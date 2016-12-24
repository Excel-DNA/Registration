using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.Registration
{
    /// <summary>
    /// A marshal-by-reference object cache must implement this interface.
    /// See class ExcelObjectCache for a readily available implementation which you will most likely want to reuse.
    /// This interface is used to store any results and look up any parameters of type T, where T was either decorated
    /// with ExcelMarshalByRefAttribute or added to ParameterConversionConfiguration.MarshalByRef property of the configuration
    /// being used.
    /// </summary>
    public interface IReferenceMarshaller
    {
        /// <summary>
        /// This method is called to look up an object id string in the cache.
        /// </summary>
        /// <param name="id">The object id string coming from Excel</param>
        /// <returns></returns>
        object Lookup(string id);

        /// <summary>
        /// This method is called to store any objects returned by a function.
        /// </summary>
        /// <param name="o">The object</param>
        /// <returns>The string id to display in the cell</returns>
        string Store(object o);

        /// <summary>
        /// This method is called to tell the cache that we are starting to evaluate a new cell.
        /// </summary>
        void SetCurrentCell();
    }

    // CONSIDER: Can one use an ExpressionVisitor to do these things....?
    public static class ParameterConversionRegistration
    {
        public static IEnumerable<ExcelFunctionRegistration> ProcessParameterConversions(this IEnumerable<ExcelFunctionRegistration> registrations, ParameterConversionConfiguration conversionConfig)
        {
            foreach (var reg in registrations)
            {
                // Keep a list of conversions for each parameter
                // TODO: Prevent having a cycle, but allow arbitrary ordering...?

                var paramsConversions = new List<List<LambdaExpression>>();
                for (int i = 0; i < reg.FunctionLambda.Parameters.Count; i++)
                {
                    var initialParamType = reg.FunctionLambda.Parameters[i].Type;
                    var paramReg = reg.ParameterRegistrations[i];

                    var paramConversions = GetParameterConversions(conversionConfig, initialParamType, paramReg);
                    paramsConversions.Add(paramConversions);
                } // for each parameter !

                // Process return conversions
                var returnConversions = GetReturnConversions(conversionConfig, reg.FunctionLambda.ReturnType, reg.ReturnRegistration);

                // Now we apply all the conversions
                ApplyConversions(reg, paramsConversions, returnConversions);

                yield return reg;
            }
        }

        private static T LookupFromCache<T>(IReferenceMarshaller objectCache, string idString) where T : class
        {
            object o = objectCache.Lookup(idString);
            T r = o as T;
            if (r == null)
                throw new ArgumentException($"Object '{idString}' is not of type {typeof(T)}, it is of type {o.GetType()}.");
            return r;
        }

        static LambdaExpression ComposeLambdas(IEnumerable<LambdaExpression> lambdas)
        {
            LambdaExpression result = null;
            if (lambdas != null)
            {
                var convsIter = lambdas.GetEnumerator();
                if (convsIter.MoveNext())
                {
                    result = convsIter.Current;
                    while (convsIter.MoveNext())
                    {
                        result = Expression.Lambda(Expression.Invoke(result, convsIter.Current),
                            convsIter.Current.Parameters);
                    }
                }
            }
            return result;
        }

        internal static LambdaExpression GetParameterConversion(ParameterConversionConfiguration conversionConfig,
            Type initialParamType, ExcelParameterRegistration paramRegistration)
        {
            return ComposeLambdas(GetParameterConversions(conversionConfig, initialParamType, paramRegistration));
        }

        // Should return null if there are no conversions to apply
        internal static List<LambdaExpression> GetParameterConversions(ParameterConversionConfiguration conversionConfig, Type initialParamType, ExcelParameterRegistration paramRegistration)
        {
            var appliedConversions = new List<LambdaExpression>();

            bool marshalByRef = conversionConfig.MarshalByRef.Contains(initialParamType);
            marshalByRef = marshalByRef || initialParamType.GetCustomAttributes(typeof(ExcelMarshalByRefAttribute), true).Length > 0;

            if (marshalByRef && conversionConfig.ReferenceMarshaller != null)
            {
                var idString = Expression.Parameter(typeof(string), "idString");
                var lookupFromCacheLambda = Expression.Lambda(
                    Expression.Call(
                        typeof(ParameterConversionRegistration), "LookupFromCache", new Type[] {initialParamType},
                        Expression.Constant(conversionConfig.ReferenceMarshaller), idString),
                    idString);
                appliedConversions.Add(lookupFromCacheLambda);
            }
            else
            {
                // paramReg might be modified internally by the conversions, but won't become a different object
                var paramType = initialParamType; // Might become a different type as we convert
                foreach (var paramConversion in conversionConfig.ParameterConversions)
                {
                    var lambda = paramConversion.Convert(paramType, paramRegistration);
                    if (lambda == null)
                        continue;

                    // We got one to apply...
                    // Some sanity checks
                    Debug.Assert(lambda.Parameters.Count == 1);
                    Debug.Assert(lambda.ReturnType == paramType || lambda.ReturnType.IsEquivalentTo(paramType));

                    appliedConversions.Add(lambda);

                    // Change the Parameter Type to be whatever the conversion function takes us to
                    // for the next round of processing
                    paramType = lambda.Parameters[0].Type;
                }
            }

            if (appliedConversions.Count == 0)
                return null;

            return appliedConversions;
        }

        private delegate string StoreToCacheDelegate(object o);

        internal static LambdaExpression GetReturnConversion(ParameterConversionConfiguration conversionConfig,
            Type initialReturnType, ExcelReturnRegistration returnRegistration, bool setCurrentCellInMarshalByRefCache = true)
        {
            return ComposeLambdas(GetReturnConversions(conversionConfig, initialReturnType, returnRegistration, setCurrentCellInMarshalByRefCache));
        }

        internal static List<LambdaExpression> GetReturnConversions(ParameterConversionConfiguration conversionConfig, Type initialReturnType, ExcelReturnRegistration returnRegistration,
            bool setCurrentCellInMarshalByRefCache = true)
        {
            var appliedConversions = new List<LambdaExpression>();

            bool marshalByRef = conversionConfig.IsMarshalByRef(initialReturnType);
            if (marshalByRef)
            {
                var input = Expression.Parameter(initialReturnType, "input");

                StoreToCacheDelegate storeToCacheDelegate = null;
                if (setCurrentCellInMarshalByRefCache)
                    storeToCacheDelegate = (o) =>
                    {
                        conversionConfig.ReferenceMarshaller.SetCurrentCell();
                        return conversionConfig.ReferenceMarshaller.Store((object) o);
                    };
                else storeToCacheDelegate = (o) =>
                    {
                        return conversionConfig.ReferenceMarshaller.Store((object) o);
                    };

                var storeToCacheLambda = Expression.Lambda(Expression.Call(Expression.Constant(storeToCacheDelegate.Target), storeToCacheDelegate.Method, input), input);
                appliedConversions.Add(storeToCacheLambda);
            }
            else
            {
                // paramReg might be modified internally by the conversions, but won't become a different object
                var returnType = initialReturnType; // Might become a different type as we convert

                foreach (var returnConversion in conversionConfig.ReturnConversions)
                {
                    var lambda = returnConversion.Convert(returnType, returnRegistration);
                    if (lambda == null)
                        continue;

                    // We got one to apply...
                    // Some sanity checks
                    Debug.Assert(lambda.Parameters.Count == 1);
                    Debug.Assert(lambda.Parameters[0].Type == returnType);

                    appliedConversions.Add(lambda);

                    // Change the Return Type to be whatever the conversion function returns
                    // for the next round of processing
                    returnType = lambda.ReturnType;
                }
            }

            if (appliedConversions.Count == 0)
                return null;

            return appliedConversions;
        }

        // returnsConversion and the entries in paramsConversions may be null.
        static void ApplyConversions(ExcelFunctionRegistration reg, List<List<LambdaExpression>> paramsConversions, List<LambdaExpression> returnConversions)
        {
            // CAREFUL: The parameter transformations are applied in reverse order to how they're identified.
            // We do the following transformation
            //      public static string dnaParameterConvertTest(double? optTest) {   };
            //
            // with conversions convert1 and convert2 taking us from Type1 to double?
            // 
            // to
            //      public static string dnaParameterConvertTest(Type1 optTest) 
            //      {   
            //          return convertRet2(convertRet1(
            //                      dnaParameterConvertTest(
            //                          paramConvert1(optTest)
            //                            )));
            //      };
            // 
            // and then with a conversion from object to Type1, resulting in
            //
            //      public static string dnaParameterConvertTest(object optTest) 
            //      {   
            //          return convertRet2(convertRet1(
            //                      dnaParameterConvertTest(
            //                          paramConvert1(paramConvert2(optTest))
            //                            )));
            //      };

            Debug.Assert(reg.FunctionLambda.Parameters.Count == paramsConversions.Count);

            // NOTE: To cater for the Range COM type equivalance, we need to distinguish the FunctionLambda's parameter type and the paramConversion ReturnType.
            //       These need not be the same, but the should at least be equivalent.

            // build up the invoke expression for each parameter
            var wrappingParameters = reg.FunctionLambda.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();

            // Build the nested parameter convertion expression.
            // Start with the wrapping parameters as they are. Then replace with the nesting of conversions as needed.
            var paramExprs = new List<Expression>(wrappingParameters);
            for (int i = 0; i < paramsConversions.Count; i++)
            {
                var paramConversions = paramsConversions[i];
                if (paramConversions == null)
                    continue;

                // If we have a list, there should be at least one conversion in it.
                Debug.Assert(paramConversions.Count > 0);
                // Update the calling parameter type to be the outer one in the conversion chain.
                wrappingParameters[i] = Expr.Parameter(paramConversions.Last().Parameters[0].Type, wrappingParameters[i].Name);
                // Start with just the (now updated) outer param which will be the inner-most value in the conversion chain
                Expression wrappedExpr = wrappingParameters[i];
                // Need to go in reverse for the parameter wrapping
                // Need to now build from the inside out
                foreach (var conversion in Enumerable.Reverse(paramConversions))
                {
                    wrappedExpr = Expr.Invoke(conversion, wrappedExpr);
                }
                paramExprs[i] = wrappedExpr;
            }

            var wrappingCall = Expr.Invoke(reg.FunctionLambda, paramExprs);
            if (returnConversions != null)
            {
                foreach (var conversion in returnConversions)
                    wrappingCall = Expr.Invoke(conversion, wrappingCall);
            }

            reg.FunctionLambda = Expr.Lambda(wrappingCall, reg.FunctionLambda.Name, wrappingParameters);
        }
    }
}
