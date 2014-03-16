using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.CustomRegistration
{
    // CONSIDER: Can one use an ExpressionVisitor to do these things....?

    // Some concerns: What about native async function, they return 'void' type?
    // (Might interfere with our abuse of void in the Dictionary)

    // TODO: Maybe need to turn these into objects with type and name so that we can trace and debug....
    public delegate LambdaExpression ParameterConversion(Type parameterType, ExcelParameterRegistration parameterRegistration);
    public delegate LambdaExpression ReturnConversion(Type returnType, List<object> returnCustomAttributes);

    public static class ParameterConversionRegistration
    {
        // UGLY: We use Void as a special value to indicate the conversions to be processed for all types
        //       I try to hide that as an implementation, to the external functions use null to indicate the universal case.
        readonly static Dictionary<Type, List<ParameterConversion>> _parameterConversions = new Dictionary<Type, List<ParameterConversion>>();
        readonly static Dictionary<Type, List<ReturnConversion>> _returnConversions = new Dictionary<Type, List<ReturnConversion>>();

        #region Various overloads for adding conversions

        // Most general case - called by the overloads below
        /// <summary>
        /// Converts a parameter from an Excel-friendly type (e.g. object, or string) to an add-in friendly type, e.g. double? or InternalType.
        /// Will only be considered for those parameters that have a 'to' type that matches targetTypeOrNull,
        ///  or for all types if null is passes for the first parameter.
        /// </summary>
        /// <param name="targetTypeOrNull"></param>
        /// <param name="parameterConversion"></param>
        public static void AddParameterConversion(Type targetTypeOrNull, ParameterConversion parameterConversion)
        {
            var targetTypeOrVoid = targetTypeOrNull ?? typeof(void);

            List<ParameterConversion> typeConversions;
            if (_parameterConversions.TryGetValue(targetTypeOrVoid, out typeConversions))
            {
                typeConversions.Add(parameterConversion);
            }
            else
            {
                _parameterConversions[targetTypeOrVoid] = new List<ParameterConversion> { parameterConversion };
            }
        }

        public static void AddParameterConversion(ParameterConversion parameterConversion)
        {
            AddParameterConversion(null, parameterConversion);
        }

        public static void AddParameterConversion<TTo>(ParameterConversion parameterConversion)
        {
            AddParameterConversion(typeof(TTo), parameterConversion);
        }

        public static void AddParameterConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddParameterConversion<TTo>((unusedParamType, unusedParamReg) => convert);
        }

        // This is a nice signature for registering conversions, but is worse than Expression<...> when applying
        public static void AddParameterConversionFunc<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddParameterConversion<TTo>(
                (unusedParamType, unusedParamReg) => 
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
        }

        public static void AddParameterConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            // CONSIDER: We really don't want our the CustomRegistration compilation to build out a closure object here...
            AddParameterConversion<TTo>(
                (unusedParamType, paramReg) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(paramReg.CustomAttributes, value)));
        }

        // Most general case - called by the overloads below
        public static void AddReturnConversion(Type targetTypeOrNull, ReturnConversion returnConversion)
        {
            var targetTypeOrVoid = targetTypeOrNull ?? typeof(void);

            List<ReturnConversion> typeConversions;
            if (_returnConversions.TryGetValue(targetTypeOrVoid, out typeConversions))
            {
                typeConversions.Add(returnConversion);
            }
            else
            {
                _returnConversions[targetTypeOrVoid] = new List<ReturnConversion> { returnConversion };
            }
        }

        public static void AddReturnConversion<TFrom>(ReturnConversion returnConversion)
        {
            AddReturnConversion(typeof(TFrom), returnConversion);
        }

        public static void AddReturnConversion<TFrom, TTo>(Func<TFrom, TTo> convert)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, unusedAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convert(value)));
        }

        public static void AddReturnConversion<TFrom, TTo>(Func<List<object>, TFrom, TTo> convertWithAttributes)
        {
            AddReturnConversion<TFrom>(
                (unusedReturnType, returnAttributes) =>
                    (Expression<Func<TFrom, TTo>>)(value => convertWithAttributes(returnAttributes, value)));
        }
        #endregion

        public static IEnumerable<ExcelFunctionRegistration> ProcessParameterConversions(this IEnumerable<ExcelFunctionRegistration> registrations)
        {
            // Make sure that there are 'global' type and return value conversion ists registered - even though they might be empty.
            if (!_parameterConversions.ContainsKey(typeof(void)))
                _parameterConversions.Add(typeof(void), new List<ParameterConversion>());

            if (!_returnConversions.ContainsKey(typeof(void)))
                _returnConversions.Add(typeof(void), new List<ReturnConversion>());

            foreach (var reg in registrations)
            {
                // Keep a list of conversions for each parameter
                // TODO: Prevent having a cycle, but allow arbitrary ordering...?

                var paramsConversions = new List<List<LambdaExpression>>();
                for (int i = 0; i < reg.FunctionLambda.Parameters.Count; i++)
                {
                    var initialParamType = reg.FunctionLambda.Parameters[i].Type;
                    var paramReg = reg.ParameterRegistrations[i];

                    // NOTE: We add null for cases where no conversions apply.
                    var paramConversions = GetParameterConversions(initialParamType, paramReg);
                    paramsConversions.Add(paramConversions);
                } // for each parameter

                // Process return conversions
                var returnConversions = GetReturnConversions(reg.FunctionLambda.ReturnType, reg.ReturnCustomAttributes);

                // Now we apply all the conversions
                ApplyConversions(reg, paramsConversions, returnConversions);

                yield return reg;
            }
        }

        // Should return null if there are no conversions to apply
        static List<LambdaExpression> GetParameterConversions(Type initialParamType, ExcelParameterRegistration paramReg)
        {
            // paramReg Might be modified internally, should not become a different object
            var paramType = initialParamType; // Might become a different type as we convert

            // Assume most parameters will need no conversion
            List<LambdaExpression> paramConversions = null;

            // Get hold of the global conversions list (which we assume is always present)
            var globalParameterConversions = _parameterConversions[typeof(void)];

            // Try to repeatedly apply conversions until none are applicable.
            // We add a simple guard to covers for cycles and ill-behaved conversions functions
            // TODO: Improve tracing and log better error
            const int maxConversionDepth = 16;
            var depth = 0;
            while (depth < maxConversionDepth)
            {
                // First check specific type conversions, 
                // then also the global type conversions (that are not restricted to a specific type)
                List<ParameterConversion> typeConversions;
                if (_parameterConversions.TryGetValue(paramType, out typeConversions))
                    typeConversions.AddRange(globalParameterConversions);
                else
                    typeConversions = globalParameterConversions;

                var applied = false;
                // we have conversions that might be applied to this type...
                // see if we can find one to be applied
                // Note that convert might also make modifications to the paramReg object...
                foreach (var convert in typeConversions)
                {
                    var lambda = convert(paramType, paramReg);
                    if (lambda == null)
                        continue; // Try next conversion for this type

                    // We got one to apply...
                    // Some sanity checks
                    Debug.Assert(lambda.Parameters.Count == 1);
                    Debug.Assert(lambda.ReturnType == paramType);

                    // Check if we need to make a new conversion list
                    if (paramConversions == null)
                        paramConversions = new List<LambdaExpression>();

                    paramConversions.Add(lambda);
                    // Change the Parameter Type to be whatever the conversion function takes us to
                    // for the next round of processing
                    paramType = lambda.Parameters[0].Type;
                    applied = true;
                    break;
                }
                if (applied)
                    depth++;
                else
                    break; // None of the conversions were applied - stop trying
            } // while checking types

            return paramConversions;
        }

        static List<LambdaExpression> GetReturnConversions(Type initialReturnType, List<object> returnCustomAttributes)
        {
            // returnCustomAttributes list might be modified, should not become a different object
            var returnType = initialReturnType; // Might become a different type as we convert

            // Assume most returns will need no conversion
            List<LambdaExpression> returnConversions = null;

            // Get hold of the global conversions list (which we assume is always present)
            var globalReturnConversions = _returnConversions[typeof(void)];

            // Try to repeatedly apply conversions until none are applicable.
            // We add a simple guard to covers for cycles and ill-behaved conversions functions
            // TODO: Improve tracing and log better error
            const int maxConversionDepth = 16;
            var depth = 0;
            while (depth < maxConversionDepth)
            {
                // First check specific type conversions, 
                // then also the global type conversions (that are not restricted to a specific type)
                List<ReturnConversion> typeConversions;
                if (_returnConversions.TryGetValue(returnType, out typeConversions))
                    typeConversions.AddRange(globalReturnConversions);
                else
                    typeConversions = globalReturnConversions;

                var applied = false;
                // we have conversions that might be applied to this type...
                // see if we can find one to be applied
                // Note that convert might also make modifications to the return attributes list...
                foreach (var convert in typeConversions)
                {
                    var lambda = convert(returnType, returnCustomAttributes);
                    if (lambda == null)
                        continue; // Try next conversion for this type

                    // We got one to apply...
                    // Some sanity checks
                    Debug.Assert(lambda.Parameters.Count == 1);
                    Debug.Assert(lambda.Parameters[0].Type == returnType);

                    // Check if we need to make a new conversion list
                    if (returnConversions == null)
                        returnConversions = new List<LambdaExpression>();

                    returnConversions.Add(lambda);
                    // Change the Return Type to be whatever the conversion function returns
                    // for the next round of processing
                    returnType = lambda.ReturnType;
                    applied = true;
                    break;
                }
                if (applied)
                    depth++;
                else
                    break; // None of the conversions were applied - stop trying
            } // while checking types

            return returnConversions;
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

            // NOTE NOTE - Don't need the Variables....

            // build up the invoke expression for each parameter
            var wrappingParameters = new List<ParameterExpression>(reg.FunctionLambda.Parameters);
            var paramExprs = reg.FunctionLambda.Parameters.Select((param, i) =>
                {
                    var paramConversions = paramsConversions[i];

                    // Starting point is just the parameter expression
                    Expression wrappedExpr = param;
                    if (paramConversions != null)
                    {
                        // Need to go in reverse for the parameter wrapping
                        // Need to now build from the inside out
                        wrappingParameters[i] = Expr.Parameter(paramConversions.Last().Parameters[0].Type, param.Name);
                        // Start with just the final outer param
                        wrappedExpr = wrappingParameters[i];
                        foreach (var conversion in Enumerable.Reverse(paramConversions))
                        {
                            wrappedExpr = Expr.Invoke(conversion, wrappedExpr);
                        }
                    }

                    return (Expression)wrappedExpr;
                });

            var wrappingCall = Expr.Invoke(reg.FunctionLambda, paramExprs);
            if (returnConversions != null)
            {
                foreach (var conversion in returnConversions)
                    wrappingCall = Expr.Invoke(conversion, wrappingCall);
            }

            reg.FunctionLambda = Expr.Lambda(wrappingCall, wrappingParameters);
        }
    }
}
