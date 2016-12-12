using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelDna.Registration
{
    // Used for parameter and return type conversions (when these can be done without interfering with the rest of the function).

    // CONSIDER: Add a name to the XXXConversion for tracing and debugging
    // CONSIDER: Do we need to consider Co-/Contravariance and allow processing of sub-/super-types?
    // What about native async function, they return 'void' type?

    public class ParameterConversionConfiguration
    {
        internal class ParameterConversion
        {
            // Conversion receives the parameter type and parameter registration info, 
            // and should return an Expression<Func<TTo, TFrom>> 
            // (and may optionally update the information in the ExcelParameterRegistration.
            // May return null to indicate that no conversion should be applied.
            public Func<Type, ExcelParameterRegistration, LambdaExpression> Conversion { get; private set; }

            // The TypeFilter is used as a quick filter to decide whether the Conversion function should be called for a parameter.
            // TypeFilter may be null to indicate that conversion should be applied for all types.
            // (The Conversion function may anyway return null to indicate that no conversion should be applied.)
            public Type TypeFilter { get; private set; }

            public ParameterConversion(Func<Type, ExcelParameterRegistration, LambdaExpression> conversion, Type typeFilter = null)
            {
                if (conversion == null)
                    throw new ArgumentNullException("conversion");

                Conversion = conversion;
                TypeFilter = typeFilter;
            }

            internal LambdaExpression Convert(Type paramType, ExcelParameterRegistration paramReg)
            {
                if (TypeFilter != null && paramType != TypeFilter)
                    return null;

                return Conversion(paramType, paramReg);
            }
        }

        internal class ReturnConversion
        {
            // Conversion receives the return type and list of custom attributes applied to the return value,
            // and should return  an Expression<Func<TTo, TFrom>> 
            // (and may optionally update the information in the ExcelParameterRegistration.
            // May return null to indicate that no conversion should be applied.
            public Func<Type, ExcelReturnRegistration, LambdaExpression> Conversion { get; private set; }

            // TypeFilter is used as a quick filter to decide whether the conversion function should be called for a return value.
            // TypeFilter be null to indicate that conversion should be applied for all types
            // The Conversion function may anyway return null to indicate that no conversion should be applied.
            public Type TypeFilter { get; private set; }

            public ReturnConversion(Func<Type, ExcelReturnRegistration, LambdaExpression> conversion, Type typeFilter = null)
            {
                if (conversion == null)
                    throw new ArgumentNullException("conversion");

                Conversion = conversion;
                TypeFilter = typeFilter;
            }
        
            internal LambdaExpression Convert(Type returnType, ExcelReturnRegistration returnRegistration)
            {
                if (TypeFilter != null && returnType != TypeFilter && !returnType.IsSubclassOf(TypeFilter))
                    return null;

 	            LambdaExpression result = Conversion(returnType, returnRegistration);

                if (TypeFilter != null && returnType != TypeFilter)
                {
                    var returnValue = Expression.Parameter(returnType, "returnValue");
                    var castExpr = Expression.Convert(returnValue, TypeFilter);
                    var composeExpr = Expression.Invoke(result, castExpr);
                    result = Expression.Lambda(composeExpr, returnValue);
                }
                return result;
            }
        }

        internal List<ParameterConversion> ParameterConversions { get; private set; }
        internal List<ReturnConversion>    ReturnConversions    { get; private set; }

        public ParameterConversionConfiguration()
        {
            ParameterConversions = new List<ParameterConversion>();
            ReturnConversions    = new List<ReturnConversion>();
        }

        #region Various overloads for adding conversions

        // Most general case - called by the overloads below
        /// <summary>
        /// Converts a parameter from an Excel-friendly type (e.g. object, or string) to an add-in friendly type, e.g. double? or InternalType.
        /// Will only be considered for those parameters that have a 'to' type that matches targetTypeOrNull,
        ///  or for all types if null is passes for the first parameter.
        /// </summary>
        /// <param name="parameterConversion"></param>
        /// <param name="targetTypeOrNull"></param>
        public ParameterConversionConfiguration AddParameterConversion(Func<Type, ExcelParameterRegistration, LambdaExpression> parameterConversion, Type targetTypeOrNull = null)
        {
            var pc = new ParameterConversion(parameterConversion, targetTypeOrNull);
            ParameterConversions.Add(pc);
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TTo>(Func<Type, ExcelParameterRegistration, LambdaExpression> parameterConversion)
        {
            AddParameterConversion(parameterConversion, typeof(TTo));
            return this;
        }

        public ParameterConversionConfiguration AddParameterConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddParameterConversion<TTo>((unusedParamType, unusedParamReg) => convert);
            return this;
        }

        // Most general case - called by the overloads below
        public ParameterConversionConfiguration AddReturnConversion(Func<Type, ExcelReturnRegistration, LambdaExpression> returnConversion, Type targetTypeOrNull = null)
        {
            var rc = new ReturnConversion(returnConversion, targetTypeOrNull);
            ReturnConversions.Add(rc);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom>(Func<Type, ExcelReturnRegistration, LambdaExpression> returnConversion, Type targetTypeOrNull = null)
        {
            AddReturnConversion(returnConversion, typeof(TFrom));
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert)
        {
            AddReturnConversion<TFrom>((unusedReturnType, unusedAttributes) => convert);
            return this;
        }
        #endregion

        LambdaExpression NullableConversion(Type type, ExcelParameterRegistration paramReg, bool treatEmptyAsMissing)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(Nullable<>)) // E.g. type is Nullable<Complex>
                return null;

            var innerType = type.GetGenericArguments()[0]; // E.g. innerType is Complex
            ParameterConversion innerTypeParameterConversion = ParameterConversions.First(c => c.Convert(innerType, paramReg)!=null); // Try to find a converter for innerType in the config
            ParameterExpression input = null;
            Expression innerTypeConversion = null;
            // if we have a converter for innertype in the config, then use it. Otherwise try one of the conversions for the basic types
            if (innerTypeParameterConversion == null)
            {
                input = Expression.Parameter(typeof(object), "input");
                innerTypeConversion = TypeConversion.GetConversion(input, innerType);
            }
            else
            {
                var innerTypeParamConverter = innerTypeParameterConversion.Convert(innerType, paramReg);
                input = Expression.Parameter(innerTypeParamConverter.Parameters[0].Type, "input");
                innerTypeConversion = Expression.Invoke(innerTypeParamConverter, input);
            }
            // Here's the actual conversion function
            var result =
                Expression.Lambda(
                    Expression.Condition(
                        // if the value is missing (or possibly empty)
                        MissingTest(input, treatEmptyAsMissing),
                        // cast null to int?
                        Expression.Constant(null, type),
                        // else convert to int, and cast that to int?
                        Expression.Convert(innerTypeConversion, type)),
                    input);
            return result;
        }

        Func<Type, ExcelParameterRegistration, LambdaExpression> GetNullableConversion(bool treatEmptyAsMissing)
        {
            return (type, paramReg) => NullableConversion(type, paramReg, treatEmptyAsMissing);
        }

        public ParameterConversionConfiguration AddNullableConversion(bool treatEmptyAsMissing = false)
        {
            return AddParameterConversion(GetNullableConversion(treatEmptyAsMissing));
        }

        static bool MissingOrNATest(object input)
        {
            var inputArray = input as object[];
            if (inputArray != null && inputArray.Length == 1)
                input = inputArray[0];
            Type inputType = input.GetType();
            bool result = (inputType == typeof(ExcelMissing)) || (inputType == typeof(ExcelEmpty));
            if (!result && inputType == typeof(ExcelError))
                result = (ExcelError)input == ExcelError.ExcelErrorNA;
            return result;
        }

        internal static Expression MissingTest(ParameterExpression input, bool treatEmptyAsMissing)
        {
            Expression r = null;
            if (!treatEmptyAsMissing)
                r = Expression.TypeIs(input, typeof(ExcelMissing));
            else
            {
                var methodMissingorNATest = typeof(ParameterConversionConfiguration).GetMethod("MissingOrNATest", BindingFlags.NonPublic| BindingFlags.Static);
                r = Expression.Call(null, methodMissingorNATest, input);
            }
            return r;
        }
    }
}
