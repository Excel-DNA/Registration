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

            /// <summary>
            /// If true, the conversion will also convert all subtypes of its input type
            /// </summary>
            public bool HandleSubTypes { get; private set; }

            public ReturnConversion(Func<Type, ExcelReturnRegistration, LambdaExpression> conversion, Type typeFilter = null, bool handleSubTypes = false)
            {
                if (conversion == null)
                    throw new ArgumentNullException("conversion");

                Conversion = conversion;
                TypeFilter = typeFilter;
                HandleSubTypes = handleSubTypes;
            }

            internal LambdaExpression Convert(Type returnType, ExcelReturnRegistration returnRegistration)
            {
                if (TypeFilter != null && returnType != TypeFilter && (!HandleSubTypes || !returnType.IsSubclassOf(TypeFilter)))
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
        public ParameterConversionConfiguration AddReturnConversion(Func<Type, ExcelReturnRegistration, LambdaExpression> returnConversion, Type targetTypeOrNull = null, bool handleSubTypes = false)
        {
            var rc = new ReturnConversion(returnConversion, targetTypeOrNull, handleSubTypes);
            ReturnConversions.Add(rc);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom>(Func<Type, ExcelReturnRegistration, LambdaExpression> returnConversion, Type targetTypeOrNull = null, bool handleSubTypes = false)
        {
            AddReturnConversion(returnConversion, typeof(TFrom), handleSubTypes);
            return this;
        }

        public ParameterConversionConfiguration AddReturnConversion<TFrom, TTo>(Expression<Func<TFrom, TTo>> convert, bool handleSubTypes = false)
        {
            AddReturnConversion<TFrom>((unusedReturnType, unusedAttributes) => convert, null, handleSubTypes);
            return this;
        }
        #endregion

        internal static LambdaExpression NullableConversion(IEnumerable<ParameterConversion> parameterConversions, Type type, ExcelParameterRegistration paramReg, bool treatEmptyAsMissing,
            bool treatNAErrorAsMissing)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(Nullable<>)) // E.g. type is Nullable<Complex>
                return null;

            var innerType = type.GetGenericArguments()[0]; // E.g. innerType is Complex
            // Try to find a converter for innerType in the config
            ParameterConversion innerTypeParameterConversion = null;
            if (parameterConversions != null)
                innerTypeParameterConversion =
                    parameterConversions.FirstOrDefault(c => c.Convert(innerType, paramReg) != null);
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
                        MissingTest(input, treatEmptyAsMissing, treatNAErrorAsMissing),
                        // cast null to int?
                        Expression.Constant(null, type),
                        // else convert to int, and cast that to int?
                        Expression.Convert(innerTypeConversion, type)),
                    input);
            return result;
        }

        Func<Type, ExcelParameterRegistration, LambdaExpression> GetNullableConversion(bool treatEmptyAsMissing, bool treatNAErrorAsMissing)
        {
            return (type, paramReg) => NullableConversion(ParameterConversions, type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        /// <summary>
        /// Adds a Nullable conversion that will also translate any type parameter T of Nullable[T] for which there is a conversion in the configutation.
        /// Note that the added rule is quite generic and only has access to the T conversion rules that have already been added before it, so you should
        /// call this at the very bottom of your configuration setup sequence.
        /// </summary>
        /// <param name="treatEmptyAsMissing">If true, any empty cells will be treated as null values</param>
        /// <param name="treatNAErrorAsMissing">If true, any #NA! errors will be treated as null values</param>
        /// <returns>The parameter conversion configuration with the new added rule</returns>
        public ParameterConversionConfiguration AddNullableConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return AddParameterConversion(GetNullableConversion(treatEmptyAsMissing, treatNAErrorAsMissing));
        }

        static bool MissingOrNATest(object input, bool treatEmptyAsMissing)
        {
            var inputArray = input as object[];
            if (inputArray != null && inputArray.Length == 1)
                input = inputArray[0];
            Type inputType = input.GetType();
            bool result = (inputType == typeof(ExcelMissing)) ||
                          (treatEmptyAsMissing && inputType == typeof(ExcelEmpty));
            if (!result && inputType == typeof(ExcelError))
                result = (ExcelError) input == ExcelError.ExcelErrorNA;
            return result;
        }

        internal static Expression MissingTest(ParameterExpression input, bool treatEmptyAsMissing,
            bool treatNAErrorAsMissing)
        {
            Expression r = null;
            if (treatNAErrorAsMissing)
            {
                var methodMissingOrNATest = typeof(ParameterConversionConfiguration).GetMethod("MissingOrNATest",
                    BindingFlags.NonPublic | BindingFlags.Static);
                r = Expression.Call(null, methodMissingOrNATest, input, Expression.Constant(treatEmptyAsMissing));
            }
            else
            {
                r = Expression.TypeIs(input, typeof(ExcelMissing));
                if (treatEmptyAsMissing)
                    r = Expression.OrElse(r, Expression.TypeIs(input, typeof(ExcelEmpty)));
            }
            return r;
        }
    }
}
