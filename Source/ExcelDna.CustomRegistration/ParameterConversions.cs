using System;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    /// <summary>
    /// Defines some standard Parameter Conversions.
    /// Register by calling ParameterConversionConfiguration.AddParameterConversion(ParameterConversions.NullableConversion);
    /// </summary>
    public static class ParameterConversions
    {
        public static LambdaExpression NullableConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(Nullable<>))                 // E.g. type is Nullable<double>
                return null;

            var innerType = type.GetGenericArguments()[0]; // E.g. innerType is double
            // Here's the actual conversion function
            var input = Expression.Parameter(typeof(object));
            var result =
                Expression.Lambda(
                    Expression.Condition(
                // if the value is missing
                        Expression.TypeIs(input, typeof(ExcelMissing)),
                // cast null to int?
                        Expression.Constant(null, type),
                // else convert to int, and cast that to int?
                        Expression.Convert(TypeConversion.GetConversion(input, innerType), type)),
                    input);
            return result;
        }

        public static LambdaExpression OptionalConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!paramReg.CustomAttributes.OfType<OptionalAttribute>().Any())
                return null;

            var defaultAttribute = paramReg.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
            var defaultValue = defaultAttribute == null ? TypeConversion.GetDefault(type) : defaultAttribute.Value;
            // var returnType = type.GetGenericArguments()[0]; // E.g. returnType is double

            // Consume the attributes
            paramReg.CustomAttributes.RemoveAll(att => att is OptionalAttribute);
            paramReg.CustomAttributes.RemoveAll(att => att is DefaultParameterValueAttribute);

            // Here's the actual conversion function
            var input = Expression.Parameter(typeof(object));
            return
                Expression.Lambda(
                    Expression.Condition(
                        Expression.TypeIs(input, typeof(ExcelMissing)),
                        Expression.Constant(defaultValue, type),
                        TypeConversion.GetConversion(input, type)),
                    input);
        }
    }
}
