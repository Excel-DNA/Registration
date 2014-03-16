using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    /// <summary>
    /// Defines some standard Parameter Conversions.
    /// Register by calling ParameterConversionRegistration.AddParameterConversion(ParameterConversions.NullableConversion);
    /// </summary>
    public static class ParameterConversions
    {
        // TODO: Need to take the non-null case of the conversion more seriously.
        public static LambdaExpression NullableConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsGenericType || type.GetGenericTypeDefinition() != typeof(Nullable<>))                 // E.g. type is Nullable<double>
                return null;

            // var returnType = type.GetGenericArguments()[0]; // E.g. returnType is double

            // Here's the actual conversion function
            // TODO: Sort out type conversion from object to type...?
            // E.g. Fails for optional int parameters (because doubles are passed in)
            var input = Expression.Parameter(typeof(object));
            var result =
                Expression.Lambda(
                    Expression.Condition(
                        Expression.TypeIs(input, typeof(ExcelMissing)),
                        Expression.Constant(null, type),
                        Expression.Convert(input, type)),
                    input);
            return result;
        }

        // TODO: Need to take the non-null case of the conversion more seriously.
        public static LambdaExpression OptionalConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!paramReg.CustomAttributes.OfType<OptionalAttribute>().Any())
                return null;

            var defaultAttribute = paramReg.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
            var defaultValue = defaultAttribute == null ? GetDefault(type) : defaultAttribute.Value;
            // var returnType = type.GetGenericArguments()[0]; // E.g. returnType is double

            // Consume the attributes
            paramReg.CustomAttributes.RemoveAll(att => att is OptionalAttribute);
            paramReg.CustomAttributes.RemoveAll(att => att is DefaultParameterValueAttribute);

            // Here's the actual conversion function
            // TODO: Sort out type conversion from object to type...?
            // E.g. Fails for nullable int parameters
            var input = Expression.Parameter(typeof(object));
            return
                Expression.Lambda(
                    Expression.Condition(
                        Expression.TypeIs(input, typeof(ExcelMissing)),
                        Expression.Constant(defaultValue, type),
                        Expression.Convert(input, type)),
                    input);
        }

        // TODO: Simple type conversion.....

        static object GetDefault(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }
    }
}
