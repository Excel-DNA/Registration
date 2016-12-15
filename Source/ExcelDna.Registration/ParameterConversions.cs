using System;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelDna.Registration
{
    /// <summary>
    /// Defines some standard Parameter Conversions.
    /// Register by calling ParameterConversionConfiguration.AddParameterConversion(ParameterConversions.NullableConversion);
    /// </summary>
    public static class ParameterConversions
    {
        public static Func<Type, ExcelParameterRegistration, LambdaExpression> GetOptionalConversion(bool treatEmptyAsMissing = false)
        {
            return (type, paramReg) => OptionalConversion(type, paramReg, treatEmptyAsMissing);
        }

        public static Func<Type, ExcelParameterRegistration, LambdaExpression> GetEnumConversion()
        {
            return (type, paramReg) => EnumConversion(type, paramReg);
        }

        internal static object EnumParse(Type enumType, object obj)
        {
            object result;
            string objToString = obj.ToString().Trim();
            try
            {
                result = Enum.Parse(enumType, objToString, true);
            }
            catch (ArgumentException x)
            {
                throw new ArgumentException($"'{objToString}' is not a value of enum '{enumType.Name}'. Legal values are: {enumType.GetEnumNames()}");
            }
            return result;
        }

        static LambdaExpression EnumConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsEnum)
                return null;

            var input = Expression.Parameter(typeof(object), "input");
            var enumTypeParam = Expression.Parameter(typeof(Type), "enumType");
            Expression<Func<Type, object, object>> enumParse = (t, s) => EnumParse(t, s.ToString().Trim());
            var result =
                Expression.Lambda(
                    Expression.Convert(
                        Expression.Invoke(enumParse, Expression.Constant(type), input),
                        type),
                    input);
            return result;
        }

        static LambdaExpression OptionalConversion(Type type, ExcelParameterRegistration paramReg, bool treatEmptyAsMissing)
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
            var input = Expression.Parameter(typeof(object), "input");
            return
                Expression.Lambda(
                    Expression.Condition(
                        ParameterConversionConfiguration.MissingTest(input, treatEmptyAsMissing),
                        Expression.Constant(defaultValue, type),
                        TypeConversion.GetConversion(input, type)),
                    input);
        }
    }
}
