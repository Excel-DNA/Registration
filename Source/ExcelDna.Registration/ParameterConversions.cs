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
        // These can be used directly in .AddParameterConversion

        /// <summary>
        /// Legacy method: this returns a converter for Nullable[T] where T is one of the basic types that do not require any converter.
        /// If you need a Nullable[T] converter that can call into another for T, then use ParameterConversionConfiguration.AddNullableConversion.
        /// </summary>
        /// <param name="treatEmptyAsMissing"></param>
        /// <param name="treatNAErrorAsMissing"></param>
        /// <returns></returns>
        public static Func<Type, ExcelParameterRegistration, LambdaExpression> GetNullableConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return (type, paramReg) => ParameterConversionConfiguration.NullableConversion(null, type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        public static Func<Type, ExcelParameterRegistration, LambdaExpression> GetOptionalConversion(bool treatEmptyAsMissing = false, bool treatNAErrorAsMissing = false)
        {
            return (type, paramReg) => OptionalConversion(type, paramReg, treatEmptyAsMissing, treatNAErrorAsMissing);
        }

        public static Func<Type, ExcelParameterRegistration, LambdaExpression> GetEnumStringConversion()
        {
            return (type, paramReg) => EnumStringConversion(type, paramReg);
        }

        internal static object EnumParse(Type enumType, object obj)
        {
            object result;
            string objToString = obj.ToString().Trim();
            try
            {
                result = Enum.Parse(enumType, objToString, true);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException($"'{objToString}' is not a value of enum '{enumType.Name}'. Legal values are: {string.Join(", ", enumType.GetEnumNames())}");
            }
            return result;
        }

        static LambdaExpression EnumStringConversion(Type type, ExcelParameterRegistration paramReg)
        {
            // Decide whether to return a conversion function for this parameter
            if (!type.IsEnum)
                return null;

            var input = Expression.Parameter(typeof(object), "input");
            var enumTypeParam = Expression.Parameter(typeof(Type), "enumType");
            Expression<Func<Type, object, object>> enumParse = (t, s) => EnumParse(t, s);
            var result =
                Expression.Lambda(
                    Expression.Convert(
                        Expression.Invoke(enumParse, Expression.Constant(type), input),
                        type),
                    input);
            return result;
        }

        static LambdaExpression OptionalConversion(Type type, ExcelParameterRegistration paramReg, bool treatEmptyAsMissing, bool treatNAErrorAsMissing)
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
                        ParameterConversionConfiguration.MissingTest(input, treatEmptyAsMissing, treatNAErrorAsMissing),
                        Expression.Constant(defaultValue, type),
                        TypeConversion.GetConversion(input, type)),
                    input);
        }
    }
}
