using System;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    /// <summary>
    /// Defines some standard Parameter Conversions.
    /// Register by calling ParameterConversionRegistration.AddParameterConversion(ParameterConversions.NullableConversion);
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
                        Expression.Convert(GetConversion(input, innerType), type)),
                    input);
            return result;
        }

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
            var input = Expression.Parameter(typeof(object));
            return
                Expression.Lambda(
                    Expression.Condition(
                        Expression.TypeIs(input, typeof(ExcelMissing)),
                        Expression.Constant(defaultValue, type),
                        GetConversion(input, type)),
                    input);
        }

        static Expression GetConversion(Expression input, Type type)
        {
            if (type == typeof(Object))
                return input;
            if (type == typeof(Double))
                return Expression.Call(((Func<Object, Double>)ConvertToDouble).Method, input);
            if (type == typeof(String))
                return Expression.Call(((Func<Object, String>)ConvertToString).Method, input);
            if (type == typeof(DateTime))
                return Expression.Call(((Func<Object, DateTime>)ConvertToDateTime).Method, input);
            if (type == typeof(Boolean))
                return Expression.Call(((Func<Object, Boolean>)ConvertToBoolean).Method, input);
            if (type == typeof(Int64))
                return Expression.Call(((Func<Object, Int64>)ConvertToInt64).Method, input);
            if (type == typeof(Int32))
                return Expression.Call(((Func<Object, Int32>)ConvertToInt32).Method, input);
            if (type == typeof(Int16))
                return Expression.Call(((Func<Object, Int16>)ConvertToInt16).Method, input);
            if (type == typeof(UInt16))
                return Expression.Call(((Func<Object, UInt16>)ConvertToUInt16).Method, input);
            if (type == typeof(Decimal))
                return Expression.Call(((Func<Object, Decimal>)ConvertToDecimal).Method, input);

            // Fallback - not likely to be useful
            return Expression.Convert(input, type);
        }

        static double ConvertToDouble(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (double)result;
            }

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Int32.");
        }

        static string ConvertToString(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeString);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
            {
                return (string)result;
            }

            // Not sure how this can happen...
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to String.");
        }

        static DateTime ConvertToDateTime(object value)
        {
            try
            {
                return DateTime.FromOADate(ConvertToDouble(value));
            }
            catch
            {
                // Might exceed range of DateTime
                throw new InvalidCastException("Value " + value.ToString() + " could not be converted to DateTime.");
            }
        }

        static bool ConvertToBoolean(object value)
        {
            object result;
            var retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeBoolean);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return (bool)result;

            // failed - as a fallback, try to convert to a double
            retVal = XlCall.TryExcel(XlCall.xlCoerce, out result, value, (int)XlType.XlTypeNumber);
            if (retVal == XlCall.XlReturn.XlReturnSuccess)
                return ((double)result != 0.0);

            // We give up.
            throw new InvalidCastException("Value " + value.ToString() + " could not be converted to Boolean.");
        }

        static int ConvertToInt32(object value)
        {
            return checked((int)ConvertToInt64(value));
        }

        static short ConvertToInt16(object value)
        {
            return checked((short)ConvertToInt64(value));
        }

        static ushort ConvertToUInt16(object value)
        {
            return checked((ushort)ConvertToInt64(value));
        }

        static decimal ConvertToDecimal(object value)
        {
            return checked((decimal)ConvertToDouble(value));
        }

        static long ConvertToInt64(object value)
        {
            return checked((long)Math.Round(ConvertToDouble(value), MidpointRounding.ToEven));
        }

        static object GetDefault(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }
    }

    internal enum XlType : int
    {
        XlTypeNumber = 0x0001,
        XlTypeString = 0x0002,
        XlTypeBoolean = 0x0004,
        XlTypeReference = 0x0008,
        XlTypeError = 0x0010,
        XlTypeFlow = 0x0020, // Unused
        XlTypeArray = 0x0040,
        XlTypeMissing = 0x0080,
        XlTypeEmpty = 0x0100,
        XlTypeInt = 0x0800,     // int16 in XlOper, int32 in XlOper12, never passed into UDF
    }
}
