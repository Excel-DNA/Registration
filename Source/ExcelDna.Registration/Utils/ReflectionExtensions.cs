using System;
using System.Linq;
using System.Reflection;

namespace ExcelDna.CustomRegistration.Utils
{
    static class ReflectionExtensions
    {
        public static T GetCustomAttribute<T>(this MethodInfo mi, bool inherit = false) where T : Attribute
        {
            return mi.GetCustomAttributes(typeof(T), inherit).Cast<T>().FirstOrDefault();
        }

        public static T GetCustomAttribute<T>(this ParameterInfo mi, bool inherit = false) where T : Attribute
        {
            return mi.GetCustomAttributes(typeof(T), inherit).Cast<T>().FirstOrDefault();
        }
    }
}
