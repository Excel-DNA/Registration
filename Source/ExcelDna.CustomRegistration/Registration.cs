using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.CustomRegistration.Utils;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration
{
    // Ideas:
    // * Parameter and Return Type conversions, including Range, Optional etc.
    // * Object Instances - Methods and Properties (with INotifyPropertyChanged support, and Disposable from Observable handles)
    // * Struct semantics, like built-in COMPLEX data
    // * Logging and other pre / post / catch / finally handlers

    // A first attempt to allow chaining of the CustomRegistration rewrites.
    public static class Registration
    {
        /// <summary>
        ///  
        /// </summary>
        /// <returns>All public static methods in registered assemblies that are decorated with an [ExcelFunction] attribute 
        /// (or a derived attribute, like [ExcelAsyncFunction]
        /// </returns>
        public static IEnumerable<RegistrationEntry> GetExcelFunctions()
        {
            return from ass in ExcelIntegration.GetExportedAssemblies()
                   from typ in ass.GetTypes()
                   from mi in typ.GetMethods(BindingFlags.Public | BindingFlags.Static)
                   where mi.GetCustomAttribute<ExcelFunctionAttribute>() != null
                   select new RegistrationEntry(mi);
        }

        /// <summary>
        /// Registers the given of functions with Excel-DNA.
        /// </summary>
        /// <param name="registrationEntries"></param>
        public static void RegisterFunctions(this IEnumerable<RegistrationEntry> registrationEntries)
        {
            var delList = new List<Delegate>();
            var attList = new List<object>();
            var argAttList = new List<List<object>>();
            foreach (var entry in registrationEntries)
            {
                try
                {
                    var del = entry.FunctionLambda.Compile();
                    var att = entry.FunctionAttribute;
                    var argAtt = entry.ArgumentAttributes == null
                                       ? null
                                       : entry.ArgumentAttributes.Cast<object>().ToList();

                    delList.Add(del);
                    attList.Add(att);
                    argAttList.Add(argAtt);
                }
                catch (Exception ex)
                {
                    Logging.LogDisplay.WriteLine("Exception while registering method {0} - {1}", entry.FunctionAttribute.Name, ex.ToString());
                }
            }

            ExcelIntegration.RegisterDelegates(delList, attList, argAttList);
        }
    }
}
