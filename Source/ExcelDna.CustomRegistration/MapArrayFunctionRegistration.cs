using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Linq.Expressions;
using System.Text;
using ExcelDna.Integration;

[assembly: InternalsVisibleTo("ExcelDna.CustomRegistration.Test")]
namespace ExcelDna.CustomRegistration
{
    /// <summary>
    /// Defines an attribute to identify functions that define a mapping array function using enumerables and property reflection.
    /// </summary>
    public class ExcelMapArrayFunctionAttribute : ExcelFunctionAttribute
    {
    }

    public static class MapArrayFunctionRegistration
    {
        /// <summary>
        /// Modifies RegistrationEntries which have methods with IEnumerable signatures,
        /// allowing them to be converted to and from Excel Ranges (i.e. object[,]).
        /// The first row in each Range contains column headers, which are mapped to and from
        /// the public properties of the enumerated types.
        /// Currently just supports methods with signature IEnumerable<typeparamref name="T"/> -> IEnumerable<typeparamref name="U"/>
        /// E.g.
        ///     struct Output { int Out; }
        ///     struct Input  { int In1; int In2; }
        ///     IEnumerable<typeparamref name="Output"/> MyFunc(IEnumerable<typeparamref name="Input"/>) { ... }
        /// In Excel, use an Array Formula, e.g.
        ///       | A       B       C       
        ///     --+-------------------------
        ///     1 | In1     In2     {=MyFunc(A1:B3)} -> Out
        ///     2 | 1.0     2.0     {=MyFunc(A1:B3)} -> 1.5
        ///     3 | 2.0     3.0     {=MyFunc(A1:B3)} -> 2.5
        /// </summary>
        public static IEnumerable<ExcelFunctionRegistration> ProcessMapArrayFunctions(
            this IEnumerable<ExcelFunctionRegistration> registrations)
        {
            foreach (var reg in registrations)
            {
                if (!(reg.FunctionAttribute is ExcelMapArrayFunctionAttribute))
                {
                    // Not considered at all
                    yield return reg;
                    continue;
                }

                // avoid manipulating reg until we're sure it can succeed
                try
                {
                    var shim = reg.FunctionLambda.MakeObjectArrayShim();

                    // replace the Function attribute, with a description of the output fields
                    var functionDescription = "Returns " + new ShimParameter(reg.FunctionLambda.ReturnType).HelpString;

                    // replace the Argument description, with a description of the input fields
                    var parameterDescriptions = new string[reg.FunctionLambda.Parameters.Count];
                    for (int param = 0; param != reg.FunctionLambda.Parameters.Count; ++param)
                    {
                        parameterDescriptions[param] = "Input " +
                                                       new ShimParameter(reg.FunctionLambda.Parameters[param].Type)
                                                           .HelpString;
                    }

                    if (parameterDescriptions.GetLength(0) != reg.ParameterRegistrations.Count)
                        throw new InvalidOperationException(
                            string.Format("Unexpected number of parameter registrations {0} vs {1}",
                                parameterDescriptions.GetLength(0), reg.ParameterRegistrations.Count));

                    // all ok - modify the registration
                    reg.FunctionLambda = shim;
                    if(String.IsNullOrEmpty(reg.FunctionAttribute.Description))
                        reg.FunctionAttribute.Description = functionDescription;
                    for (int param = 0; param != reg.ParameterRegistrations.Count; ++param)
                    {
                        if (String.IsNullOrEmpty(reg.ParameterRegistrations[param].ArgumentAttribute.Description))
                            reg.ParameterRegistrations[param].ArgumentAttribute.Description =
                                parameterDescriptions[param];
                    }
                }
                catch
                {
                    // failed to shim, just pass on the original
                }
                yield return reg;
            }
        }

        private delegate object ParamsDelegate(params object[] args);

        /// <summary>
        /// Function which creates a shim for a target method.
        /// The target method is expected to take 1 or more enumerables of various types, and return a single enumerable of another type.
        /// The shim is a lambda expression which takes 1 or more object[,] parameters, and returns a single object[,]
        /// The first row of each array defines the field names, which are mapped to the public properties of the
        /// input and return types.
        /// </summary>
        /// <param name="targetMethod"></param>
        /// <returns></returns>
        internal static LambdaExpression MakeObjectArrayShim(this LambdaExpression targetMethod)
        {
            var nParams = targetMethod.Parameters.Count;

            // validate and extract info for the input and return types
            var inputShimParameters = targetMethod.Parameters.Select(param => new ShimParameter(param.Type)).ToList();
            var resultShimParameter = new ShimParameter(targetMethod.ReturnType);

            var compiledTargetMethod = targetMethod.Compile();

            // create a delegate, object*n -> object
            // (simpler, but probably slower, alternative to building it all out of Expressions)
            ParamsDelegate shimDelegate = inputObjectArray =>
            {
                if (inputObjectArray.GetLength(0) != nParams)
                    throw new InvalidOperationException(string.Format("Expected {0} params, received {1}", nParams,
                        inputObjectArray.GetLength(0)));

                var targetMethodInputs = new object[nParams];

                for (int i = 0; i != nParams; ++i)
                    targetMethodInputs[i] = inputShimParameters[i].ConvertShimToTarget(inputObjectArray[i]);

                var targetMethodResult = compiledTargetMethod.DynamicInvoke(targetMethodInputs);

                return resultShimParameter.ConvertTargetToShim(targetMethodResult);
            };

            // convert the delegate back to a LambdaExpression
            var args = targetMethod.Parameters.Select(param => Expression.Parameter(typeof(object))).ToList();
            var paramsParam = Expression.NewArrayInit(typeof(object), args);
            var closure = Expression.Constant(shimDelegate.Target);
            var call = Expression.Call(closure, shimDelegate.Method, paramsParam);
            return Expression.Lambda(call, args);
        }

        /// <summary>
        /// Class which does the work of translating a parameter or return value
        /// between the shim (e.g. object[,] or object) and the target (e.g. IEnumerable or a value type)
        /// </summary>
        private class ShimParameter
        {
            private Type Type { set; get; }
            private bool CanMapToArray { get { return MappedRecordProperties != null && MappedRecordType != null; } }
            private Type MappedRecordType { set; get; }
            private PropertyInfo[] MappedRecordProperties { set; get; }

            public string HelpString
            {
                get
                {
                    if (CanMapToArray)
                        return "array, with header row containing:\n" + String.Join(",", this.MappedRecordProperties as IEnumerable<PropertyInfo>);

                    return "value, of type " + Type.Name;
                }
            }

            public ShimParameter(Type type)
            {
                Type = type;
                if (!type.IsGenericType || type.Name != typeof(IEnumerable<>).Name)
                    return;

                var typeArgs = type.GetGenericArguments();
                if (typeArgs.Length != 1)
                    return;

                Type recordType = typeArgs[0];
                PropertyInfo[] recordProperties =
                    recordType.GetMembers(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).
                        OfType<PropertyInfo>().ToArray();
                if (recordProperties.Length == 0)
                    return;

                MappedRecordType = recordType;
                MappedRecordProperties = recordProperties;
            }

            /// <summary>
            /// Converts a parameter from the shim (e.g. object[,]) to the target (e.g. IEnumerable)
            /// </summary>
            /// <param name="inputObject"></param>
            /// <returns></returns>
            public object ConvertShimToTarget(object inputObject)
            {
                var objectArray = inputObject as object[,];
                if (!this.CanMapToArray || objectArray == null)
                {
                    // can't map target to array, so just convert raw value
                    return ConvertFromExcelObject(inputObject, this.Type);
                }

                if (objectArray.GetLength(0) == 0)
                    throw new ArgumentException("objectArray");

                // extract nrows and ncols for each input array
                int nInputRows = objectArray.GetLength(0) - 1;
                int nInputCols = objectArray.GetLength(1);

                // Decorate the input record properties with the matching
                // column indices from the input array. We have to do this each time
                // the shim is invoked to map column headers dynamically.
                // Would this be better as a SelectMany?
                var inputPropertyCols = this.MappedRecordProperties.Select(propInfo =>
                {
                    int colIndex = -1;

                    for (int inputCol = 0; inputCol != nInputCols; ++inputCol)
                    {
                        var colName = objectArray[0, inputCol] as string;
                        if (colName == null)
                            continue;

                        if (propInfo.Name.Equals(colName, StringComparison.OrdinalIgnoreCase))
                        {
                            colIndex = inputCol;
                            break;
                        }
                    }
                    return Tuple.Create(propInfo, colIndex);
                }).ToArray();

                // create a sequence of InputRecords
                Array records = Array.CreateInstance(this.MappedRecordType, nInputRows);

                // populate it
                for (int row = 0; row != nInputRows; ++row)
                {
                    object inputRecord;
                    try
                    {
                        // try using constructor which takes parameters in their declared order
                        inputRecord = Activator.CreateInstance(this.MappedRecordType,
                            inputPropertyCols.Select(
                                prop =>
                                    ConvertFromExcelObject(objectArray[row + 1, prop.Item2],
                                        prop.Item1.PropertyType)).ToArray());
                    }
                    catch (MissingMethodException)
                    {
                        // try a different way... default constructor and then set properties
                        inputRecord = Activator.CreateInstance(this.MappedRecordType);

                        // populate the record
                        foreach (var prop in inputPropertyCols)
                            prop.Item1.SetValue(inputRecord,
                                ConvertFromExcelObject(objectArray[row + 1, prop.Item2],
                                    prop.Item1.PropertyType), null);
                    }

                    records.SetValue(inputRecord, row);
                }
                return records;
            }

            /// <summary>
            /// Converts a parameter from the target (e.g. IEnumerable) to the shim (e.g. object[,])
            /// </summary>
            /// <param name="outputObject"></param>
            /// <returns></returns>
            public object ConvertTargetToShim(object outputObject)
            {
                if (!this.CanMapToArray)
                    return outputObject;

                var genericToArray =
                    typeof(Enumerable).GetMethods(BindingFlags.Static | BindingFlags.Public)
                        .First(mi => mi.Name == "ToArray");
                if (genericToArray == null)
                    throw new InvalidOperationException("Internal error. Failed to find Enumerable.ToArray");
                var toArray = genericToArray.MakeGenericMethod(this.MappedRecordType);
                var returnRecordArray = toArray.Invoke(null, new object[] { outputObject }) as Array;
                if (returnRecordArray == null)
                    throw new InvalidOperationException("Internal error. Failed to convert return record to Array");

                // create a return object array and populate the first row
                var nReturnRows = returnRecordArray.Length;
                var returnObjectArray = new object[nReturnRows + 1, this.MappedRecordProperties.Length];
                for (int outputCol = 0; outputCol != this.MappedRecordProperties.Length; ++outputCol)
                    returnObjectArray[0, outputCol] = this.MappedRecordProperties[outputCol].Name;

                // iterate through the entire array and populate the output
                for (int returnRow = 0; returnRow != nReturnRows; ++returnRow)
                {
                    for (int returnCol = 0; returnCol != this.MappedRecordProperties.Length; ++returnCol)
                    {
                        returnObjectArray[returnRow + 1, returnCol] = this.MappedRecordProperties[returnCol].
                            GetValue(returnRecordArray.GetValue(returnRow), null);
                    }
                }

                return returnObjectArray;
            }

            /// <summary>
            /// Wrapper for Convert.ChangeType which understands Excel's use of doubles as OADates.
            /// </summary>
            /// <param name="from">Excel object to convert into a different .NET type</param>
            /// <param name="toType">Type to convert to</param>
            /// <returns>Converted object</returns>
            private static object ConvertFromExcelObject(object from, Type toType)
            {
                // special case when converting from Excel double to DateTime
                // no need for special case in reverse, because Excel-DNA understands a DateTime object
                if (toType == typeof(DateTime) && @from is double)
                {
                    return DateTime.FromOADate((double)@from);
                }
                return Convert.ChangeType(@from, toType);
            }
        }
    }
}
