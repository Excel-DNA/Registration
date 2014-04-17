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
                    var functionDescription = "Returns an array, with header row containing:\n" + 
                        GetEnumerableRecordTypeHelpString(reg.FunctionLambda.ReturnType);

                    // replace the Argument description, with a description of the input fields
                    var parameterDescriptions = new string[reg.FunctionLambda.Parameters.Count];
                    for (int param = 0; param != reg.FunctionLambda.Parameters.Count; ++param)
                    {
                        parameterDescriptions[param] = "Input array, with header row containing:\n" +
                            GetEnumerableRecordTypeHelpString(reg.FunctionLambda.Parameters[param].Type);
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
            if (toType == typeof(DateTime) && from is double)
            {
                return DateTime.FromOADate((double)from);
            }
            return Convert.ChangeType(from, toType);
        }

        /// <summary>
        /// Returns a description string which is used in the Excel function dialog.
        /// E.g. IEnumerable<typeparamref name="T"/> produces a string with the public properties
        /// of type T.
        /// This really belongs in a View class somewhere...
        /// </summary>
        /// <param name="enumerableType">Enumerable type for which the help string is required</param>
        /// <returns>Formatted help string, complete with whitespace and newlines</returns>
        private static string GetEnumerableRecordTypeHelpString(Type enumerableType)
        {
            PropertyInfo[] recordProperties = GetEnumerableRecordInfo(enumerableType).Item2;
            var headerFields = new StringBuilder();
            string delim = "";
            foreach (var prop in recordProperties)
            {
                headerFields.Append(delim);
                headerFields.Append(prop.Name);
                delim = ",";
            }
            return headerFields.ToString();
        }

        /// <summary>
        /// Returns information about the Record type which an enumerator yields.
        /// A valid record type must have at least 1 public property.
        /// </summary>
        /// <param name="enumerableType">Enumerable type to reflect</param>
        /// <returns>A tuple containing the Type of the record, and an array of its public member properties</returns>
        private static Tuple<Type, PropertyInfo[]> GetEnumerableRecordInfo(Type enumerableType)
        {
            if (!enumerableType.IsGenericType || enumerableType.Name != "IEnumerable`1")
                throw new ArgumentException(string.Format("Type {0} is not IEnumerable<>", enumerableType), "enumerableType");

            var typeArgs = enumerableType.GetGenericArguments();
            if (typeArgs.Length != 1)
                throw new ArgumentException(string.Format("Unexpected {0} generic args for type {1}",
                    typeArgs.Length, enumerableType.Name));

            Type recordType = typeArgs[0];
            PropertyInfo[] recordProperties =
                recordType.GetMembers(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).
                    OfType<PropertyInfo>().ToArray();
            if (recordProperties.Length == 0)
                throw new ArgumentException(string.Format("Unsupported record type {0} with 0 public properties",
                    recordType.Name));

            return Tuple.Create(recordType, recordProperties);
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

            // validate and extract property info for the input and return types
            var inputRecordInfo = targetMethod.Parameters.Select(param => GetEnumerableRecordInfo(param.Type)).ToList();
            var returnRecordInfo = GetEnumerableRecordInfo(targetMethod.ReturnType);
            Type[] inputRecordType = inputRecordInfo.Select(info => info.Item1).ToArray();
            Type returnRecordType = returnRecordInfo.Item1;
            PropertyInfo[][] inputRecordProperties = inputRecordInfo.Select(info => info.Item2).ToArray();
            PropertyInfo[] returnRecordProperties = returnRecordInfo.Item2;

            // create the delegate, object[,]*n -> object[,]
            ParamsDelegate shimDelegate = inputObjectArray =>
            {
                if (inputObjectArray.GetLength(0) != nParams)
                    throw new InvalidOperationException(string.Format("Expected {0} params, received {1}", nParams,
                        inputObjectArray.GetLength(0)));

                var inputRecordArray = new Array[nParams];

                for (int i = 0; i != nParams; ++i)
                {
                    var inputObject = inputObjectArray[i] as object[,];
                    if (inputObject == null)
                    {
                        throw new InvalidOperationException(string.Format("Parameter {0} of {1} is not an Array", i + 1,
                            nParams));
                    }
                    if (inputObject.GetLength(0) == 0)
                        throw new ArgumentException();

                    // extract nrows and ncols for each input array
                    int nInputRows = inputObject.GetLength(0) - 1;
                    int nInputCols = inputObject.GetLength(1);

                    // Decorate the input record properties with the matching
                    // column indices from the input array. We have to do this each time
                    // the shim is invoked to map column headers dynamically.
                    // Would this be better as a SelectMany?
                    var inputPropertyCols = inputRecordProperties[i].Select(propInfo =>
                    {
                        int colIndex = -1;

                        for (int inputCol = 0; inputCol != nInputCols; ++inputCol)
                        {
                            var colName = inputObject[0, inputCol] as string;
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
                    inputRecordArray[i] = Array.CreateInstance(inputRecordType[i], nInputRows);

                    // populate it
                    for (int row = 0; row != nInputRows; ++row)
                    {
                        object inputRecord;
                        try
                        {
                            // try using constructor which takes parameters in their declared order
                            inputRecord = Activator.CreateInstance(inputRecordType[i],
                                inputPropertyCols.Select(
                                    prop =>
                                        ConvertFromExcelObject(inputObject[row + 1, prop.Item2],
                                            prop.Item1.PropertyType)).ToArray());
                        }
                        catch (MissingMethodException)
                        {
                            // try a different way... default constructor and then set properties
                            inputRecord = Activator.CreateInstance(inputRecordType[i]);

                            // populate the record
                            foreach (var prop in inputPropertyCols)
                                prop.Item1.SetValue(inputRecord,
                                    ConvertFromExcelObject(inputObject[row + 1, prop.Item2],
                                        prop.Item1.PropertyType), null);
                        }

                        inputRecordArray[i].SetValue(inputRecord, row);
                    }
                }

                // invoke the method
                var returnRecordSequence = targetMethod.Compile().DynamicInvoke(inputRecordArray);

                // turn it ito an Array<OutputRecordType>
                var genericToArray =
                    typeof(Enumerable).GetMethods(BindingFlags.Static | BindingFlags.Public)
                        .First(mi => mi.Name == "ToArray");
                if (genericToArray == null)
                    throw new InvalidOperationException("Internal error. Failed to find Enumerable.ToArray");
                var toArray = genericToArray.MakeGenericMethod(returnRecordType);
                var returnRecordArray = toArray.Invoke(null, new object[] { returnRecordSequence }) as Array;
                if (returnRecordArray == null)
                    throw new InvalidOperationException("Internal error. Failed to convert return record to Array");

                // create a return object array and populate the first row
                var nReturnRows = returnRecordArray.Length;
                var returnObjectArray = new object[nReturnRows + 1, returnRecordProperties.Length];
                for (int outputCol = 0; outputCol != returnRecordProperties.Length; ++outputCol)
                    returnObjectArray[0, outputCol] = returnRecordProperties[outputCol].Name;

                // iterate through the entire array and populate the output
                for (int returnRow = 0; returnRow != nReturnRows; ++returnRow)
                {
                    for (int returnCol = 0; returnCol != returnRecordProperties.Length; ++returnCol)
                    {
                        returnObjectArray[returnRow + 1, returnCol] = returnRecordProperties[returnCol].
                            GetValue(returnRecordArray.GetValue(returnRow), null);
                    }
                }

                return returnObjectArray;
            };

            var args = targetMethod.Parameters.Select(param => Expression.Parameter(typeof(object[,]))).ToList();
            var paramsParam = Expression.NewArrayInit(typeof(object), args);
            var closure = Expression.Constant(shimDelegate.Target);
            var call = Expression.Call(closure, shimDelegate.Method, paramsParam);
            return Expression.Lambda(call, args);
        }
    }
}
