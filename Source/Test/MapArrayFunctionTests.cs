using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using NUnit.Framework;
using Orientation = ExcelDna.Registration.MapArrayFunctionRegistration.Orientation;

namespace ExcelDna.Registration.Test
{
    [TestFixture]
    public static class MapArrayFunctionTests
    {
        //////////////////////////////////////////////////////////////////////////////////////////////
        // Define 4 classes which we'll use as IEnumerable record types.
        #region Record Classes

        public interface IHaveSomeProperties
        {
            double D { set; get; }
            int I { set; get; }
            string S { set; get; }
            DateTime Dt { set; get; }
        }
        /// <summary>
        /// Struct, with initialising constructor
        /// </summary>
        public struct TestStructWithCtor
        {
            public TestStructWithCtor(double d, int i, string s, DateTime dt)
            {
                _d = d;
                _i = i;
                _s = s;
                _dt = dt;
            }

            private double _d;
            public double D { set { _d = value; } get { return _d; } }

            private int _i;
            public int I { set { _i = value; } get { return _i; } }

            private string _s;
            public string S { set { _s = value; } get { return _s; } }

            private DateTime _dt;
            public DateTime Dt { set { _dt = value; } get { return _dt; } }
        }

        /// <summary>
        /// Class, with initialising constructor, equivalent to F# Record type
        /// </summary>
        public class TestClassWithCtor
        {
            public TestClassWithCtor(double d, int i, string s, DateTime dt)
            {
                D = d;
                I = i;
                S = s;
                Dt = dt;
            }

            public double D { get; set; }
            public int I { get; set; }
            public string S { get; set; }
            public DateTime Dt { get; set; }
        }

        /// <summary>
        /// Struct, with no initialising constructor
        /// </summary>
        public struct TestStructDefaultCtor : IHaveSomeProperties
        {
            public double D { get; set; }
            public int I { get; set; }
            public string S { get; set; }
            public DateTime Dt { get; set; }
        }

        /// <summary>
        /// Class, with no initialising constructor
        /// </summary>
        public class TestClassDefaultCtor : IHaveSomeProperties
        {
            public double D { get; set; }
            public int I { get; set; }
            public string S { get; set; }
            public DateTime Dt { get; set; }
        }

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Test Case class

        public class TestCase
        {
            public TestCase(MethodInfo methodInfo, object[] inputData, object expectedOutputData, string name = "")
            {
                MethodInfo = methodInfo;
                InputData = inputData;
                ExpectedOutputData = expectedOutputData;
                Name =
                    (String.IsNullOrEmpty(name) ? MethodInfo.ToString() : name);
            }
            public MethodInfo MethodInfo { set; get; }
            public object[] InputData { set; get; }
            public object ExpectedOutputData { set; get; }
            public string Name { set; get; }

            // for nunit's convenience
            public override string ToString()
            {
                return Name;
            }
        }

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Test Data

        private static readonly object[,] _recordsInputData =
        {
            // input data can have fields in any order, with different case
            { "I", "S", "D", "DT" },

            { 123, "123", 123.0, new DateTime(2014,03,10,17,40,21) },
            { 456, "456", 456.0, new DateTime(2001,11,23,22,45,00) },
            { 56789.3, 56789, 56789, 41910.0 }  // Test conversion from non-regular types
        };
        private static readonly TestClassWithCtor _mixedInputsRecord = new TestClassWithCtor(111.1, 222, "333", new DateTime(2044, 4, 4, 4, 44, 44));

        private static readonly object[,] _expectedRecordsOutputData1To1 = 
        {
            // output data fields are determined by type, not data
            { "D", "I", "S", "Dt" },

            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) }
        };
        private static readonly object[,] _expectedRecordsOutputData2To1 = 
        {
            { "D", "I", "S", "Dt" },

            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) },
            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) }
        };
        private static readonly object[,] _expectedOutputDataMixedWithAppend = 
        {
            { "D", "I", "S", "Dt" },

            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
            { _mixedInputsRecord.D, _mixedInputsRecord.I, _mixedInputsRecord.S, _mixedInputsRecord.Dt }
        };
        private static readonly object[,] _expectedOutputDataMixedNoAppend = 
        {
            { "D", "I", "S", "Dt" },

            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
        };

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Target Methods for Shim

        /// <summary>
        /// A method for us to test, using ExcelMapPropertiesToColumnHeaders
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [return: ExcelMapPropertiesToColumnHeaders]
        public static IEnumerable<T> ReverseEnumerable<T>(
            [ExcelMapPropertiesToColumnHeaders] IEnumerable<T> input)
        {
            return input.Reverse();
        }

        /// <summary>
        /// Same again, but using a more specific IEnumerable type
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        [return: ExcelMapPropertiesToColumnHeaders]
        public static IList<T> ReverseList<T>(
            [ExcelMapPropertiesToColumnHeaders] IList<T> input)
        {
            return input.Reverse().ToList();
        }

        /// <summary>
        /// Same again, but with a non-generic enumerable
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static IEnumerable ReverseNonGenericNoMapping(
             IEnumerable input)
        {
            return input.Cast<object>().Reverse();
        }

        /// <summary>
        /// A method for us to test, without ExcelMapPropertiesToColumnHeaders
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static IEnumerable<T> ReverseNoMapping<T>(
            IEnumerable<T> input)
        {
            return input.Reverse();
        }

        [return: ExcelMapPropertiesToColumnHeaders]
        public static IEnumerable<T> CombineAndReverse<T>(
            [ExcelMapPropertiesToColumnHeaders] IEnumerable<T> input1, 
            [ExcelMapPropertiesToColumnHeaders] IEnumerable<T> input2)
        {
            return input1.Concat(input2).Reverse();
        }

        /// <summary>
        /// The purpose of this method is to test shimming of a function with a mix of sequence and plain value
        /// parameters
        /// </summary>
        [return: ExcelMapPropertiesToColumnHeaders]
        public static IEnumerable<T> AppendOne<T>(
            [ExcelMapPropertiesToColumnHeaders] IEnumerable<T> input, bool append, double doubleValue, int intValue, string stringValue, DateTime dateTimeValue)
            where T:IHaveSomeProperties,new()
        {
            var t = new T {I=intValue, S=stringValue, Dt = dateTimeValue, D=doubleValue };
            if(append)
                return input.Concat(new[] { t });
            return input;
        }

        #endregion

        #region Static methods for tests
        public static bool Not(bool x) { return !x; }
        public static bool And(bool x, bool y) { return x && y; }
        public static double Plus1Double(double x) { return x + 1; }
        public static int Plus1Int(int x) { return x+1; }
        public static string ToUpper(string s) { return s.ToUpper(); }
        public static DateTime Identity(DateTime x) { return x; }
        public static bool ClassDefaultCtorParam(TestClassDefaultCtor x) { return x.I == 0; }
        public static bool StructDefaultCtorParam(TestStructDefaultCtor x) { return x.I == 0; }
        public static bool ClassWithCtorParam(TestClassWithCtor x) { return x == null; }
        public static bool StructWithCtorParam(TestStructWithCtor x) { return x.I == 0; }
        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Test Cases

        public static TestCase[] TestCases =
        {
            // Functions which take and return simple value types - no mapping
            new TestCase(((Func<bool, bool>)Not).Method, new object[] { true }, false),
            new TestCase(((Func<double, double>)Plus1Double).Method, new object[] { 123.0 }, 124.0 ),
            new TestCase(((Func<int, int>)Plus1Int).Method, new object[] { 123 }, 124 ),
            new TestCase(((Func<string, string>)ToUpper).Method, new object[] { "hello" }, "HELLO" ),
                // Excel provides dates as doubles. The shim will pass them back as DateTime, because Excel-DNA will convert for us.
            new TestCase(((Func<DateTime, DateTime>)Identity).Method, new object[] { 41757 }, DateTime.FromOADate(41757)),
            new TestCase(((Func<DateTime, DateTime>)Identity).Method, new object[] { 41757.123 }, DateTime.FromOADate(41757.123)),

            // Functions which take and return sequences of plain value types - no mapping
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (int)),
                new [] { (object)23 },  // pass in a single item instead of array
                new [,] { { (object)23 } },  // still produces an array output
                "ReverseNoMapping int (single)"), 
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (int)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)i).ToArray2D(Orientation.Horizontal) },
                Enumerable.Range(0, 10).Select(i => (object)(9-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping int (horizontal)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (int)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)i).ToArray2D(Orientation.Vertical) },
                Enumerable.Range(0, 10).Select(i => (object)(9-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping int (vertical)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (double)),
                new [] { (object)23.45 },  // pass in a single item instead of array
                new [,] { { (object)23.45 } },  // still produces an array output
                "ReverseNoMapping double (single)"), 
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (double)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)(double)i).ToArray2D(Orientation.Horizontal) },
                Enumerable.Range(0, 10).Select(i => (object)(double)(9-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping double (horizontal)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (double)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)(double)i).ToArray2D(Orientation.Vertical) },
                Enumerable.Range(0, 10).Select(i => (object)(double)(9-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping double (vertical)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (DateTime)),
                new [] { (object)40000 },
                new [,] { { (object)DateTime.FromOADate(40000) } },
                "ReverseNoMapping datetime (single)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (DateTime)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)(i+40000)).ToArray2D(Orientation.Horizontal) },
                Enumerable.Range(0, 10).Select(i => (object)DateTime.FromOADate(40009-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping datetime (horizontal)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNoMapping").MakeGenericMethod(typeof (DateTime)),
                new [] { Enumerable.Range(0, 10).Select(i => (object)(i+40000)).ToArray2D(Orientation.Vertical) },
                Enumerable.Range(0, 10).Select(i => (object)DateTime.FromOADate(40009-i)).ToArray2D(Orientation.Vertical),
                "ReverseNoMapping datetime (vertical)"),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseNonGenericNoMapping"),
                new [] { Enumerable.Range(0, 10).ToArray2D(Orientation.Vertical) },
                Enumerable.Range(0, 10).Select(i => 9 - i).ToArray2D(Orientation.Vertical),
                "ReverseNonGenericNoMapping"),

            // Check handling of empty values, for value types, strings, and classes
            new TestCase(((Func<bool, bool>)Not).Method, new object[] { ExcelEmpty.Value }, true),
            new TestCase(((Func<double, double>)Plus1Double).Method, new object[] { ExcelEmpty.Value }, 1.0 ),
            new TestCase(((Func<int, int>)Plus1Int).Method, new object[] { ExcelEmpty.Value }, 1 ),
            new TestCase(((Func<string, string>)ToUpper).Method, new object[] { ExcelEmpty.Value }, "" ),
            new TestCase(((Func<DateTime, DateTime>)Identity).Method, new object[] { ExcelEmpty.Value }, new DateTime()),
            new TestCase(((Func<TestClassDefaultCtor, bool>)ClassDefaultCtorParam).Method, new object[] { ExcelEmpty.Value }, true),
            new TestCase(((Func<TestStructDefaultCtor, bool>)StructDefaultCtorParam).Method, new object[] { ExcelEmpty.Value }, true),
            new TestCase(((Func<TestClassWithCtor, bool>)ClassWithCtorParam).Method, new object[] { ExcelEmpty.Value }, true),
            new TestCase(((Func<TestStructWithCtor, bool>)StructWithCtorParam).Method, new object[] { ExcelEmpty.Value }, true),

            // Functions which take a sequence, and return a sequence, using ExcelMapPropertiesToColumnHeaders
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseEnumerable").MakeGenericMethod(typeof (TestStructWithCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseEnumerable").MakeGenericMethod(typeof (TestClassWithCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseEnumerable").MakeGenericMethod(typeof (TestStructDefaultCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseEnumerable").MakeGenericMethod(typeof (TestClassDefaultCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseList").MakeGenericMethod(typeof (TestStructWithCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseList").MakeGenericMethod(typeof (TestClassWithCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseList").MakeGenericMethod(typeof (TestStructDefaultCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("ReverseList").MakeGenericMethod(typeof (TestClassDefaultCtor)),
                new [] { _recordsInputData },
                _expectedRecordsOutputData1To1),

            // Functions which take 2 sequences of records, and return a sequence of records, using ExcelMapPropertiesToColumnHeaders
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestStructWithCtor)),
                new [] { _recordsInputData, _recordsInputData },
                _expectedRecordsOutputData2To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestClassWithCtor)),
                new [] { _recordsInputData, _recordsInputData },
                _expectedRecordsOutputData2To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestStructDefaultCtor)),
                new [] { _recordsInputData, _recordsInputData },
                _expectedRecordsOutputData2To1),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestClassDefaultCtor)),
                new [] { _recordsInputData, _recordsInputData },
                _expectedRecordsOutputData2To1),

            // Functions which take a mixture of sequence and value types, and return a sequence
            // with ExcelMapPropertiesToColumnHeaders
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestStructDefaultCtor)),
                new object[] { _recordsInputData, true, _mixedInputsRecord.D, _mixedInputsRecord.I, _mixedInputsRecord.S, _mixedInputsRecord.Dt },
                _expectedOutputDataMixedWithAppend),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestClassDefaultCtor)),
                new object[] { _recordsInputData, true, _mixedInputsRecord.D, _mixedInputsRecord.I, _mixedInputsRecord.S, _mixedInputsRecord.Dt },
                _expectedOutputDataMixedWithAppend),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestStructDefaultCtor)),
                new object[] { _recordsInputData, false, _mixedInputsRecord.D, _mixedInputsRecord.I, _mixedInputsRecord.S, _mixedInputsRecord.Dt },
                _expectedOutputDataMixedNoAppend),
            new TestCase(typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestClassDefaultCtor)),
                new object[] { _recordsInputData, false, _mixedInputsRecord.D, _mixedInputsRecord.I, _mixedInputsRecord.S, _mixedInputsRecord.Dt },
                _expectedOutputDataMixedNoAppend),

            // Failure to execute shim
            new TestCase(((Func<bool, bool, bool>)And).Method, 
                new object[] { true, "this is not a bool" }, 
                new object[,] { {ExcelError.ExcelErrorValue}, {"Failed to convert parameter 2: String was not recognized as a valid Boolean."} },
                "Execute failure"),
        };

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Test Methods

        [Test, TestCaseSource("TestCases")]
        public static void TestMapArrayRegistrations(TestCase testCase)
        {
            var methodInfo = testCase.MethodInfo;

            ////////////////////////////////////////////
            // arrange

            Assert.IsNotNull(methodInfo);

            // wrap the method in a registrationentry
            var registration = new ExcelFunctionRegistration(methodInfo);
            registration.FunctionAttribute = new ExcelMapArrayFunctionAttribute();

            //////////////////////////////////////////
            // act - process the registration object

            var processed = Enumerable.Repeat(registration, 1).ProcessMapArrayFunctions().ToList();

            //////////////////////////////////////////
            // assert

            Assert.AreEqual(1, processed.Count());
            var processedRegistration = processed.First();

            var functionDescription = registration.FunctionAttribute.Description;
            Assert.IsTrue(functionDescription.StartsWith("Returns "));
            var expectedOutputArray = testCase.ExpectedOutputData as object[,];
            if (expectedOutputArray != null && expectedOutputArray.GetLength(0) != 1 && expectedOutputArray.GetLength(1) != 1)
            {
                // confirm function description contains list of field names for array result
                for (int col = 0; col != expectedOutputArray.GetLength(1); ++col)
                {
                    var colHeader = expectedOutputArray[0, col] as string;
                    Assert.IsNotNull(colHeader);
                    var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                    Assert.IsTrue(regex.Match(functionDescription).Success, functionDescription);
                }
            }

            Assert.AreEqual(testCase.InputData.GetLength(0), registration.ParameterRegistrations.Count);
            for (int param = 0; param != testCase.InputData.GetLength(0); ++param)
            {
                var inputArray = testCase.InputData[param] as object[,];
                if(inputArray != null && inputArray.GetLength(0) != 1 && inputArray.GetLength(1) != 1)
                    // confirm function description contains list of field names for this array param
                    for (int col = 0; col != inputArray.GetLength(1); ++col)
                    {
                        var colHeader = _recordsInputData[0, col] as string;
                        Assert.IsNotNull(colHeader);
                        var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                        var parameterDescription = registration.ParameterRegistrations[param].ArgumentAttribute.Description;
                        Assert.IsTrue(regex.Match(parameterDescription).Success, parameterDescription);
                    }
            }

            //////////////////////////////////////////
            // act - invoke the delegate

            var output = processedRegistration.FunctionLambda.Compile().DynamicInvoke(testCase.InputData);

            //////////////////////////////////////////
            // assert

            if(output == null)
                Assert.IsNull(output);
            else
                Assert.AreEqual(testCase.ExpectedOutputData, output);
        }

        #endregion
    }
}
