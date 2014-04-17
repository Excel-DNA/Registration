using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ExcelDna.CustomRegistration.Test
{
    [TestFixture]
    public class MapArrayFunctionTests
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
        // Define some methods which we want to create Shims for, to use in Excel
        // Signature is IEnumerable<T> -> IEnumerable<T>
        #region Target Methods for Shim

        /// <summary>
        /// A method for us to test
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static IEnumerable<T> Reverse<T>(IEnumerable<T> input)
        {
            return input.Reverse();
        }

        public static IEnumerable<T> CombineAndReverse<T>(IEnumerable<T> input1, IEnumerable<T> input2)
        {
            return input1.Concat(input2).Reverse();
        }

        public static IEnumerable<T> AppendOne<T>(IEnumerable<T> input, bool append, double doubleValue, int intValue, string stringValue, DateTime dateTimeValue)
            where T:IHaveSomeProperties,new()
        {
            var t = new T {I=intValue, S=stringValue, Dt = dateTimeValue, D=doubleValue };
            if(append)
                return input.Concat(new[] { t });
            return input;
        }

        public static MethodInfo[] Methods1To1 =
        {
            typeof (MapArrayFunctionTests).GetMethod("Reverse").MakeGenericMethod(typeof (TestStructWithCtor)),
            typeof (MapArrayFunctionTests).GetMethod("Reverse").MakeGenericMethod(typeof (TestClassWithCtor)),
            typeof (MapArrayFunctionTests).GetMethod("Reverse").MakeGenericMethod(typeof (TestStructDefaultCtor)),
            typeof (MapArrayFunctionTests).GetMethod("Reverse").MakeGenericMethod(typeof (TestClassDefaultCtor))
        };

        public static MethodInfo[] Methods2To1 =
        {
            typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestStructWithCtor)),
            typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestClassWithCtor)),
            typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestStructDefaultCtor)),
            typeof (MapArrayFunctionTests).GetMethod("CombineAndReverse").MakeGenericMethod(typeof(TestClassDefaultCtor)),
        };

        public static MethodInfo[] MixedTypeMethods =
        {
            typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestStructDefaultCtor)),
            typeof (MapArrayFunctionTests).GetMethod("AppendOne").MakeGenericMethod(typeof(TestClassDefaultCtor)),
        };

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////
        #region Tests

        private static readonly object[,] _inputData =
        {
            // input data can have fields in any order, with different case
            { "I", "S", "D", "DT" },

            { 123, "123", 123.0, new DateTime(2014,03,10,17,40,21) },
            { 456, "456", 456.0, new DateTime(2001,11,23,22,45,00) },
            { 56789.3, 56789, 56789, 41910.0 }  // Test conversion from non-regular types
        };
        private static readonly TestClassWithCtor _mixedInputs = new TestClassWithCtor(111.1, 222, "333", new DateTime(2044,4,4,4,44,44));

        private static readonly object[,] _expectedOutputData1To1 = 
        {
            // output data fields are determined by type, not data
            { "D", "I", "S", "Dt" },

            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) }
        };
        private static readonly object[,] _expectedOutputData2To1 = 
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
            { _mixedInputs.D, _mixedInputs.I, _mixedInputs.S, _mixedInputs.Dt }
        };
        private static readonly object[,] _expectedOutputDataMixedNoAppend = 
        {
            { "D", "I", "S", "Dt" },

            { 123.0, 123, "123", new DateTime(2014,03,10,17,40,21) },
            { 456.0, 456, "456", new DateTime(2001,11,23,22,45,00) },
            { 56789.0, 56789, "56789", DateTime.FromOADate(41910.0) },
        };

        [Test, TestCaseSource("Methods1To1")]
        public void TestMapArrayRegistrations1To1(MethodInfo methodInfo)
        {
            ////////////////////////////////////////////
            // arrange

            Assert.IsNotNull(methodInfo);

            // wrap the method in a registrationentry
            var registration = new ExcelFunctionRegistration(methodInfo)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };

            //////////////////////////////////////////
            // act

            var processed = Enumerable.Repeat(registration, 1).ProcessMapArrayFunctions().ToList();

            //////////////////////////////////////////
            // assert

            Assert.AreEqual(1, processed.Count());
            var processedRegistration = processed.First();

            var functionDescription = registration.FunctionAttribute.Description;
            for (int col = 0; col != _expectedOutputData1To1.GetLength(1); ++col)
            {
                var colHeader = _expectedOutputData1To1[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(functionDescription).Success, functionDescription);
            }

            Assert.AreEqual(1, registration.ParameterRegistrations.Count);
            var parameterDescription = registration.ParameterRegistrations[0].ArgumentAttribute.Description;
            for (int col = 0; col != _inputData.GetLength(1); ++col)
            {
                var colHeader = _inputData[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(parameterDescription).Success, parameterDescription);
            }

            //////////////////////////////////////////
            // act

            // invoke the delegate
            var output = processedRegistration.FunctionLambda.Compile().DynamicInvoke(_inputData);

            //////////////////////////////////////////
            // assert

            Assert.IsNotNull(output);
            Assert.AreEqual(_expectedOutputData1To1, output);
        }

        [Test, TestCaseSource("Methods2To1")]
        public void TestMapArrayRegistrations2To1(MethodInfo methodInfo)
        {
            ////////////////////////////////////////////
            // arrange

            Assert.IsNotNull(methodInfo);

            // wrap the method in a registrationentry
            var registration = new ExcelFunctionRegistration(methodInfo);
            registration.FunctionAttribute = new ExcelMapArrayFunctionAttribute();

            //////////////////////////////////////////
            // act

            var processed = Enumerable.Repeat(registration, 1).ProcessMapArrayFunctions().ToList();

            //////////////////////////////////////////
            // assert

            Assert.AreEqual(1, processed.Count());
            var processedRegistration = processed.First();

            var functionDescription = registration.FunctionAttribute.Description;
            for (int col = 0; col != _expectedOutputData1To1.GetLength(1); ++col)
            {
                var colHeader = _expectedOutputData2To1[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(functionDescription).Success, functionDescription);
            }

            Assert.AreEqual(2, registration.ParameterRegistrations.Count);
            var parameterDescription0 = registration.ParameterRegistrations[0].ArgumentAttribute.Description;
            var parameterDescription1 = registration.ParameterRegistrations[0].ArgumentAttribute.Description;
            for (int col = 0; col != _inputData.GetLength(1); ++col)
            {
                var colHeader = _inputData[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(parameterDescription0).Success, parameterDescription0);
                Assert.IsTrue(regex.Match(parameterDescription1).Success, parameterDescription1);
            }

            //////////////////////////////////////////
            // act

            // invoke the delegate
            var output = processedRegistration.FunctionLambda.Compile().DynamicInvoke(_inputData, _inputData);

            //////////////////////////////////////////
            // assert

            Assert.IsNotNull(output);
            Assert.AreEqual(_expectedOutputData2To1, output);
        }

        [Test, TestCaseSource("MixedTypeMethods")]
        public void TestMapArrayRegistrationsMixedTypes(MethodInfo methodInfo)
        {
            ////////////////////////////////////////////
            // arrange

            Assert.IsNotNull(methodInfo);

            // wrap the method in a registrationentry
            var registration = new ExcelFunctionRegistration(methodInfo)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };

            //////////////////////////////////////////
            // act

            var processed = Enumerable.Repeat(registration, 1).ProcessMapArrayFunctions().ToList();

            //////////////////////////////////////////
            // assert

            Assert.AreEqual(1, processed.Count());
            var processedRegistration = processed.First();

            var functionDescription = registration.FunctionAttribute.Description;
            for (int col = 0; col != _expectedOutputDataMixedWithAppend.GetLength(1); ++col)
            {
                var colHeader = _expectedOutputData1To1[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(functionDescription).Success, functionDescription);
            }

            Assert.AreEqual(6, registration.ParameterRegistrations.Count);
            var parameterDescription = registration.ParameterRegistrations[0].ArgumentAttribute.Description;
            for (int col = 0; col != _inputData.GetLength(1); ++col)
            {
                var colHeader = _inputData[0, col] as string;
                Assert.IsNotNull(colHeader);
                var regex = new Regex(@"\b" + colHeader + @"\b", RegexOptions.IgnoreCase);
                Assert.IsTrue(regex.Match(parameterDescription).Success, parameterDescription);
            }
            foreach (
                var pair in registration.ParameterRegistrations.Skip(1).Zip(new[] {"bool", "double", "int", "string", "datetime"}, Tuple.Create))
            {
                Assert.IsTrue(
                    pair.Item1.ArgumentAttribute.Description.IndexOf(pair.Item2, StringComparison.OrdinalIgnoreCase) >= 0,
                        string.Format("Could not find {0} in {1}", pair.Item2, pair.Item1));
            }

            //////////////////////////////////////////
            // act

            // invoke the delegate
            var output = processedRegistration.FunctionLambda.Compile().DynamicInvoke(
                _inputData, true, _mixedInputs.D, _mixedInputs.I, _mixedInputs.S, _mixedInputs.Dt);

            //////////////////////////////////////////
            // assert

            Assert.IsNotNull(output);
            Assert.AreEqual(_expectedOutputDataMixedWithAppend, output);

            //////////////////////////////////////////
            // act

            // invoke the delegate again with append = false
            output = processedRegistration.FunctionLambda.Compile().DynamicInvoke(
                _inputData, false, _mixedInputs.D, _mixedInputs.I, _mixedInputs.S, _mixedInputs.Dt);

            //////////////////////////////////////////
            // assert

            Assert.IsNotNull(output);
            Assert.AreEqual(_expectedOutputDataMixedNoAppend, output);
        }

        [Test]
        public void TestMapArrayRegistrationsNonArrayTypes()
        {
            ////////////////////////////////////////////
            // arrange

            Func<bool, bool> funcBool = x => x;
            Func<double, double> funcDouble = x => x;
            Func<int, int> funcInt = x => x;
            Func<string, string> funcString = x => x;
            Func<DateTime, DateTime> funcDateTime = x => x;

            var registrationBool = new ExcelFunctionRegistration(funcBool.Method)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };
            var registrationDouble = new ExcelFunctionRegistration(funcDouble.Method)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };
            var registrationInt = new ExcelFunctionRegistration(funcInt.Method)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };
            var registrationString = new ExcelFunctionRegistration(funcString.Method)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };
            var registrationDateTime = new ExcelFunctionRegistration(funcDateTime.Method)
            {
                FunctionAttribute = new ExcelMapArrayFunctionAttribute()
            };

            //////////////////////////////////////////
            // act

            var processedBool = Enumerable.Repeat(registrationBool, 1).ProcessMapArrayFunctions().ToList().First();
            var processedDouble = Enumerable.Repeat(registrationDouble, 1).ProcessMapArrayFunctions().ToList().First();
            var processedInt = Enumerable.Repeat(registrationInt, 1).ProcessMapArrayFunctions().ToList().First();
            var processedString = Enumerable.Repeat(registrationString, 1).ProcessMapArrayFunctions().ToList().First();
            var processedDateTime = Enumerable.Repeat(registrationDateTime, 1).ProcessMapArrayFunctions().ToList().First();

            //////////////////////////////////////////
            // assert

            Assert.AreEqual(true, processedBool.FunctionLambda.Compile().DynamicInvoke(true));
            Assert.AreEqual(123.0, processedDouble.FunctionLambda.Compile().DynamicInvoke(123.0));
            Assert.AreEqual(123, processedInt.FunctionLambda.Compile().DynamicInvoke(123));
            Assert.AreEqual("123", processedString.FunctionLambda.Compile().DynamicInvoke("123"));

            // Excel provides dates as doubles. The shim will pass them back as DateTime, because Excel-DNA will convert for us.
            // There's a slight precision lost in this process, so check difference in milliseconds, not ticks.
            var now = DateTime.Now;
            Assert.Less(
                Math.Abs((now - (DateTime) processedDateTime.FunctionLambda.Compile().DynamicInvoke(now.ToOADate())).TotalMilliseconds),
                1.0);
        }

        #endregion
    }
}
