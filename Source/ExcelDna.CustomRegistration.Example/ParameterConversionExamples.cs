using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration.Example
{
    // TODO: Fix double->int and similar conversions when objects are received
    //       Test multi-hop return type conversions
    //       Test ExcelArgumentAttributes are preserved.

    public static class ParameterConversionExamples
    {
        [ExcelFunction]
        public static string dnaParameterConvertTest(double? optTest)
        {
            if (!optTest.HasValue) return "NULL!!!";

            return optTest.Value.ToString("F1");
        }

        [ExcelFunction]
        public static string dnaParameterConvertOptionalTest(double optOptTest = 42.0)
        {
            return "VALUE: " + optOptTest.ToString("F1");
        }

        [ExcelFunction]
        public static string dnaMultipleOptional(double optOptTest1 = 3.14159265, string optOptTest2 = "@42@")
        {
            return "VALUES: " + optOptTest1.ToString("F7") + " & " + optOptTest2;
        }

        // Problem function
        // This function cannot be called yet, since the cast from what Excel passes in (object) to the int we expect fails.
        // It will need some improved conversion function in the OptionalParameterConversion.
        [ExcelFunction]
        public static string dnaOptionalIntFail(int optOptTest = 42)
        {
            return "VALUE: " + optOptTest.ToString("F1");
        }
    }

    // Here I test some custom conversions, including a two-hop conversion
    public class TestType1
    {
        public string Value;
        public TestType1(string value)
        {
            Value = value;
        }

        public override string ToString()
        {
            return "From Type 1 with " + Value;
        }

        [ExcelFunction]
        public static string dnaTestFunction1(TestType1 tt)
        {
            return "The Test (1) value is " + tt.Value;
        }
    }

    public class TestType2
    {
        readonly TestType1 _value;
        public TestType2(TestType1 value)
        {
            _value = value;
        }

        // Must be converted using TestType1 => TestType2, then string => TestType1
        [ExcelFunction]
        public static string dnaTestFunction2(TestType2 tt)
        {
            return "The Test (2) value is " + tt._value.Value;
        }

        [ExcelFunction]
        public static TestType1 dnaTestFunction2Ret1(TestType2 tt)
        {
            return new TestType1("The Test (2) value is " + tt._value.Value);
        }
    }
}
