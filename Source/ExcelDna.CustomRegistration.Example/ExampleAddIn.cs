using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration.Example
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! ERROR: " + ex.ToString());

            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            ConfigureParameterConversion();

            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit regisrations are the only ones - there is no default processing
            Registration.GetExcelFunctions()
                        .ProcessParameterConversions()
                        .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                        .ProcessParamsRegistrations()
                        .RegisterFunctions();
        }

        static void ConfigureParameterConversion()
        {
            // CONSIDER: This might have to change if we want to add improved tracing to the conversions.
            // TODO: Parameter vs Return conversions...?

            // Register the Standard Parameter Conversions
            ParameterConversionRegistration.AddParameterConversion(ParameterConversions.NullableConversion);
            ParameterConversionRegistration.AddParameterConversion(ParameterConversions.OptionalConversion);

            // Some ideas ways to define and register conversions
            // These are for a particular parameter type
            // (Func<object, MyType> would allow MyType to be taken as parameter)

            // Inline Lambda - one way
            ParameterConversionRegistration.AddParameterConversion((string value) => new TestType1(value));
            ParameterConversionRegistration.AddParameterConversion((TestType1 value) => new TestType2(value));

            ParameterConversionRegistration.AddReturnConversion((TestType1 value) => value.ToString());
            //ParameterConversionRegistration.AddParameterConversion((string value) => convert2(convert1(value)));

            // Alternative - use method via lambda
            // ParameterConversionRegistration.AddParameterConversion((string input) => ConvertToTestType(input));

            // Pass Delegate - different name and needs the signature types, but also works...
            // ParameterConversionRegistration.AddParameterConversionFunc<string, TestType>(ConvertToTestType);
        }

        public void AutoClose()
        {
        }

    }
}
