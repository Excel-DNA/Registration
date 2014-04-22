using System.Linq;
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
            var conversionConfig = GetParameterConversionConfig();

            var functionHandlerConfig = GetFunctionExecutionHandlerConfig();

            // Get all the ExcelFunction functions, process and register
            // Since the .dna file has ExplicitExports="true", these explicit regisrations are the only ones - there is no default processing
            Registration.GetExcelFunctions()
                        .ProcessParameterConversions(conversionConfig)
                        .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                        .ProcessParamsRegistrations()
                        .ProcessFunctionExecutionHandlers(functionHandlerConfig)
                        .RegisterFunctions();
        }

        static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            return new ParameterConversionConfiguration()
                // CONSIDER: This might have to change if we want to add improved tracing to the conversions.
                // TODO: Parameter vs Return conversions...?

            // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))

            // Some ideas ways to define and register conversions
                // These are for a particular parameter type
                // (Func<object, MyType> would allow MyType to be taken as parameter)

            // Inline Lambda - one way
                .AddParameterConversion((string value) => new TestType1(value))
                .AddParameterConversion((TestType1 value) => new TestType2(value))

                .AddReturnConversion((TestType1 value) => value.ToString())
            //  .AddParameterConversion((string value) => convert2(convert1(value)));

            // Alternative - use method via lambda
                // This adds a conversion to allow string[] parameters (by accepting object[] instead).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray());

            // Pass Delegate - different name and needs the signature types, but also works...
            //  .AddParameterConversionFunc<string, TestType>(ConvertToTestType);
        }

        static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
        {
            return new FunctionExecutionConfiguration()
                .AddFunctionExecutionHandler(FunctionLoggingHandler.LoggingHandlerSelector)
                .AddFunctionExecutionHandler(CacheFunctionExecutionHandler.CacheHandlerSelector)
                .AddFunctionExecutionHandler(TimingFunctionExecutionHandler.TimingHandlerSelector)
                .AddFunctionExecutionHandler(SuppressInDialogFunctionExecutionHandler.SuppressInDialogSelector);
        }

        public void AutoClose()
        {
        }

    }
}
