using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration.Example
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Registration.GetExcelFunctions()
                        .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                        .ProcessParamsRegistrations()
                        .RegisterFunctions();
        }

        public void AutoClose()
        {
        }

    }
}
