using System.Diagnostics;

namespace ExcelDna.CustomRegistration.Example
{
    public class FunctionLoggingHandler : FunctionExecutionHandler
    {
        string Tag;
        public override void OnEntry(FunctionExecutionArgs args)
        {
            Debug.Print("{0} - OnEntry", Tag);
        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {
            Debug.Print("{0} - OnSuccess", Tag);
        }

        public override void OnException(FunctionExecutionArgs args)
        {
            Debug.Print("{0} - OnException", Tag);
        }

        public override void OnExit(FunctionExecutionArgs args)
        {
            Debug.Print("{0} - OnExit", Tag);
        }

        // The configuration part - should move somewhere else.
        // Just to show we can attach arbitrary data to the captured handler.
        static int _tagIndex = 0;
        internal static FunctionExecutionHandler LoggingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            return new FunctionLoggingHandler { Tag = "Function: " + functionRegistration.FunctionAttribute.Name + ":" + _tagIndex++ };
        }
    }


}
