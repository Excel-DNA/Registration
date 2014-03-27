using System.Diagnostics;

namespace ExcelDna.CustomRegistration.Example
{
    public class MethodLoggingHandler : MethodExecutionHandler
    {
        string Tag;
        public override void OnEntry(MethodExecutionArgs args)
        {
            Debug.Print("{0} - OnEntry", Tag);
        }

        public override void OnSuccess(MethodExecutionArgs args)
        {
            Debug.Print("{0} - OnSuccess", Tag);
        }

        public override void OnException(MethodExecutionArgs args)
        {
            Debug.Print("{0} - OnException", Tag);
        }

        public override void OnExit(MethodExecutionArgs args)
        {
            Debug.Print("{0} - OnExit", Tag);
        }

        // The configuration part - should move somewhere else.
        // Just to show we can attach arbitrary data to the captured handler.
        static int _tagIndex = 0;
        internal static MethodExecutionHandler LoggingHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            return new MethodLoggingHandler { Tag = "Function: " + functionRegistration.FunctionAttribute.Name + ":" + _tagIndex++ };
        }
    }


}
