using System;

namespace ExcelDna.CustomRegistration
{
    public enum FlowBehavior
    {
        /// <summary>
        /// Default behaviour - Same as continue for OnEnter, OnSuccess and OnExit; same as RethrowException for OnException.
        /// </summary>
        Default = 0,
        // Makes no sense to me yet.
        ///// <summary>
        ///// Continue normally - For an OnException handler would suppress the exception and continue as if the method ware successful
        ///// </summary>
        //Continue = 1,
        /// <summary>
        /// Rethrow the current exception - only valid for OnException handlers.
        /// </summary>
        RethrowException = 2,
        /// Return the value of ReturnValue immediately  - For OnEnter will skip the method execution and the OnSuccess handlers, but will run OnExit handlers
        Return = 3,
        /// <summary>
        /// Throw the Exception in the Exception property - For OnException handlers only.
        /// </summary>
        ThrowException = 4
    }

    // CONSIDER: One might make a generic typed version of this...
    public class MethodExecutionArgs
    {
        // Can't change arguments - Make ReadOnly collection?
        public object[] Arguments { get; private set; }
        public object ReturnValue { get; set; }
        // Can't change exception
        public Exception Exception { get; set; }
        public FlowBehavior FlowBehavior { get; set; }
        // public Method ...?
        public object Tag { get; set; }

        public MethodExecutionArgs(object[] arguments)
        {
            Arguments = arguments;
        }
    }

    /*
        public static int MyMethod(object arg0, int arg1)
        {
          OnEntry();
          try
          {
            // Original method body. 
            OnSuccess();
            return returnValue;
          }
          catch ( Exception e )
          {
            OnException();
          }
          finally
          {
            OnExit();
          }
        }
    */
    public abstract class MethodExecutionHandler
    {
        public abstract void OnEntry(MethodExecutionArgs args);
        public abstract void OnSuccess(MethodExecutionArgs args);
        public abstract void OnException(MethodExecutionArgs args);
        public abstract void OnExit(MethodExecutionArgs args);
    }
}
