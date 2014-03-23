using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace ExcelDna.CustomRegistration
{
    public static class MethodExecutionRegistration
    {
        public static IEnumerable<ExcelFunctionRegistration> ProcessMethodHandlers(this IEnumerable<ExcelFunctionRegistration> registrations, MethodExecutionConfiguration methodHandlerConfig)
        {
            foreach (var reg in registrations)
            {
                var registration = reg; // Safe semantics for captures foreach variable
                var handlers = methodHandlerConfig.MethodHandlers
                                                  .Select(mhSelector => mhSelector(registration))
                                                  .Where(mh => mh != null);
                ApplyMethodHandlers(reg, handlers);

                yield return reg;
            }
        }

        static void ApplyMethodHandlers(ExcelFunctionRegistration reg, IEnumerable<MethodExecutionHandler> handlers)
        {
            // The order of method handlers is important.
            // The are passed from high priority (most inside) to low priority (most outside)
            // Imagine 2 MethodHandlers, mh1 then mh2
            // So mh1 will be 'inside' (highest priority) and mh2 will be outside (lower priority)
        }
            
            //  public static int MyMethod(object arg0, int arg1) { return 0; }
             
            // becomes:
             
                //public static int MyMethodWrap(object arg0, int arg1)
                //{
                //    MethodExecutionHandler mh1 = null;
                //    MethodExecutionHandler  mh2 = null;

                //    var mh2Args = new MethodExecutionArgs(new object[] { arg0, arg1});
                //    int result = default(int);
                //    
                //    try
                //    {
                //        mh2.OnEntry(mh2Args);
                //        if (mh2Args.FlowBehavior == FlowBehavior.Return)
                //        {
                //            result = (int)mh2Args.ReturnValue;
                //        }
                //        else
                //        {
                               // Inner call
                               
                //                 // var mh1Args = ...                            
                //                 // OnEntry...
                //                     result = MyMethod(arg0, arg1);
                //             mh2Args.ReturnValue = result;
                //             mh2.OnSuccess(mh2Args);
                //             result = (int)mh2Args.ReturnValue;
                //        }
                //    }
                //    catch ( Exception ex )
                //    {
                //        mh2Args.Exception = ex;
                //        mh2.OnException(mh2Args);
                //        // Makes no sense to me yet.
                //        // if (mh2Args.FlowBehavior == FlowBehavior.Continue)
                //        // {
        // TODO: .......?????????????
                //        //     // Finally will run, but can't change return value
                //        //     // So Default value will be returned....?????
                //        //     mh2Args.Exception = null;
                //        // }
                //        // else 
                //        if (mh2Args.FlowBehavior == FlowBehavior.Return)
                //        {
                //            // Finally will run, but can't change return value
                //            mh2Args.Exception = null;
                //            result = (int)mh2Args.ReturnValue;
                //        }
                //        else if (mh2Args.FlowBehavior == FlowBehavior.ThrowException)
                //        {
                //            throw mh2Args.Exception;
                //        }
                //        else // if (mhArgs.FlowBehavior == FlowBehavior.Default || mhArgs.FlowBehavior == FlowBehavior.RethrowException)
                //        {
                //            throw;
                //        }
                //    }
                //    finally
                //    {
                //        mh2Args.ReturnValue = result;
                //        mh2.OnExit(mh2Args);
                //        // mh2Args.ReturnValue is not used here...!
                //    }
                //    
                //    return result;
                //  }
                //}

    }
}
