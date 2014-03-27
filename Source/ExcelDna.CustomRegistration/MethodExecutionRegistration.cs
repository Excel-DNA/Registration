using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Expr = System.Linq.Expressions.Expression;

namespace ExcelDna.CustomRegistration
{
    public static class MethodExecutionRegistration
    {
        public static IEnumerable<ExcelFunctionRegistration> ProcessMethodExecutionHandlers(this IEnumerable<ExcelFunctionRegistration> registrations, MethodExecutionConfiguration methodHandlerConfig)
        {
            foreach (var registration in registrations)
            {
                var reg = registration; // Ensure safe semantics for captured foreach variable
                var handlers = methodHandlerConfig.MethodHandlerSelectors
                                                  .Select(mhSelector => mhSelector(reg))
                                                  .Where(mh => mh != null);
                ApplyMethodHandlers(reg, handlers);

                yield return reg;
            }
        }

        static void ApplyMethodHandlers(ExcelFunctionRegistration reg, IEnumerable<MethodExecutionHandler> handlers)
        {
            // The order of method handlers is important - we follow PostSharp's convention.
            // The are passed from high priority (most inside) to low priority (most outside)
            // Imagine 2 MethodHandlers, mh1 then mh2
            // So mh1 (highest priority)  will be 'inside' and mh2 will be outside (lower priority)
            foreach (var handler in handlers)
            {
                reg.FunctionLambda = ApplyMethodHandler(reg.FunctionLambda, handler);
            }
        }
        
        static LambdaExpression ApplyMethodHandler(LambdaExpression functionLambda, MethodExecutionHandler handler)
        {
            //  public static int MyMethod(object arg0, int arg1) { ... }
             
            // becomes:

            // (the 'handler' object is captured and called mh)
            //public static int MyMethodWrapped(object arg0, int arg1)
            //{
            //    var mhArgs = new MethodExecutionArgs(new object[] { arg0, arg1});
            //    int result = default(int);
            //    try
            //    {
            //        mh.OnEntry(mhArgs);
            //        if (mhArgs.FlowBehavior == FlowBehavior.Return)
            //        {
            //            result = (int)mhArgs.ReturnValue;
            //        }
            //        else
            //        {
            //             // Inner call
            //             result = MyMethod(arg0, arg1);
            //             mhArgs.ReturnValue = result;
            //             mh.OnSuccess(mhArgs);
            //             result = (int)mhArgs.ReturnValue;
            //        }
            //    }
            //    catch ( Exception ex )
            //    {
            //        mhArgs.Exception = ex;
            //        mh.OnException(mhArgs);
            //        // Makes no sense to me yet - I've removed this FlowBehavior enum value.
            //        // if (mhArgs.FlowBehavior == FlowBehavior.Continue)
            //        // {
            //        //     // Finally will run, but can't change return value
            //        //     // Should we assign result...?
            //        //     // So Default value will be returned....?????
            //        //     mhArgs.Exception = null;
            //        // }
            //        // else 
            //        if (mhArgs.FlowBehavior == FlowBehavior.Return)
            //        {
            //            // Clear the Exception and return the ReturnValue instead
            //            // Finally will run, but can't further change return value
            //            mhArgs.Exception = null;
            //            result = (int)mhArgs.ReturnValue;
            //        }
            //        else if (mhArgs.FlowBehavior == FlowBehavior.ThrowException)
            //        {
            //            throw mhArgs.Exception;
            //        }
            //        else // if (mhArgs.FlowBehavior == FlowBehavior.Default || mhArgs.FlowBehavior == FlowBehavior.RethrowException)
            //        {
            //            throw;
            //        }
            //    }
            //    finally
            //    {
            //        mh.OnExit(mhArgs);
            //        // NOTE: mhArgs.ReturnValue is not used again here...!
            //    }
            //    
            //    return result;
            //  }
            //}

            // Ensure the handler object is captured.
            var mh = Expression.Constant(handler);

            // Prepare the methodHandlerArgs that will be threaded through the handler, 
            // and a bunch of expressions that access various properties on it.
            var mhArgs = Expr.Variable(typeof(MethodExecutionArgs), "mhArgs");
            var mhArgsReturnValue = SymbolExtensions.GetProperty(mhArgs, (MethodExecutionArgs mea) => mea.ReturnValue);
            var mhArgsException = SymbolExtensions.GetProperty(mhArgs, (MethodExecutionArgs mea) => mea.Exception);
            var mhArgsFlowBehaviour = SymbolExtensions.GetProperty(mhArgs, (MethodExecutionArgs mea) => mea.FlowBehavior);

            // Set up expressions to call the various handler methods.
            // TODO: Later we can determine which of these are actually implemented, and only write out the code needed in the particular case.
            var onEntry = Expr.Call(mh, SymbolExtensions.GetMethodInfo<MethodExecutionHandler>(meh => meh.OnEntry(null)), mhArgs);
            var onSuccess = Expr.Call(mh, SymbolExtensions.GetMethodInfo<MethodExecutionHandler>(meh => meh.OnSuccess(null)), mhArgs);
            var onException = Expr.Call(mh, SymbolExtensions.GetMethodInfo<MethodExecutionHandler>(meh => meh.OnException(null)), mhArgs);
            var onExit = Expr.Call(mh, SymbolExtensions.GetMethodInfo<MethodExecutionHandler>(meh => meh.OnExit(null)), mhArgs);

            // Create the array of parameter values that will be put into the method handler args.
            var paramsArray = Expr.NewArrayInit(typeof(object), functionLambda.Parameters.Select(p => Expr.Convert(p, typeof(object))));

            // Prepare the result and ex(ception) local variables
            var result = Expr.Variable(functionLambda.ReturnType, "result");
            var ex = Expression.Parameter(typeof(Exception), "ex");

            // A bunch of helper expressions:
            // : new MethodExecutionArgs(new object[] { arg0, arg1 })
            var mhArgsConstr = typeof(MethodExecutionArgs).GetConstructor(new[] { typeof(object[]) });
            var newMhArgs = Expr.New(mhArgsConstr, paramsArray);
            // : result = (int)mhArgs.ReturnValue
            var resultFromReturnValue = Expr.Assign(result, Expr.Convert(mhArgsReturnValue, functionLambda.ReturnType));
            // : mhArgs.ReturnValue = (object)result
            var returnValueFromResult = Expr.Assign(mhArgsReturnValue, Expr.Convert(result, typeof(object)));
            // : result = function(arg0, arg1)
            var resultFromInnerCall = Expr.Assign(result, Expr.Invoke(functionLambda, functionLambda.Parameters));

            // Build the Lambda wrapper, with the original parameters
            return Expr.Lambda(
                Expr.Block(new[] { mhArgs, result },
                     Expr.Assign(mhArgs, newMhArgs),
                     Expr.Assign(result, Expr.Default(result.Type)),
                     Expr.TryCatchFinally(
                        Expr.Block( 
                            onEntry,
                            Expr.IfThenElse(
                                Expr.Equal(mhArgsFlowBehaviour, Expr.Constant(FlowBehavior.Return)),
                                resultFromReturnValue,
                                Expr.Block(
                                    resultFromInnerCall, 
                                    returnValueFromResult,
                                    onSuccess,
                                    resultFromReturnValue))),
                        onExit, // finally
                        Expr.Catch(ex,
                            Expr.Block(
                                Expr.Assign(mhArgsException, ex),
                                onException,
                                Expr.IfThenElse(
                                    Expr.Equal(mhArgsFlowBehaviour, Expr.Constant(FlowBehavior.Return)),
                                    Expr.Block(
                                        Expr.Assign(mhArgsException, Expr.Constant(null, typeof(Exception))),
                                        resultFromReturnValue),
                                    Expr.IfThenElse(
                                        Expr.Equal(mhArgsFlowBehaviour, Expr.Constant(FlowBehavior.ThrowException)),
                                        Expr.Throw(mhArgsException),
                                        Expr.Rethrow()))))
                        ),
                    result),
                functionLambda.Parameters);
        }
            
        

    }
}
