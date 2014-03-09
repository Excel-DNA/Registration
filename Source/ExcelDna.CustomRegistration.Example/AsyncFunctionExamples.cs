using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.CustomRegistration.Utils;
using ExcelDna.Integration;

namespace ExcelDna.CustomRegistration.Example
{
    public static class AsyncFunctionExamples
    {
        // Will be registered in Excel by Excel-DNA, without being picked up by our CustomRegistration processing
        // since there is no ExcelFunction attribute.
        // Adding ExplicitRegistration="true" in the .dna file would prevent this function from being registered.
        public static string dnaSayHello(string name)
        {
            return "Hello " + name + "!";
        }

        // A simple function that can take a long time to complete.
        public static string dnaDelayedHello(string name, int msToSleep)
        {
            Thread.Sleep(msToSleep);
            return "Hello " + name + "!";
        }

        // Explicitly marked with ExcelAsyncFunction, so it will be wrapped by CustomRegistration
        // If we marked this function with [ExcelFunction] instead of [ExcelAsyncFunction] it would
        // not be wrapped (since it doesn't return Task or IObservable).
        [ExcelAsyncFunction(Name="dnaDelayedHelloAsync", Description="A friendly async function")]
        public static string dnaDelayedHello2(string name, int msToSleep)
        {
            Thread.Sleep(msToSleep);
            return "Hello " + name + "!";
        }

        // A function that returns a Task<T> and will be wrapped by CustomRegistration
        // It doesn't matter if this function is marked with ExcelFunction or ExcelAsyncFunction
        [ExcelFunction]
        public static Task<string> dnaDelayedTaskHello(string name, int msDelay)
        {
            return Task.Factory.StartNew(() => Delay(msDelay).ContinueWith(_ => "Hello" + name)).Unwrap();
            // With .NET 4.5 one would do:
            // return Task.Run(() => Task.Delay(msDelay).ContinueWith(_ => "Hello" + name));
        }

        // .NET 4.5 function with async/await
        // Change the Example project's target runtime and uncomment
        //[ExcelAsyncFunction]
        //public static async Task<string> dnaDelayedTaskHello(string name, int msDelay)
        //{
        //    await Task.Delay(msDelay);
        //    return "Hello " + name;
        //}

        // A function that returns a Task<T>, takes a CancellationToken as last parameter, and will be wrapped by CustomRegistration
        // It doesn't matter if this function is marked with ExcelFunction or ExcelAsyncFunction.
        // Whether the registration uses the native async under Excel 2010+ will make a big difference to the cancellation here!
        [ExcelAsyncFunction]
        public static Task<string> dnaDelayedTaskHelloWithCancellation(string name, int msDelay, CancellationToken ct)
        {
            ct.Register(() => Debug.Print("Cancelled!"));

            return Task.Factory.StartNew(() =>
            {
                Debug.Print("Started calc!");
                return Delay(msDelay).ContinueWith(_ => "Hello" + name);
            }).Unwrap();

            // With .NET 4.5 one could do the same a bit simpler:
            // return Task.Run(() => Task.Delay(msDelay).ContinueWith(_ => "Hello" + name));
        }

        // This is what the Task wrappers that is generated looks like.
        // Can use the same Task helper here.
        public static object dnaExplicitWrap(string name, int msDelay)
        {
            return AsyncTaskUtil.RunTask("dnaExplicitWrap", new object[] { name, msDelay }, () => dnaDelayedTaskHello(name, msDelay));
        }

        // private function used here to create a 'Delay' Task, but built-in under .NET 4.5
        static Task Delay(int milliseconds)
        {
            var tcs = new TaskCompletionSource<object>();
            new Timer(_ => tcs.SetResult(null)).Change(milliseconds, Timeout.Infinite);
            return tcs.Task;
        }
    }
}
