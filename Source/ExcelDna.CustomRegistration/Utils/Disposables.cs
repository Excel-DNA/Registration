using System;
using System.Threading;

namespace ExcelDna.CustomRegistration.Utils
{
    sealed class DefaultDisposable : IDisposable
    {
        public static readonly DefaultDisposable Instance = new DefaultDisposable();

        // Prevent external instantiation
        DefaultDisposable()
        {
        }

        public void Dispose()
        {
            // no op
        }
    }

    sealed class CancellationDisposable : IDisposable
    {
        bool suppress;
        readonly CancellationTokenSource cts;
        public CancellationDisposable(CancellationTokenSource cts)
        {
            if (cts == null)
            {
                throw new ArgumentNullException("cts");
            }

            this.cts = cts;
        }

        public CancellationDisposable()
            : this(new CancellationTokenSource())
        {
        }

        public void SuppressCancel()
        {
            suppress = true;
        }

        public CancellationToken Token
        {
            get { return cts.Token; }
        }

        public void Dispose()
        {
            if (!suppress) cts.Cancel();
            cts.Dispose();  // Not really needed...
        }
    }
}
