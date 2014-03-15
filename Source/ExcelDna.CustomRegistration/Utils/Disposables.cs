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
        bool _suppress;
        readonly CancellationTokenSource _cts;
        public CancellationDisposable(CancellationTokenSource cts)
        {
            if (cts == null)
            {
                throw new ArgumentNullException("cts");
            }

            _cts = cts;
        }

        public CancellationDisposable()
            : this(new CancellationTokenSource())
        {
        }

        public void SuppressCancel()
        {
            _suppress = true;
        }

        public CancellationToken Token
        {
            get { return _cts.Token; }
        }

        public void Dispose()
        {
            if (!_suppress) _cts.Cancel();
            _cts.Dispose();  // Not really needed...
        }
    }
}
