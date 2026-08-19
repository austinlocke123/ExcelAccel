using System;
using System.Threading;

namespace ExcelAccel.Core.Reliability;

public sealed class ReentrancyGate
{
    private int _entered;

    public IDisposable? TryEnter()
    {
        if (Interlocked.CompareExchange(ref _entered, 1, 0) != 0)
        {
            return null;
        }

        return new Lease(this);
    }

    private sealed class Lease : IDisposable
    {
        private ReentrancyGate? _owner;

        public Lease(ReentrancyGate owner)
        {
            _owner = owner;
        }

        public void Dispose()
        {
            var owner = Interlocked.Exchange(ref _owner, null);
            if (owner is not null)
            {
                Volatile.Write(ref owner._entered, 0);
            }
        }
    }
}
