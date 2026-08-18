using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;

namespace ExcelAccel.Core.Reliability;

public sealed class ApplicationStateGuard
{
    private readonly IApplicationStatePort _port;
    private readonly List<OwnedBooleanState> _ownedStates = new List<OwnedBooleanState>();
    private bool _restored;

    private ApplicationStateGuard(IApplicationStatePort port)
    {
        _port = port;
    }

    public static ApplicationStateGuard Enter(IApplicationStatePort port, ApplicationStateChangeSet changes)
    {
        if (port is null)
        {
            throw new ArgumentNullException(nameof(port));
        }

        if (changes is null)
        {
            throw new ArgumentNullException(nameof(changes));
        }

        var guard = new ApplicationStateGuard(port);
        try
        {
            if (changes.SuppressEvents)
            {
                guard.ChangeBoolean("EnableEvents", () => port.EnableEvents, value => port.EnableEvents = value, false);
            }

            if (changes.SuppressScreenUpdating)
            {
                guard.ChangeBoolean("ScreenUpdating", () => port.ScreenUpdating, value => port.ScreenUpdating = value, false);
            }

            return guard;
        }
        catch (Exception enterFailure)
        {
            var restoreFailures = guard.Restore();
            if (restoreFailures.Count != 0)
            {
                throw new StateRestoreException(
                    "Excel application state could not be restored after guard entry failed.",
                    restoreFailures,
                    enterFailure);
            }

            throw;
        }
    }

    public IReadOnlyList<string> Restore()
    {
        if (_restored)
        {
            return Array.Empty<string>();
        }

        _restored = true;
        var failures = new List<string>();
        for (var index = _ownedStates.Count - 1; index >= 0; index--)
        {
            var state = _ownedStates[index];
            try
            {
                var current = state.Read();
                if (current == state.OriginalValue)
                {
                    continue;
                }

                if (current != state.OwnedValue)
                {
                    failures.Add(state.Name);
                    continue;
                }

                state.Write(state.OriginalValue);
            }
            catch (Exception)
            {
                failures.Add(state.Name);
            }
        }

        return failures;
    }

    public static void Run(
        IApplicationStatePort port,
        ApplicationStateChangeSet changes,
        Action mutation)
    {
        if (mutation is null)
        {
            throw new ArgumentNullException(nameof(mutation));
        }

        var guard = Enter(port, changes);
        Exception? mutationFailure = null;
        try
        {
            mutation();
        }
        catch (Exception exception)
        {
            mutationFailure = exception;
        }

        var restoreFailures = guard.Restore();
        if (restoreFailures.Count != 0)
        {
            throw new StateRestoreException(
                "Excel application state could not be fully restored.",
                restoreFailures,
                mutationFailure);
        }

        if (mutationFailure is not null)
        {
            ExceptionDispatchInfo.Capture(mutationFailure).Throw();
        }
    }

    private void ChangeBoolean(string name, Func<bool> read, Action<bool> write, bool desiredValue)
    {
        var original = read();
        if (original == desiredValue)
        {
            return;
        }

        try
        {
            write(desiredValue);
        }
        catch (Exception writeFailure)
        {
            try
            {
                if (read() == desiredValue)
                {
                    _ownedStates.Add(new OwnedBooleanState(name, original, desiredValue, read, write));
                }
            }
            catch (Exception stateUnknown)
            {
                throw new StateRestoreException(
                    $"The '{name}' state is unknown after a failed state change.",
                    new[] { name },
                    stateUnknown);
            }

            ExceptionDispatchInfo.Capture(writeFailure).Throw();
        }

        _ownedStates.Add(new OwnedBooleanState(name, original, desiredValue, read, write));
    }

    private sealed class OwnedBooleanState
    {
        public OwnedBooleanState(string name, bool originalValue, bool ownedValue, Func<bool> read, Action<bool> write)
        {
            Name = name;
            OriginalValue = originalValue;
            OwnedValue = ownedValue;
            Read = read;
            Write = write;
        }

        public string Name { get; }

        public bool OriginalValue { get; }

        public bool OwnedValue { get; }

        public Func<bool> Read { get; }

        public Action<bool> Write { get; }
    }
}
