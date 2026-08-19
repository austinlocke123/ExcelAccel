using System;
using ExcelAccel.Core.Reliability;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ApplicationStateGuardTests
{
    [Fact]
    public void RestoresOwnedStateAfterSuccessfulMutation()
    {
        var port = new FakeApplicationStatePort();

        ApplicationStateGuard.Run(port, ApplicationStateChangeSet.PropertyMutation(), () =>
        {
            Assert.False(port.ScreenUpdating);
            Assert.False(port.EnableEvents);
        });

        Assert.True(port.ScreenUpdating);
        Assert.True(port.EnableEvents);
    }

    [Fact]
    public void RestoresOwnedStateAndRethrowsMutationFailure()
    {
        var port = new FakeApplicationStatePort();
        var expected = new InvalidOperationException("injected mutation failure");

        var actual = Assert.Throws<InvalidOperationException>(() =>
            ApplicationStateGuard.Run(
                port,
                ApplicationStateChangeSet.PropertyMutation(),
                () => throw expected));

        Assert.Same(expected, actual);
        Assert.True(port.ScreenUpdating);
        Assert.True(port.EnableEvents);
    }

    [Fact]
    public void RestorationFailureCannotAppearAsSuccess()
    {
        var port = new FakeApplicationStatePort
        {
            SetFailure = (property, value) =>
                property == "ScreenUpdating" && value
                    ? new InvalidOperationException("injected restore failure")
                    : null,
        };

        var exception = Assert.Throws<StateRestoreException>(() =>
            ApplicationStateGuard.Run(port, ApplicationStateChangeSet.PropertyMutation(), () => { }));

        Assert.Contains("ScreenUpdating", exception.FailedProperties);
        Assert.True(port.EnableEvents);
        Assert.False(port.ScreenUpdating);
    }

    [Fact]
    public void GuardEntryFailureRestoresPreviouslyOwnedState()
    {
        var port = new FakeApplicationStatePort
        {
            SetFailure = (property, value) =>
                property == "ScreenUpdating" && !value
                    ? new InvalidOperationException("injected entry failure")
                    : null,
        };

        Assert.Throws<InvalidOperationException>(() =>
            ApplicationStateGuard.Enter(port, ApplicationStateChangeSet.PropertyMutation()));

        Assert.True(port.ScreenUpdating);
        Assert.True(port.EnableEvents);
    }

    [Fact]
    public void AlreadySuppressedExternalStateIsNotClaimedOrChanged()
    {
        var port = new FakeApplicationStatePort
        {
            ScreenUpdating = false,
            EnableEvents = false,
        };
        port.ResetSetCount();

        ApplicationStateGuard.Run(port, ApplicationStateChangeSet.PropertyMutation(), () => { });

        Assert.False(port.ScreenUpdating);
        Assert.False(port.EnableEvents);
        Assert.Equal(0, port.SetCount);
    }

    private sealed class FakeApplicationStatePort : IApplicationStatePort
    {
        private bool _screenUpdating = true;
        private bool _enableEvents = true;

        public Func<string, bool, Exception?>? SetFailure { get; set; }

        public int SetCount { get; private set; }

        public void ResetSetCount() => SetCount = 0;

        public bool ScreenUpdating
        {
            get => _screenUpdating;
            set
            {
                Set("ScreenUpdating", value);
                _screenUpdating = value;
            }
        }

        public bool EnableEvents
        {
            get => _enableEvents;
            set
            {
                Set("EnableEvents", value);
                _enableEvents = value;
            }
        }

        private void Set(string property, bool value)
        {
            SetCount++;
            var failure = SetFailure?.Invoke(property, value);
            if (failure is not null)
            {
                throw failure;
            }
        }
    }
}
