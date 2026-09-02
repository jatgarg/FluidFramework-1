// -----------------------------------------------------------------------------
// Test-only IFluidLogger that captures events for assertion.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

using Microsoft.Office.Web.Fluid;

namespace Microsoft.Office.Web.Fluid.Tests
{
	internal sealed class RecordingLogger : IFluidLogger
	{
		public List<RecordedEvent> Telemetry { get; } = new();
		public List<RecordedEvent> Errors { get; } = new();
		public List<RecordedEvent> Performance { get; } = new();

		public void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
			Telemetry.Add(new RecordedEvent(evt.EventName, evt.Properties, error, logLevel));
		}

		public void SendErrorEvent(FluidTelemetryEvent evt, Exception? error = null)
		{
			Errors.Add(new RecordedEvent(evt.EventName, evt.Properties, error, LogLevel: null));
		}

		public void SendPerformanceEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
			Performance.Add(new RecordedEvent(evt.EventName, evt.Properties, error, logLevel));
		}

		public IFluidLogger CreateChildLogger(string namespaceName)
		{
			return NamespacedLogger.Create(this, namespaceName);
		}

		internal sealed record RecordedEvent(
			string EventName,
			IReadOnlyDictionary<string, object?>? Properties,
			Exception? Error,
			FluidLogLevel? LogLevel);
	}
}
