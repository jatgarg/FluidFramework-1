// -----------------------------------------------------------------------------
// No-op IFluidLogger. Default when a host supplies no logger.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// <see cref="IFluidLogger" /> that drops every event. Functionally
	/// equivalent to disabling telemetry, with zero allocation.
	/// </summary>
	public sealed class NullLogger : IFluidLogger
	{
		public static NullLogger Instance { get; } = new();

		private NullLogger()
		{
		}

		public void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
		}

		public void SendErrorEvent(FluidTelemetryEvent evt, Exception? error = null)
		{
		}

		public void SendPerformanceEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
		}

		public IFluidLogger CreateChildLogger(string namespaceName) => this;
	}
}
