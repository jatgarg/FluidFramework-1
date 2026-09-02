// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils ITelemetryLoggerExt.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Sink for the port's telemetry events. Concrete implementations are
	/// responsible for serialization, sampling gates, and routing to the
	/// underlying logging backend (host-provided; never called directly by
	/// port code).
	/// </summary>
	public interface IFluidLogger
	{
		/// <summary>
		/// Emit a routine telemetry event.
		/// </summary>
		/// <param name="evt">Event descriptor.</param>
		/// <param name="error">Optional exception. Impls should render its
		/// full <c>ToString()</c> and merge <see cref="ILoggingError" />
		/// payload when present.</param>
		/// <param name="logLevel">Optional severity. When omitted, impls
		/// should treat as <see cref="FluidLogLevel.Essential" />
		/// (matches TS default).</param>
		void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null);

		/// <summary>
		/// Emit an error-severity telemetry event. TS auto-upgrades error
		/// events to <see cref="FluidLogLevel.Essential" />; no explicit
		/// level param.
		/// </summary>
		void SendErrorEvent(FluidTelemetryEvent evt, Exception? error = null);

		/// <summary>
		/// Emit a performance-scoped telemetry event.
		/// </summary>
		void SendPerformanceEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null);

		/// <summary>
		/// Return a child logger that prepends <paramref name="namespaceName" />
		/// (dot-separated) to every event name. Nested calls concatenate
		/// namespaces.
		/// </summary>
		IFluidLogger CreateChildLogger(string namespaceName);
	}
}
