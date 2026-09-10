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
	/// <remarks>
	/// <b>Compliance:</b> the <c>error</c> parameter must be treated as
	/// PII-suspect by default. See the per-method remarks below and the
	/// property-tagging contract on <see cref="FluidTelemetryEvent.Properties" />.
	/// Never render <see cref="Exception.ToString" /> verbatim to a shipped
	/// telemetry sink.
	/// </remarks>
	public interface IFluidLogger
	{
		/// <summary>
		/// Emit a routine telemetry event.
		/// </summary>
		/// <param name="evt">Event descriptor.</param>
		/// <param name="error">Optional exception. See remarks for
		/// compliance handling.</param>
		/// <param name="logLevel">Optional severity. When omitted, impls
		/// should treat as <see cref="FluidLogLevel.Essential" />
		/// (matches TS default).</param>
		/// <remarks>
		/// If <paramref name="error" /> is supplied, impls should split it
		/// into three compliance buckets before rendering:
		/// <list type="bullet">
		/// <item><description>Type name (e.g., <c>error.GetType().FullName</c>) —
		///   <see cref="FluidTelemetryDataTag.CodeArtifact" /> (safe).</description></item>
		/// <item><description><see cref="Exception.Message" /> —
		///   <see cref="FluidTelemetryDataTag.UserData" /> by default. Message
		///   strings may embed user content from throw-site interpolation
		///   (e.g., <c>$"unknown key: {key}"</c>). Redact / hash / drop for
		///   shipped telemetry.</description></item>
		/// <item><description>Stack trace (<see cref="Exception.StackTrace" />) —
		///   <see cref="FluidTelemetryDataTag.CodeArtifact" /> (safe), but should
		///   be sanitized to strip the leading <c>"[TypeName]: [Message]"</c>
		///   line so the message doesn't slip in via the stack. Matches TS
		///   <c>extractLogSafeErrorProperties(sanitizeStack: true)</c>.</description></item>
		/// </list>
		/// If the error implements <see cref="ILoggingError" />, its
		/// <c>GetTelemetryProperties()</c> payload is merged into the event
		/// under the same tagging contract as
		/// <see cref="FluidTelemetryEvent.Properties" />.
		/// <para />
		/// Rendering the exception via a bare <see cref="Exception.ToString" />
		/// concatenates all three parts (type, message, stack) into one
		/// blob and defeats the compliance boundary. Do not do that on the
		/// shipped-telemetry path.
		/// </remarks>
		void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null);

		/// <summary>
		/// Emit an error-severity telemetry event. TS auto-upgrades error
		/// events to <see cref="FluidLogLevel.Essential" />; no explicit
		/// level param.
		/// </summary>
		/// <remarks>
		/// Same exception-compliance handling as <see cref="SendTelemetryEvent" />
		/// applies to <paramref name="error" />.
		/// </remarks>
		void SendErrorEvent(FluidTelemetryEvent evt, Exception? error = null);

		/// <summary>
		/// Emit a performance-scoped telemetry event.
		/// </summary>
		/// <remarks>
		/// Same exception-compliance handling as <see cref="SendTelemetryEvent" />
		/// applies to <paramref name="error" />.
		/// </remarks>
		void SendPerformanceEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null);

		/// <summary>
		/// Return a child logger that prepends <paramref name="namespaceName" />
		/// to every event name, separated by
		/// <see cref="NamespacedLogger.EventNamespaceSeparator" />
		/// (<c>":"</c>). Nested calls concatenate namespaces.
		/// </summary>
		IFluidLogger CreateChildLogger(string namespaceName);
	}
}
