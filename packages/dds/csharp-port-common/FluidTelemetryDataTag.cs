// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils TelemetryDataTag.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Compliance tag for a telemetry property value. Every value on
	/// <see cref="FluidTelemetryEvent.Properties" /> and
	/// <see cref="ILoggingError.GetTelemetryProperties" /> is either wrapped
	/// in a <see cref="TaggedTelemetryValue" /> (bearing one of these tags)
	/// or bare (implicitly <see cref="CodeArtifact" /> — safe by contract).
	/// </summary>
	/// <remarks>
	/// Concrete <see cref="IFluidLogger" /> implementations must inspect the
	/// tag and route the value appropriately: local diagnostic sinks may
	/// render every value, shipped telemetry sinks must strip or redact
	/// <see cref="UserData" />.
	/// </remarks>
	public enum FluidTelemetryDataTag
	{
		/// <summary>
		/// Data containing terms or IDs from code — event names, op types,
		/// segment counts, sequence numbers. Safe to log to shipped telemetry.
		/// </summary>
		CodeArtifact,

		/// <summary>
		/// Personal data pertaining to the user (segment text, property
		/// values, marker text, etc.). Local logs may print; shipped
		/// telemetry sinks must strip, hash, or redact.
		/// </summary>
		UserData,
	}
}
