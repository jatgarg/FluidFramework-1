// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-interfaces Tagged<V, T>.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// A telemetry-property value wrapped with a
	/// <see cref="FluidTelemetryDataTag" />. Appears alongside bare
	/// values inside <see cref="FluidTelemetryEvent.Properties" /> and
	/// <see cref="ILoggingError.GetTelemetryProperties" />.
	/// </summary>
	/// <remarks>
	/// TS parallel: <c>Tagged&lt;V, T&gt;</c>. Producers wrap values
	/// whose class must be declared (e.g., user data that a wire sink
	/// should strip); consumers inspect <see cref="Tag" /> and route.
	/// </remarks>
	public sealed record TaggedTelemetryValue(object? Value, FluidTelemetryDataTag Tag);
}
