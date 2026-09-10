// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-interfaces Tagged<V, T>.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// A telemetry-property value wrapped with a <see cref="FluidTelemetryDataTag" />.
	/// Appears alongside bare primitive values inside
	/// <see cref="FluidTelemetryEvent.Properties" /> and
	/// <see cref="ILoggingError.GetTelemetryProperties" />.
	/// </summary>
	/// <remarks>
	/// TS parallel: <c>Tagged&lt;V, T&gt;</c> with fields <c>value</c> +
	/// <c>tag</c>. The C# equivalent uses PascalCase properties but keeps
	/// the same shape. Producers construct these when they attach values
	/// that might carry PII or need explicit compliance classification;
	/// consumers (concrete <see cref="IFluidLogger" /> impls) inspect the
	/// tag and route the value accordingly.
	/// </remarks>
	public sealed record TaggedTelemetryValue(object? Value, FluidTelemetryDataTag Tag);
}
