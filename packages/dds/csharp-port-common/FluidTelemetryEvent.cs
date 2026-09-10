// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils ITelemetryGenericEventExt.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Descriptor consumed by <see cref="IFluidLogger" />: stable event
	/// name plus an optional structured property bag. Event names are
	/// matched verbatim with the Fluid JS side so JS/C# telemetry queries
	/// share Kusto dashboards.
	/// </summary>
	/// <remarks>
	/// A <c>readonly record struct</c> so construction on the emission
	/// path stays allocation-free.
	/// <para />
	/// <see cref="Properties" /> is a mutable <see cref="Dictionary{TKey, TValue}" />
	/// (not <see cref="IReadOnlyDictionary{TKey, TValue}" />) so callers
	/// can use target-typed <c>new()</c> + collection initializer:
	/// <c>Properties = new() { ["k"] = v }</c>.
	/// </remarks>
	public readonly record struct FluidTelemetryEvent
	{
		public required string EventName { get; init; }

		/// <summary>
		/// Optional structured payload. Each value is either bare —
		/// safe by producer contract (must not carry user content) — or
		/// wrapped in a <see cref="TaggedTelemetryValue" /> declaring
		/// its <see cref="FluidTelemetryDataTag" />. Concrete
		/// <see cref="IFluidLogger" /> implementations must strip / hash
		/// / redact <see cref="FluidTelemetryDataTag.UserData" /> values
		/// on shipped telemetry sinks.
		/// </summary>
		public Dictionary<string, object?>? Properties { get; init; }
	}
}
