// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils ITelemetryGenericEventExt.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Descriptor consumed by <see cref="IFluidLogger" />: stable event name
	/// plus an optional structured property bag. Event names are matched
	/// verbatim with the Fluid JS side so JS/C# telemetry queries share
	/// Kusto dashboards.
	/// </summary>
	/// <remarks>
	/// A <c>readonly record struct</c> so construction on the emission path
	/// stays allocation-free. Value equality + <c>ToString()</c> from
	/// <c>record</c> are useful for tests and diagnostic sinks.
	/// <para />
	/// <see cref="Properties" /> is typed <see cref="Dictionary{TKey, TValue}" />
	/// (not <see cref="IReadOnlyDictionary{TKey, TValue}" />) so callers can
	/// use target-typed <c>new()</c> + collection-initializer syntax at the
	/// emission site: <c>Properties = new() { ["k"] = v }</c>. The read-only
	/// view remains available to consumers via
	/// <see cref="IReadOnlyDictionary{TKey, TValue}" /> access on the same
	/// dictionary reference.
	/// </remarks>
	public readonly record struct FluidTelemetryEvent
	{
		public required string EventName { get; init; }

		/// <summary>
		/// Optional structured payload. Each value is either a bare
		/// primitive / string / enum / exception (implicitly classified
		/// safe by contract — must not carry user content) or a
		/// <see cref="TaggedTelemetryValue" /> carrying an explicit
		/// <see cref="FluidTelemetryDataTag" />. Concrete
		/// <see cref="IFluidLogger" /> implementations must inspect the tag
		/// and route the value: local diagnostic sinks may render
		/// everything, shipped telemetry sinks must strip, hash, or redact
		/// <see cref="FluidTelemetryDataTag.UserData" /> values.
		/// </summary>
		public Dictionary<string, object?>? Properties { get; init; }
	}
}
