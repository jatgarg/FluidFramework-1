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
	public sealed class FluidTelemetryEvent
	{
		public required string EventName { get; init; }

		/// <summary>
		/// Optional structured payload. Values should be primitives, strings,
		/// enums (via <c>ToString()</c>), or exceptions (via
		/// <c>ToString()</c>). Nested dicts / arrays are logger-specific.
		/// </summary>
		public IReadOnlyDictionary<string, object?>? Properties { get; init; }
	}
}
