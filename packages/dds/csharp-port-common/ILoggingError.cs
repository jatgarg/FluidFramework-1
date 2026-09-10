// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-interfaces ILoggingError.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Contract for exceptions that expose structured properties for
	/// telemetry consumers to merge into their event payload.
	/// </summary>
	/// <remarks>
	/// TS <c>ILoggingError extends Error</c>; C# equivalence relies on
	/// implementers deriving from <see cref="System.Exception" />. Values
	/// in the returned dictionary follow the same tagging contract as
	/// <see cref="FluidTelemetryEvent.Properties" />: bare values are
	/// safe by producer contract, and PII-bearing values must be wrapped
	/// in <see cref="TaggedTelemetryValue" /> with
	/// <see cref="FluidTelemetryDataTag.UserData" />.
	/// </remarks>
	public interface ILoggingError
	{
		IReadOnlyDictionary<string, object?> GetTelemetryProperties();
	}
}
