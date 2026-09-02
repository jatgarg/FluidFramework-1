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
	/// implementers deriving from <see cref="System.Exception" />.
	/// Implementations must not return PII, credentials, or user content.
	/// </remarks>
	public interface ILoggingError
	{
		IReadOnlyDictionary<string, object?> GetTelemetryProperties();
	}
}
