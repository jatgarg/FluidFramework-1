// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils DataProcessingError.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Signals corrupt or unprocessable data from the server / persisted state.
	/// Distinguished from <see cref="LoggingError" /> so callers can react to
	/// data-corruption cases without confusing them with other invariant
	/// failures.
	/// </summary>
	/// <remarks>
	/// TS parallel: <c>@fluidframework/telemetry-utils</c>
	/// <c>DataProcessingError</c>. Not retryable.
	/// </remarks>
	public class DataProcessingError : LoggingError
	{
		public DataProcessingError(string message)
			: base(message)
		{
		}

		public DataProcessingError(string message, IReadOnlyDictionary<string, object?>? properties)
			: base(message, properties)
		{
		}

		public DataProcessingError(string message, Exception? innerException)
			: base(message, innerException)
		{
		}

		public DataProcessingError(string message, IReadOnlyDictionary<string, object?>? properties, Exception? innerException)
			: base(message, properties, innerException)
		{
		}
	}
}
