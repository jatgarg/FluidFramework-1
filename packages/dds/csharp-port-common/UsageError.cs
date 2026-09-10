// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils UsageError.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Public-API contract violation raised by port callers (bad arguments,
	/// illegal state transitions, unsupported input shapes). Downstream
	/// telemetry filters can key off the <c>usageError: true</c> property.
	/// </summary>
	public class UsageError : LoggingError
	{
		public UsageError(string message)
			: base(message, PropsWithUsageFlag(null))
		{
		}

		public UsageError(string message, Dictionary<string, object?>? properties)
			: base(message, PropsWithUsageFlag(properties))
		{
		}

		public UsageError(string message, Exception? innerException)
			: base(message, PropsWithUsageFlag(null), innerException)
		{
		}

		public UsageError(string message, Dictionary<string, object?>? properties, Exception? innerException)
			: base(message, PropsWithUsageFlag(properties), innerException)
		{
		}

		private static Dictionary<string, object?> PropsWithUsageFlag(Dictionary<string, object?>? properties)
		{
			Dictionary<string, object?> merged = properties is null
				? new Dictionary<string, object?>(1)
				: new Dictionary<string, object?>(properties);
			merged["usageError"] = true;
			return merged;
		}
	}
}
