// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils LoggingError.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Base exception for port-side invariant / assert failures. Carries a
	/// structured payload for consumers of <see cref="ILoggingError" />.
	/// </summary>
	/// <remarks>
	/// Prefer this over <see cref="InvalidOperationException" /> whenever
	/// the throw represents an invariant the port itself enforces
	/// (<see cref="UsageError" /> is the sibling for caller-fault throws).
	/// </remarks>
	public class LoggingError : Exception, ILoggingError
	{
		private readonly IReadOnlyDictionary<string, object?> _properties;

		public LoggingError(string message)
			: this(message, properties: null, innerException: null)
		{
		}

		public LoggingError(string message, IReadOnlyDictionary<string, object?>? properties)
			: this(message, properties, innerException: null)
		{
		}

		public LoggingError(string message, Exception? innerException)
			: this(message, properties: null, innerException)
		{
		}

		public LoggingError(string message, IReadOnlyDictionary<string, object?>? properties, Exception? innerException)
			: base(message, innerException)
		{
			_properties = properties ?? EmptyProperties;
			ErrorInstanceId = Guid.NewGuid().ToString("N");
		}

		/// <summary>
		/// Unique instance id used to correlate multiple log lines that
		/// reference the same error occurrence. TS parallel:
		/// <c>errorInstanceId</c> on <c>LoggingError</c>.
		/// </summary>
		public string ErrorInstanceId { get; private set; }

		/// <summary>
		/// Overwrite the instance id (used when wrapping an inner error to
		/// preserve its id). TS parallel: <c>overwriteErrorInstanceId</c>.
		/// </summary>
		public void OverwriteErrorInstanceId(string id)
		{
			if (string.IsNullOrEmpty(id))
			{
				throw new ArgumentException("ErrorInstanceId must not be empty.", nameof(id));
			}

			ErrorInstanceId = id;
		}

		public IReadOnlyDictionary<string, object?> GetTelemetryProperties()
		{
			Dictionary<string, object?> result = new(_properties.Count + 3)
			{
				["message"] = Message,
				["stack"] = StackTrace,
				["errorInstanceId"] = ErrorInstanceId,
			};

			foreach (KeyValuePair<string, object?> pair in _properties)
			{
				result[pair.Key] = pair.Value;
			}

			return result;
		}

		private static readonly IReadOnlyDictionary<string, object?> EmptyProperties =
			new Dictionary<string, object?>(0);
	}
}
