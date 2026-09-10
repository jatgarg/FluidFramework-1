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
	/// <para />
	/// The <c>properties</c> constructor parameter is typed
	/// <see cref="Dictionary{TKey, TValue}" /> (not
	/// <see cref="IReadOnlyDictionary{TKey, TValue}" />) so callers can
	/// use target-typed <c>new()</c> + collection-initializer syntax at
	/// the throw site (e.g., <c>throw new LoggingError(msg, new() { ["k"] = v })</c>).
	/// The read-only view is preserved on the outbound side via
	/// <see cref="GetTelemetryProperties" />.
	/// </remarks>
	public class LoggingError : Exception, ILoggingError
	{
		private readonly Dictionary<string, object?> _properties;

		public LoggingError(string message)
			: this(message, properties: null, innerException: null)
		{
		}

		public LoggingError(string message, Dictionary<string, object?>? properties)
			: this(message, properties, innerException: null)
		{
		}

		public LoggingError(string message, Exception? innerException)
			: this(message, properties: null, innerException)
		{
		}

		public LoggingError(string message, Dictionary<string, object?>? properties, Exception? innerException)
			: base(message, innerException)
		{
			// Copy the caller's dict so post-construction mutations on their
			// side can't leak into the exception payload.
			_properties = properties is null
				? new Dictionary<string, object?>(0)
				: new Dictionary<string, object?>(properties);
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
			// message + stack follow TS extractLogSafeErrorProperties(sanitizeStack: true):
			//  - message: UserData by default (may embed interpolated user content).
			//  - stack: CodeArtifact after sanitization. .NET Exception.StackTrace
			//    doesn't include a leading "Name: Message" line, so no scrubbing
			//    is needed here — the CodeArtifact classification stands as-is.
			Dictionary<string, object?> result = new(_properties.Count + 3)
			{
				["message"] = new TaggedTelemetryValue(Message, FluidTelemetryDataTag.UserData),
				["stack"] = new TaggedTelemetryValue(StackTrace, FluidTelemetryDataTag.CodeArtifact),
				["errorInstanceId"] = ErrorInstanceId,
			};

			foreach (KeyValuePair<string, object?> pair in _properties)
			{
				result[pair.Key] = pair.Value;
			}

			return result;
		}
	}
}
