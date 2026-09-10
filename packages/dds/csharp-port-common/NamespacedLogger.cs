// -----------------------------------------------------------------------------
// TS ref: @fluidframework/telemetry-utils ChildLogger (namespace prefix path).
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Wraps a parent <see cref="IFluidLogger" /> and prepends a namespace
	/// prefix to every emitted event name, separated by
	/// <see cref="EventNamespaceSeparator" /> (<c>":"</c>). Chains through
	/// nested <see cref="IFluidLogger.CreateChildLogger" /> calls.
	/// </summary>
	public sealed class NamespacedLogger : IFluidLogger
	{
		/// <summary>
		/// Separator inserted between namespace and event name, and between
		/// nested child-logger namespace segments. Matches TS
		/// <c>@fluidframework/telemetry-utils</c>
		/// <c>eventNamespaceSeparator</c> — a colon — so JS and C# share
		/// Kusto queries on the composed <c>eventName</c>.
		/// </summary>
		public const string EventNamespaceSeparator = ":";

		private readonly IFluidLogger _parent;
		private readonly string _prefix;

		private NamespacedLogger(IFluidLogger parent, string prefix)
		{
			_parent = parent;
			_prefix = prefix;
		}

		/// <summary>
		/// Wrap <paramref name="parent" /> with a namespace prefix. Impls of
		/// <see cref="IFluidLogger" /> can route their own
		/// <see cref="IFluidLogger.CreateChildLogger" /> through this factory
		/// for consistent behavior.
		/// </summary>
		public static IFluidLogger Create(IFluidLogger parent, string namespaceName)
		{
			if (parent is null)
			{
				throw new ArgumentNullException(nameof(parent));
			}

			if (string.IsNullOrEmpty(namespaceName))
			{
				throw new ArgumentException("Namespace name must not be empty.", nameof(namespaceName));
			}

			return new NamespacedLogger(parent, namespaceName);
		}

		public void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
			_parent.SendTelemetryEvent(Prefix(evt), error, logLevel);
		}

		public void SendErrorEvent(FluidTelemetryEvent evt, Exception? error = null)
		{
			_parent.SendErrorEvent(Prefix(evt), error);
		}

		public void SendPerformanceEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? logLevel = null)
		{
			_parent.SendPerformanceEvent(Prefix(evt), error, logLevel);
		}

		public IFluidLogger CreateChildLogger(string namespaceName)
		{
			if (string.IsNullOrEmpty(namespaceName))
			{
				throw new ArgumentException("Namespace name must not be empty.", nameof(namespaceName));
			}

			return new NamespacedLogger(_parent, $"{_prefix}{EventNamespaceSeparator}{namespaceName}");
		}

		private FluidTelemetryEvent Prefix(FluidTelemetryEvent evt)
		{
			return new FluidTelemetryEvent
			{
				EventName = $"{_prefix}{EventNamespaceSeparator}{evt.EventName}",
				Properties = evt.Properties,
			};
		}
	}
}
