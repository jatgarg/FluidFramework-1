// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-interfaces LogLevel.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Numeric severity for <see cref="IFluidLogger" /> events. Values match
	/// Fluid JS <c>LogLevel</c> so JS/C# telemetry queries can share Kusto
	/// dashboards. Host bridges map these to their own severity type when
	/// forwarding (e.g. wordfluidcsharp maps to <c>Log.Level</c>).
	/// </summary>
	public enum FluidLogLevel
	{
		Verbose = 10,
		Info = 20,
		Essential = 30,
	}
}
