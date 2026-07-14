// -----------------------------------------------------------------------------
// Shared representation for unresolved Fluid handle wire values.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Opaque handle value used when a received Fluid handle cannot be resolved
	/// through an <see cref="IFluidDataObjectRegistry" />.
	/// </summary>
	public sealed class SerializedFluidHandle
	{
		public SerializedFluidHandle(string url)
		{
			Url = url ?? throw new ArgumentNullException(nameof(url));
		}

		public string Url { get; }

		public override string ToString()
		{
			return Url;
		}
	}
}
