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
			: this(url, payloadPending: false)
		{
		}

		public SerializedFluidHandle(string url, bool payloadPending)
		{
			Url = url ?? throw new ArgumentNullException(nameof(url));
			PayloadPending = payloadPending;
		}

		public string Url { get; }

		/// <summary>True only when the received wire form explicitly carried <c>"payloadPending": true</c>.</summary>
		public bool PayloadPending { get; }

		public override string ToString()
		{
			return Url;
		}
	}
}
