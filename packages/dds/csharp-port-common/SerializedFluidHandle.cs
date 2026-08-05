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
	/// <remarks>
	/// Mirrors the wire shape of TS <c>ISerializedHandle</c> from
	/// <c>packages/runtime/runtime-utils/src/handles.ts</c>: <c>{ type, url, payloadPending? }</c>.
	/// The optional <c>payloadPending</c> flag is preserved so that round-trips
	/// through this port do not change the wire shape.
	/// </remarks>
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

		/// <summary>
		/// Mirrors TS <c>ISerializedHandle.payloadPending?: true</c>. Always false unless
		/// the received wire form explicitly carried <c>"payloadPending": true</c>.
		/// </summary>
		public bool PayloadPending { get; }

		public override string ToString()
		{
			return Url;
		}
	}
}
