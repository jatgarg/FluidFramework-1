// -----------------------------------------------------------------------------
// SHIM — deleted on transfer to host repo.
//
// The host repo defines these types via Bond-generated wire schemas. For the
// demo we stub only the minimum surface referenced by the DDS ports so the
// C# port compiles standalone.
// -----------------------------------------------------------------------------

#nullable enable

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Attach message announcing a new DDS instance. Server-side clients don't
	/// act on this; state is loaded from the snapshot instead.
	/// </summary>
	public sealed class fluidDataStoreMessageAttach
	{
		public string id { get; set; } = string.Empty;
		public string type { get; set; } = string.Empty;
	}
}
