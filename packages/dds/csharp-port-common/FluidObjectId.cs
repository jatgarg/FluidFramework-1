// -----------------------------------------------------------------------------
// SHIM — deleted on transfer to host repo.
// Copied verbatim from host repo:
//   waccobalt/.../Fluid/FluidObjectId.cs
// -----------------------------------------------------------------------------

using System;

namespace Microsoft.Office.Web.Fluid
{
    public static class FluidObjectId
    {
        public static string CreateId()
        {
            return Guid.NewGuid().ToString();
        }
    }
}
