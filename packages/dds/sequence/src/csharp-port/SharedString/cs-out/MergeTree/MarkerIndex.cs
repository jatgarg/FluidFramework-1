// -----------------------------------------------------------------------------
// Marker id index for the SharedString C# feasibility port — Wave 11.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    public sealed class MarkerIndex
    {
        private readonly Dictionary<string, Marker> _markers = new(StringComparer.Ordinal);

        public int Count => _markers.Count;

        public void Add(Marker marker)
        {
            ArgumentNullException.ThrowIfNull(marker);

            string? markerId = marker.GetId();
            if (!string.IsNullOrEmpty(markerId))
            {
                _markers[markerId] = marker;
            }
        }

        public bool Remove(string? markerId)
        {
            if (string.IsNullOrEmpty(markerId))
            {
                return false;
            }

            return _markers.Remove(markerId);
        }

        public bool TryGet(string id, out Marker? marker)
        {
            if (string.IsNullOrEmpty(id))
            {
                marker = null;
                return false;
            }

            return _markers.TryGetValue(id, out marker);
        }

        public void Clear()
        {
            _markers.Clear();
        }
    }
}
