// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/utils.ts
// Part of the SharedDirectory C# feasibility port.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid
{
	internal static class DirectoryUtils
	{
		/// <summary>
		/// Rough polyfill for Array.findLastIndex until we target ES2023 or greater.
		/// </summary>
		public static int FindLastIndex<T>(IReadOnlyList<T> array, Func<T, bool> predicate)
		{
			for (int i = array.Count - 1; i >= 0; i--)
			{
				if (predicate(array[i]))
				{
					return i;
				}
			}

			return -1;
		}

		/// <summary>
		/// Rough polyfill for Array.findLast until we target ES2023 or greater.
		/// </summary>
		public static T? FindLast<T>(IReadOnlyList<T> array, Func<T, bool> predicate)
			where T : class
		{
			int index = FindLastIndex(array, predicate);
			if (index < 0)
			{
				return null;
			}

			return array[index];
		}
	}
}
