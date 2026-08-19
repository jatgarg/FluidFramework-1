// -----------------------------------------------------------------------------
// Ported from packages/dds/map/src/utils.ts
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Text;

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

	/// <summary>
	/// Path helper for SharedDirectory. Implements the subset of posix-style
	/// path operations that TS SharedDirectory (packages/dds/map/src/directory.ts)
	/// uses to keep every SubDirectory path canonical (leading '/', no '.' or
	/// '..' segments, no duplicate slashes). The C# port must match those
	/// semantics because path bytes appear on the wire in every directory operation.
	/// </summary>
	public static class DirectoryPath
	{
		/// <summary>
		/// Joins two POSIX path segments and normalizes the result: drops
		/// <c>.</c> segments, resolves <c>..</c> segments up to the root, and
		/// collapses consecutive slashes. Equivalent to Node.js
		/// <c>path.posix.join(parent, child)</c> for the shapes SharedDirectory
		/// encounters:
		/// <c>Join("/", "a") == "/a"</c>,
		/// <c>Join("/a", ".") == "/a"</c>,
		/// <c>Join("/a", "..") == "/"</c>,
		/// <c>Join("/", "") == "/"</c>.
		/// </summary>
		public static string Join(string parent, string child)
		{
			ArgumentNullException.ThrowIfNull(parent);
			ArgumentNullException.ThrowIfNull(child);

			string combined = string.IsNullOrEmpty(child)
				? parent
				: string.IsNullOrEmpty(parent)
					? child
					: $"{parent}/{child}";

			return Normalize(combined);
		}

		/// <summary>
		/// Produces an absolute canonical POSIX path from a relative or
		/// absolute input. Equivalent to Node.js
		/// <c>path.posix.resolve('/', relativePath)</c> for the shapes
		/// SharedDirectory encounters:
		/// <c>ResolveAbsolute("") == "/"</c>,
		/// <c>ResolveAbsolute("a/b") == "/a/b"</c>,
		/// <c>ResolveAbsolute("/a/../b") == "/b"</c>,
		/// <c>ResolveAbsolute("../a") == "/a"</c>.
		/// Excess <c>..</c> segments stay at the root.
		/// </summary>
		public static string ResolveAbsolute(string relativePath)
		{
			ArgumentNullException.ThrowIfNull(relativePath);

			// posix.resolve('/', relativePath) is equivalent to normalizing '/' + relativePath.
			return Normalize(string.IsNullOrEmpty(relativePath) ? "/" : $"/{relativePath}");
		}

		/// <summary>
		/// Canonicalizes a POSIX path: drops empty segments and <c>.</c>
		/// segments, resolves <c>..</c> segments (up to the root for absolute
		/// paths; preserved for unresolvable leading <c>..</c> on relative
		/// paths). Preserves the absolute-vs-relative distinction based on
		/// whether the input started with <c>'/'</c>. The result uses
		/// <c>'/'</c> as the separator and never contains duplicate slashes.
		/// </summary>
		internal static string Normalize(string path)
		{
			ArgumentNullException.ThrowIfNull(path);

			bool isAbsolute = path.StartsWith("/", StringComparison.Ordinal);
			string[] rawSegments = path.Split('/');
			List<string> stack = new List<string>(rawSegments.Length);

			foreach (string segment in rawSegments)
			{
				if (segment.Length == 0 || segment == ".")
				{
					// Empty segments happen from consecutive '/' or leading/trailing '/';
					// posix.normalize drops them. '.' segments are also dropped.
					continue;
				}

				if (segment == "..")
				{
					if (stack.Count > 0 && stack[stack.Count - 1] != "..")
					{
						stack.RemoveAt(stack.Count - 1);
					}
					else if (!isAbsolute)
					{
						// Relative path with unresolvable '..' — preserve (posix.normalize
						// keeps unresolvable leading '..' segments for relative paths).
						stack.Add("..");
					}
					// For absolute paths, extra '..' above root are dropped (matches posix.resolve).

					continue;
				}

				stack.Add(segment);
			}

			if (stack.Count == 0)
			{
				return isAbsolute ? "/" : ".";
			}

			StringBuilder sb = new StringBuilder(path.Length);
			if (isAbsolute)
			{
				sb.Append('/');
			}

			for (int i = 0; i < stack.Count; i++)
			{
				if (i > 0)
				{
					sb.Append('/');
				}

				sb.Append(stack[i]);
			}

			return sb.ToString();
		}
	}
}
