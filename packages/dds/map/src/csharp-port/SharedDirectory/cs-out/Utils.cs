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
	/// Minimal port of Node.js <c>path.posix</c> semantics used by SharedDirectory
	/// for path normalization. TS SharedDirectory uses <c>posix.join</c> and
	/// <c>posix.resolve</c> (see packages/dds/map/src/directory.ts) to keep every
	/// SubDirectory path canonical (leading '/', no '.' or '..' segments, no
	/// duplicate slashes). The C# port must match those semantics because path
	/// bytes appear on the wire in every directory operation.
	/// </summary>
	/// <remarks>
	/// Not a full port of Node's posix module — only the two operations we need.
	/// </remarks>
	public static class PosixPath
	{
		/// <summary>
		/// Equivalent of Node.js <c>posix.join(parent, child)</c>. Joins two POSIX
		/// path segments and normalizes the result: drops <c>.</c> segments, resolves
		/// <c>..</c> segments up to the root, and collapses consecutive slashes.
		/// </summary>
		/// <remarks>
		/// Matches Node's behavior for the shapes SharedDirectory encounters:
		/// <c>posix.join("/", "a") == "/a"</c>,
		/// <c>posix.join("/a", ".") == "/a"</c>,
		/// <c>posix.join("/a", "..") == "/"</c>,
		/// <c>posix.join("/", "") == "/"</c>.
		/// </remarks>
		public static string Join(string parent, string child)
		{
			ArgumentNullException.ThrowIfNull(parent);
			ArgumentNullException.ThrowIfNull(child);

			// Concatenate with a separator, matching Node's posix.join preprocess step.
			string combined = string.IsNullOrEmpty(child)
				? parent
				: string.IsNullOrEmpty(parent)
					? child
					: $"{parent}/{child}";

			return Normalize(combined);
		}

		/// <summary>
		/// Equivalent of Node.js <c>posix.resolve('/', relativePath)</c>. Produces
		/// an absolute canonical POSIX path from a relative or absolute input.
		/// </summary>
		/// <remarks>
		/// Matches Node's behavior for the shapes SharedDirectory encounters:
		/// <c>posix.resolve('/', '') == '/'</c>,
		/// <c>posix.resolve('/', 'a/b') == '/a/b'</c>,
		/// <c>posix.resolve('/', '/a/../b') == '/b'</c>,
		/// <c>posix.resolve('/', '../a') == '/a'</c>.
		/// Attempts to walk above the root (via extra <c>..</c>) stay at root, matching
		/// posix.resolve semantics.
		/// </remarks>
		public static string ResolveAbsolute(string relativePath)
		{
			ArgumentNullException.ThrowIfNull(relativePath);

			// posix.resolve('/', relativePath) is equivalent to normalizing '/' + relativePath.
			// An empty relativePath resolves to just '/'.
			return Normalize(string.IsNullOrEmpty(relativePath) ? "/" : $"/{relativePath}");
		}

		/// <summary>
		/// Canonicalizes a POSIX path in-place: drops empty segments, drops <c>.</c>
		/// segments, resolves <c>..</c> segments (up to the root). Preserves the
		/// absolute-vs-relative distinction based on whether the input started with
		/// <c>'/'</c>. The result always uses <c>'/'</c> as the separator and never
		/// contains duplicate slashes.
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
