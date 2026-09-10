// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-utils assert().
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Runtime invariant check. Always throws in Debug and Release —
	/// semantically equivalent to Fluid JS <c>assert(cond, msg)</c>.
	/// </summary>
	/// <remarks>
	/// Named <c>FluidAssert</c> to disambiguate from <c>Xunit.Assert</c>,
	/// <see cref="System.Diagnostics.Debug.Assert(bool)" /> (stripped in
	/// Release), and Word's <c>Verify</c>. Tagged overloads mirror TS
	/// Fluid's build-assigned assertion tags — see
	/// <see cref="That(bool, uint)" />.
	/// </remarks>
	public static class FluidAssert
	{
		/// <summary>
		/// Throws <see cref="LoggingError" /> if <paramref name="condition" />
		/// is false. The <see cref="DoesNotReturnIfAttribute" /> mirrors TS's
		/// <c>asserts condition</c> type predicate — after a successful call
		/// the nullability analyzer treats <paramref name="condition" /> as
		/// true.
		/// </summary>
		public static void That(
			[DoesNotReturnIf(false)] bool condition,
			string message)
		{
			if (!condition)
			{
				throw new LoggingError(message);
			}
		}

		/// <summary>
		/// As <see cref="That(bool, string)" /> but attaches structured
		/// payload to the thrown <see cref="LoggingError" />.
		/// </summary>
		public static void That(
			[DoesNotReturnIf(false)] bool condition,
			string message,
			Dictionary<string, object?> properties)
		{
			if (!condition)
			{
				throw new LoggingError(message, properties);
			}
		}

		/// <summary>
		/// Tagged form for asserts with a TS ancestor. Throws
		/// <see cref="LoggingError" /> with message <c>"0xNNN"</c> —
		/// matching what TS Fluid throws — so Kusto queries key on
		/// a stable identifier across runtimes.
		/// </summary>
		/// <remarks>
		/// Use the same <paramref name="tag" /> as the TS ancestor.
		/// Description lives as a source comment at the call site
		/// (<c>FluidAssert.That(cond, 0x3fe /* description */);</c>),
		/// matching TS convention. Tag → description lookup happens
		/// via TS's generated <c>assertionShortCodesMap.ts</c>.
		/// </remarks>
		public static void That(
			[DoesNotReturnIf(false)] bool condition,
			uint tag)
		{
			if (!condition)
			{
				throw new LoggingError(FormatTag(tag));
			}
		}

		/// <summary>
		/// As <see cref="That(bool, uint)" /> but attaches structured
		/// payload to the thrown <see cref="LoggingError" />.
		/// </summary>
		public static void That(
			[DoesNotReturnIf(false)] bool condition,
			uint tag,
			Dictionary<string, object?> properties)
		{
			if (!condition)
			{
				throw new LoggingError(FormatTag(tag), properties);
			}
		}

		private static string FormatTag(uint tag)
		{
			// Matches TS assert.ts: `0x${tag.toString(16).padStart(3, "0")}`.
			return "0x" + tag.ToString("x3", CultureInfo.InvariantCulture);
		}
	}
}

