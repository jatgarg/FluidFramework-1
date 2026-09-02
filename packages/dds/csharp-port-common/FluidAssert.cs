// -----------------------------------------------------------------------------
// TS ref: @fluidframework/core-utils assert().
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.Office.Web.Fluid
{
	/// <summary>
	/// Runtime invariant check. Always throws in Debug and Release —
	/// semantically equivalent to Fluid JS <c>assert(cond, msg)</c>.
	/// </summary>
	/// <remarks>
	/// Named <c>FluidAssert</c> to disambiguate from <c>Xunit.Assert</c>,
	/// <see cref="System.Diagnostics.Debug.Assert(bool)" /> (stripped in
	/// Release), and Word's <c>Verify</c>.
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
			IReadOnlyDictionary<string, object?> properties)
		{
			if (!condition)
			{
				throw new LoggingError(message, properties);
			}
		}
	}
}
