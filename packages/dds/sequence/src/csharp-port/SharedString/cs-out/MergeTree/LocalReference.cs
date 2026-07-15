// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/localReference.ts (subset for POC)
// Part of the SharedString C# feasibility port — Wave 12a.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
	/// <summary>
	/// Directional preference used when a local reference slides after edits.
	/// </summary>
	public enum SlidingPreference
	{
		Backward,
		Forward,
	}

	internal enum ReferenceEndpointKind
	{
		None,
		Start,
		End,
	}

	/// <summary>
	/// A local reference attached to an offset within a merge-tree segment.
	/// </summary>
	public sealed class LocalReferencePosition
	{
		internal LocalReferencePosition(
			ISegment segment,
			int offset,
			ReferenceType refType,
			SlidingPreference slidingPreference = SlidingPreference.Forward,
			PropertySet? properties = null,
			bool canSlideToEndpoint = false)
		{
			ValidateReferenceType(refType);
			Segment = segment;
			Offset = offset;
			RefType = refType;
			SlidingPreference = slidingPreference;
			Properties = properties;
			CanSlideToEndpoint = canSlideToEndpoint;
		}

		internal ISegment? Segment { get; set; }

		internal int Offset { get; set; }

		internal ReferenceEndpointKind Endpoint { get; set; }

		/// <summary>
		/// Gets the reference type flags for this reference.
		/// </summary>
		public ReferenceType RefType { get; init; }

		/// <summary>
		/// Gets the preferred slide direction for this reference.
		/// </summary>
		public SlidingPreference SlidingPreference { get; init; }

		/// <summary>
		/// Gets a value indicating whether this reference may slide to a start/end endpoint sentinel.
		/// </summary>
		public bool CanSlideToEndpoint { get; init; }

		/// <summary>
		/// Gets or sets properties associated with this reference.
		/// </summary>
		public PropertySet? Properties { get; set; }

		/// <summary>
		/// Gets a value indicating whether this reference is detached from the merge tree.
		/// </summary>
		public bool IsDetached => Segment == null && Endpoint == ReferenceEndpointKind.None;

		internal bool Detach()
		{
			if (Segment is null)
			{
				if (Endpoint != ReferenceEndpointKind.None)
				{
					Endpoint = ReferenceEndpointKind.None;
					Offset = 0;
					return true;
				}

				return false;
			}

			bool removed = Segment.LocalRefs.Remove(this);
			Segment = null;
			Endpoint = ReferenceEndpointKind.None;
			Offset = 0;
			return removed;
		}

		private static void ValidateReferenceType(ReferenceType refType)
		{
			int exclusiveCount = 0;
			if ((refType & ReferenceType.Transient) != 0)
			{
				exclusiveCount++;
			}

			if ((refType & ReferenceType.SlideOnRemove) != 0)
			{
				exclusiveCount++;
			}

			if ((refType & ReferenceType.StayOnRemove) != 0)
			{
				exclusiveCount++;
			}

			if (exclusiveCount > 1)
			{
				throw new ArgumentException(
					"Reference types can only be one of Transient, SlideOnRemove, and StayOnRemove.",
					nameof(refType));
			}
		}
	}
}
