// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/textSegment.ts
// Part of the SharedString C# feasibility port — Wave 1.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Text.Json.Serialization;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    // BaseSegment, ISegment, IJSONSegment, and PropertySet are Wave 1 sibling ports expected
    // to land alongside this file.

    /// <summary>
    /// JSON wire shape for a serialized <see cref="TextSegment"/>.
    /// </summary>
    public interface IJSONTextSegment : IJSONSegment
    {
        /// <summary>
        /// Gets or sets the segment text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        /// Gets or sets the segment properties, or <c>null</c> if none.
        /// </summary>
        PropertySet? Props { get; set; }
    }

    /// <summary>
    /// Default JSON wire object for a serialized <see cref="TextSegment"/>.
    /// </summary>
    public sealed class JSONTextSegment : IJSONTextSegment
    {
        /// <inheritdoc />
        [JsonPropertyName("text")]
        public string Text { get; set; } = string.Empty;

        /// <inheritdoc />
        [JsonPropertyName("props")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public PropertySet? Props { get; set; }
    }

    /// <summary>
    /// A merge-tree segment containing text.
    /// </summary>
    public sealed class TextSegment : BaseSegment
    {
        /// <summary>
        /// Maximum length of a text segment to consider when coalescing adjacent segments.
        /// </summary>
        public const int TextSegmentGranularity = 256;

        /// <summary>
        /// Type discriminator for text segments.
        /// </summary>
        public const string TypeName = "TextSegment";

        private string _text;

        /// <summary>
        /// Initializes a new instance of the <see cref="TextSegment"/> class.
        /// </summary>
        /// <param name="text">The text content for the segment.</param>
        /// <param name="props">Optional segment properties.</param>
        public TextSegment(string text, PropertySet? props = null)
            : base(props)
        {
            _text = text ?? throw new ArgumentNullException(nameof(text));
            CachedLength = _text.Length;
        }

        /// <inheritdoc />
        public override string Type => TypeName;

        /// <summary>
        /// Gets or sets the text content of this segment.
        /// </summary>
        public string Text
        {
            get => _text;
            set
            {
                _text = value ?? throw new ArgumentNullException(nameof(value));
                CachedLength = _text.Length;
            }
        }

        /// <summary>
        /// Determines whether the supplied segment is a <see cref="TextSegment"/>.
        /// </summary>
        /// <param name="segment">The segment to test.</param>
        /// <returns><c>true</c> if the segment is a text segment; otherwise, <c>false</c>.</returns>
        public static bool Is(ISegment? segment)
        {
            return segment is TextSegment textSegment && textSegment.Type == TypeName;
        }

        /// <summary>
        /// Creates a text segment.
        /// </summary>
        /// <param name="text">The segment text.</param>
        /// <param name="props">Optional segment properties.</param>
        /// <returns>The created text segment.</returns>
        public static TextSegment Make(string text, PropertySet? props = null)
        {
            return new TextSegment(text, props);
        }

        /// <summary>
        /// Deserializes a text segment from its JSON wire representation.
        /// </summary>
        /// <param name="spec">A plain string or JSON text-segment object.</param>
        /// <returns>The deserialized segment, or <c>null</c> if the spec is not a text segment.</returns>
        public static TextSegment? FromJSONObject(object? spec)
        {
            switch (spec)
            {
                case string text:
                    return new TextSegment(text);

                case IJSONTextSegment textSpec:
                    return Make(textSpec.Text, textSpec.Props);

                default:
                    return null;
            }
        }

        /// <inheritdoc />
        public override object ToJSONObject()
        {
            if (Properties == null)
            {
                return _text;
            }

            return new JSONTextSegment()
            {
                Text = _text,
                Props = PropertyMap.ClonePropertySet(Properties),
            };
        }

        /// <inheritdoc />
        public override TextSegment Clone()
        {
            return Clone(0, null);
        }

        /// <summary>
        /// Clones a slice of this segment.
        /// </summary>
        /// <param name="start">The inclusive start offset.</param>
        /// <param name="end">The exclusive end offset, or <c>null</c> to clone through the end.</param>
        /// <returns>A cloned text segment containing the requested slice.</returns>
        public TextSegment Clone(int start, int? end = null)
        {
            TextSegment clone = Make(Slice(_text, start, end), Properties);
            CopyMetadataTo(clone);
            return clone;
        }

        /// <inheritdoc />
        public override bool CanAppend(ISegment segment)
        {
            if (segment is not TextSegment textSegment)
            {
                return false;
            }

            return !_text.EndsWith("\n", StringComparison.Ordinal)
                && (CachedLength <= TextSegmentGranularity || textSegment.CachedLength <= TextSegmentGranularity);
        }

        /// <inheritdoc />
        public override string ToString()
        {
            return _text;
        }

        /// <inheritdoc />
        public override void Append(ISegment segment)
        {
            if (segment is not TextSegment textSegment)
            {
                throw new ArgumentException("Can only append text segment.", nameof(segment));
            }

            base.Append(segment);
            _text += textSegment.Text;
        }

        /// <inheritdoc />
        protected override TextSegment? CreateSplitSegmentAt(int pos)
        {
            if (pos <= 0)
            {
                return null;
            }

            int splitPos = Math.Min(Math.Max(0, pos), _text.Length);
            string remainingText = _text.Substring(splitPos);
            _text = _text.Substring(0, splitPos);
            CachedLength = _text.Length;
            return new TextSegment(remainingText, PropertyMap.ClonePropertySet(Properties));
        }

        private static string Slice(string text, int start, int? end)
        {
            int normalizedStart = NormalizeSliceIndex(start, text.Length);
            int normalizedEnd = end.HasValue ? NormalizeSliceIndex(end.Value, text.Length) : text.Length;

            if (normalizedEnd <= normalizedStart)
            {
                return string.Empty;
            }

            return text.Substring(normalizedStart, normalizedEnd - normalizedStart);
        }

        private static int NormalizeSliceIndex(int index, int length)
        {
            int normalized = index < 0 ? length + index : index;
            return Math.Min(Math.Max(0, normalized), length);
        }

    }
}
