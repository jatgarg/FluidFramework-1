// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/snapshotV1.ts + snapshotLoader.ts
// (subset for POC — V1 chunked format, TextSegment only).
// Part of the SharedString C# feasibility port — Wave 5.
//
// Also handles the synthetic V1-lite format produced by
// SharedString.Tests/Fixtures/generate.mjs when real Fluid packages aren't
// resolvable. Both flavors decoded via the same Parse() → DTO → Populate() path.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Microsoft.Office.Web.Fluid
{
	// DTO for the parsed snapshot
	public sealed class SharedStringSnapshotDto
	{
		public string Version { get; set; } = "1";

		public int SegmentCount { get; set; }

		public int Length { get; set; }

		public int StartIndex { get; set; }

		public List<SharedStringSnapshotSegmentDto> Segments { get; set; } = new();

		public SharedStringSnapshotHeaderMetadata? HeaderMetadata { get; set; }

		public List<CatchupOpDto> CatchupOps { get; set; } = new();
	}

	public sealed class SharedStringSnapshotSegmentDto
	{
		public string? Text { get; set; } = string.Empty;

		public MergeTree.PropertySet? Props { get; set; }

		public bool IsMarker { get; set; }

		public MergeTree.ReferenceType MarkerRefType { get; set; } = MergeTree.ReferenceType.Simple;

		public bool IsUnknown { get; set; }

		// Merge info (only present in real V1 IJSONSegmentWithMergeInfo — optional for POC)
		public long? Seq { get; set; }

		public string? ClientId { get; set; }

		public long? RemovedSeq { get; set; }

		public string? RemovedClientId { get; set; }

		public long? ObliteratedSeq { get; set; }

		public string? ObliteratedClientId { get; set; }

		public List<SharedStringSnapshotRemoveStampDto> RemoveStamps { get; } = new();
	}

	public sealed class SharedStringSnapshotRemoveStampDto
	{
		public string Type { get; set; } = "setRemove";

		public long Seq { get; set; }

		public string ClientId { get; set; } = string.Empty;
	}

	public sealed class MergeTreeChunkV1Dto
	{
		public string Version { get; set; } = "1";

		public int SegmentCount { get; set; }

		public int Length { get; set; }

		public int StartIndex { get; set; }

		public List<SharedStringSnapshotSegmentDto> Segments { get; set; } = new();

		public MergeTreeHeaderMetadataDto? HeaderMetadata { get; set; }
	}

	public sealed class MergeTreeHeaderChunkMetadataDto
	{
		public string Id { get; set; } = string.Empty;
	}

	public class MergeTreeHeaderMetadataDto
	{
		public long MinSequenceNumber { get; set; }

		public long SequenceNumber { get; set; }

		public int TotalLength { get; set; }

		public int TotalSegmentCount { get; set; }

		public List<MergeTreeHeaderChunkMetadataDto> OrderedChunkMetadata { get; set; } = new();

		public List<string> CatchupOpsBlobNames { get; set; } = new();
	}

	public sealed class SharedStringSnapshotHeaderMetadata : MergeTreeHeaderMetadataDto
	{
	}

	public sealed class CatchupOpsBlobDto
	{
		public List<CatchupOpDto> Ops { get; set; } = new();
	}

	public sealed class CatchupOpDto
	{
		public string OpJson { get; set; } = string.Empty;

		public long? SequenceNumber { get; set; }

		public long? ReferenceSequenceNumber { get; set; }

		public long? MinimumSequenceNumber { get; set; }

		public string? ClientId { get; set; }
	}

	public static class SharedStringSnapshotLoader
	{
		/// <summary>Loads a SharedString snapshot JSON into an intermediate DTO tree.</summary>
		public static SharedStringSnapshotDto Load(
			string json,
			Func<string, string>? blobResolver = null,
			IFluidDataObjectRegistry? registry = null)
		{
			if (json is null)
			{
				throw new ArgumentNullException(nameof(json));
			}

			byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
			return Load(jsonBytes, blobResolver, registry);
		}

		/// <summary>Loads a SharedString snapshot JSON (UTF-8 bytes) into an intermediate DTO tree.</summary>
		public static SharedStringSnapshotDto Load(
			ReadOnlySpan<byte> jsonBytes,
			Func<string, string>? blobResolver = null,
			IFluidDataObjectRegistry? registry = null)
		{
			var reader = new Utf8JsonReader(jsonBytes);
			return ReadFrom(ref reader, blobResolver, registry);
		}

		/// <summary>Parses a SharedString snapshot JSON into an intermediate DTO tree.</summary>
		public static SharedStringSnapshotDto Parse(string json, IFluidDataObjectRegistry? registry = null)
		{
			if (json is null)
			{
				throw new ArgumentNullException(nameof(json));
			}

			byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
			return Load(jsonBytes, blobResolver: null, registry: registry);
		}

		/// <summary>Parses a SharedString snapshot JSON (UTF-8 bytes) into an intermediate DTO tree.</summary>
		public static SharedStringSnapshotDto Parse(ReadOnlySpan<byte> jsonBytes, IFluidDataObjectRegistry? registry = null)
		{
			return Load(jsonBytes, blobResolver: null, registry: registry);
		}

		/// <summary>Reads a SharedString snapshot from an open reader.</summary>
		public static SharedStringSnapshotDto ReadFrom(ref Utf8JsonReader reader, IFluidDataObjectRegistry? registry = null)
		{
			return ReadFrom(ref reader, blobResolver: null, registry: registry);
		}

		/// <summary>Reads a SharedString snapshot from an open reader.</summary>
		public static SharedStringSnapshotDto ReadFrom(
			ref Utf8JsonReader reader,
			Func<string, string>? blobResolver,
			IFluidDataObjectRegistry? registry = null)
		{
			using JsonDocument document = JsonDocument.ParseValue(ref reader);
			return ReadSnapshot(document.RootElement, blobResolver, registry);
		}

		/// <summary>
		/// Populates the target MergeTree/Client with the segments from a parsed snapshot.
		/// The target must be fresh (empty). No ops emitted; no events fired.
		/// After loading, MergeTree.CollabWindowCurrentSeq/MinSeq reflect the snapshot's headerMetadata.
		/// </summary>
		public static void PopulateFromSnapshot(
			MergeTree.Client target,
			SharedStringSnapshotDto snapshot,
			IFluidDataObjectRegistry? registry = null)
		{
			if (target is null)
			{
				throw new ArgumentNullException(nameof(target));
			}

			if (snapshot is null)
			{
				throw new ArgumentNullException(nameof(snapshot));
			}

			if (target.MergeTree.FirstSegment() is not null)
			{
				throw new OcsException(OcsGateErrorCode.InvalidState, "PopulateFromSnapshot requires empty target");
			}

			ValidateWholeSnapshotForPopulate(snapshot);
			SharedStringSnapshotHeaderMetadata headerMetadata = snapshot.HeaderMetadata!;
			List<MergeTree.ISegment> segments = new();

			foreach (SharedStringSnapshotSegmentDto segmentDto in snapshot.Segments)
			{
				if (segmentDto is null)
				{
					throw InvalidSnapshot("Snapshot segments cannot contain null entries.");
				}

				MergeTree.ISegment? segment = CreateSegment(segmentDto, registry);
				if (segment is null)
				{
					continue;
				}

				if (segment.CachedLength == 0)
				{
					continue;
				}

				long segmentSeq = segmentDto.Seq ?? MergeTree.Constants.UniversalSequenceNumber;
				int segmentClientId = segmentDto.ClientId is string clientId
					? target.MergeTree.GetOrAddShortClientId(clientId)
					: MergeTree.Constants.NonCollabClient;
				segment.Seq = segmentSeq;
				segment.ClientId = segmentClientId;

				foreach (SharedStringSnapshotRemoveStampDto removeStamp in segmentDto.RemoveStamps)
				{
					MergeTree.RemoveOperationStamp stamp = CreateSnapshotRemoveStamp(
						removeStamp,
						target.MergeTree.GetOrAddShortClientId(removeStamp.ClientId));
					MergeTree.Stamps.SpliceIntoList(segment.RemoveStamps, stamp);
				}

				if (!HasSnapshotStampType(segmentDto, "setRemove") && segmentDto.RemovedSeq is long removedSeq)
				{
					MergeTree.Stamps.SpliceIntoList(
						segment.RemoveStamps,
						new MergeTree.SetRemoveOperationStamp()
						{
							Seq = removedSeq,
							ClientId = segmentDto.RemovedClientId is string removedClientId
								? target.MergeTree.GetOrAddShortClientId(removedClientId)
								: segmentClientId,
						});
				}

				if (!HasSnapshotStampType(segmentDto, "sliceRemove") && segmentDto.ObliteratedSeq is long obliteratedSeq)
				{
					MergeTree.Stamps.SpliceIntoList(
						segment.RemoveStamps,
						new MergeTree.SliceRemoveOperationStamp()
						{
							Seq = obliteratedSeq,
							ClientId = segmentDto.ObliteratedClientId is string obliteratedClientId
								? target.MergeTree.GetOrAddShortClientId(obliteratedClientId)
								: segmentClientId,
						});
				}

				segments.Add(segment);
				if (segment is MergeTree.Marker marker && segment.RemoveStamps.Count == 0)
				{
					target.Markers.Add(marker);
				}
			}

			target.MergeTree.RebuildFromSegments(segments);
			target.SetCollaborationWindow(headerMetadata.MinSequenceNumber, headerMetadata.SequenceNumber, runZamboni: false);
			ApplyCatchupOps(target, snapshot, registry);
		}

		private static MergeTree.ISegment? CreateSegment(
			SharedStringSnapshotSegmentDto segmentDto,
			IFluidDataObjectRegistry? registry)
		{
			MergeTree.PropertySet? properties = MergeTree.PropertyMap.ClonePropertySet(
				ResolvePropertySetHandles(segmentDto.Props, registry));
			if (segmentDto.IsUnknown)
			{
				return null;
			}

			if (segmentDto.IsMarker)
			{
				return MergeTree.Marker.Make(segmentDto.MarkerRefType, properties);
			}

			return MergeTree.TextSegment.Make(segmentDto.Text ?? string.Empty, properties);
		}

		private static MergeTree.PropertySet? ResolvePropertySetHandles(
			MergeTree.PropertySet? properties,
			IFluidDataObjectRegistry? registry)
		{
			if (properties is null)
			{
				return null;
			}

			MergeTree.PropertySet resolvedProperties = new();
			foreach (KeyValuePair<string, object?> property in properties)
			{
				resolvedProperties[property.Key] = HandleWireFormat.ResolveSerializedHandles(property.Value, registry);
			}

			return resolvedProperties;
		}

		private static void ApplyCatchupOps(
			MergeTree.Client target,
			SharedStringSnapshotDto snapshot,
			IFluidDataObjectRegistry? registry)
		{
			int index = 0;
			foreach (CatchupOpDto catchupOp in snapshot.CatchupOps)
			{
				// All sequence-numbering + clientId fields are required — reject rather than
				// default to synthetic values that would mask malformed snapshots.
				long sequenceNumber = catchupOp.SequenceNumber
					?? throw InvalidSnapshot($"SharedString catchup operation at index {index} is missing 'sequenceNumber'.");
				long referenceSequenceNumber = catchupOp.ReferenceSequenceNumber
					?? throw InvalidSnapshot($"SharedString catchup operation at index {index} is missing 'referenceSequenceNumber'.");
				long minimumSequenceNumber = catchupOp.MinimumSequenceNumber
					?? throw InvalidSnapshot($"SharedString catchup operation at index {index} is missing 'minimumSequenceNumber'.");
				string clientId = catchupOp.ClientId
					?? throw InvalidSnapshot($"SharedString catchup operation at index {index} is missing 'clientId'.");

				if (minimumSequenceNumber < target.CollabWindowMinSeq
					|| referenceSequenceNumber < target.CollabWindowMinSeq
					|| sequenceNumber <= target.CollabWindowMinSeq
					|| sequenceNumber < target.CollabWindowCurrentSeq)
				{
					throw InvalidSnapshot($"Invalid SharedString catchup operation sequence numbers at index {index}.");
				}

				MergeTree.IMergeTreeOp op = SharedStringOpSerializer.Deserialize(catchupOp.OpJson, registry);
				target.ApplyOp(
					op,
					seq: sequenceNumber,
					refSeq: referenceSequenceNumber,
					clientId: clientId);

				if (minimumSequenceNumber > target.CollabWindowMinSeq)
				{
					target.SetCollaborationWindow(minimumSequenceNumber, target.CollabWindowCurrentSeq, runZamboni: false);
				}

				index++;
			}
		}

		private static SharedStringSnapshotDto ReadSnapshot(
			JsonElement element,
			Func<string, string>? blobResolver,
			IFluidDataObjectRegistry? registry)
		{
			SharedStringSnapshotDto dto = ReadHeaderChunk(element, registry);
			LoadAdditionalChunks(dto, blobResolver, registry);
			LoadCatchupOps(dto, blobResolver);
			return dto;
		}

		private static SharedStringSnapshotDto ReadHeaderChunk(JsonElement element, IFluidDataObjectRegistry? registry)
		{
			MergeTreeChunkV1Dto chunk = ReadChunk(element, "root", registry);
			var dto = new SharedStringSnapshotDto()
			{
				Version = chunk.Version,
				SegmentCount = chunk.SegmentCount,
				Length = chunk.Length,
				StartIndex = chunk.StartIndex,
				Segments = chunk.Segments,
				HeaderMetadata = chunk.HeaderMetadata as SharedStringSnapshotHeaderMetadata,
			};

			if (chunk.HeaderMetadata is not null && dto.HeaderMetadata is null)
			{
				dto.HeaderMetadata = CopyHeaderMetadata(chunk.HeaderMetadata);
			}

			return dto;
		}

		private static MergeTreeChunkV1Dto ReadChunk(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			if (element.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"SharedString snapshot chunk at {path} must be a JSON object.");
			}

			string? version = ReadOptionalStringProperty(element, "version", path);
			if (version is null)
			{
				return ReadLegacyChunk(element, path, registry);
			}

			if (!string.Equals(version, "1", StringComparison.Ordinal))
			{
				throw InvalidSnapshot($"Unsupported SharedString snapshot version '{version}'.");
			}

			var dto = new MergeTreeChunkV1Dto()
			{
				Version = version,
				SegmentCount = ReadRequiredIntProperty(element, "segmentCount", path),
				Length = ReadRequiredIntProperty(element, "length", path),
				StartIndex = ReadRequiredIntProperty(element, "startIndex", path),
			};

			JsonElement segmentsElement = ReadRequiredProperty(element, "segments", path);
			ReadSegments(segmentsElement, dto.Segments, $"{path}.segments", registry);

			if (element.TryGetProperty("headerMetadata", out JsonElement headerMetadataElement)
				&& headerMetadataElement.ValueKind != JsonValueKind.Null
				&& headerMetadataElement.ValueKind != JsonValueKind.Undefined)
			{
				dto.HeaderMetadata = ReadHeaderMetadata(headerMetadataElement, $"{path}.headerMetadata");
			}

			ValidateChunkShape(dto, path);
			return dto;
		}

		private static MergeTreeChunkV1Dto ReadLegacyChunk(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			var dto = new MergeTreeChunkV1Dto()
			{
				Version = "1",
				SegmentCount = ReadRequiredIntProperty(element, "chunkSegmentCount", path),
				Length = ReadRequiredIntProperty(element, "chunkLengthChars", path),
				StartIndex = ReadRequiredIntProperty(element, "chunkStartSegmentIndex", path),
			};

			JsonElement segmentsElement = ReadRequiredProperty(element, "segmentTexts", path);
			ReadSegments(segmentsElement, dto.Segments, $"{path}.segmentTexts", registry);

			if (element.TryGetProperty("headerMetadata", out JsonElement headerMetadataElement)
				&& headerMetadataElement.ValueKind != JsonValueKind.Null
				&& headerMetadataElement.ValueKind != JsonValueKind.Undefined)
			{
				dto.HeaderMetadata = ReadHeaderMetadata(headerMetadataElement, $"{path}.headerMetadata");
			}
			else
			{
				dto.HeaderMetadata = TryBuildLegacyHeaderMetadata(element, dto, path);
			}

			ValidateChunkShape(dto, path);
			return dto;
		}

		private static void ReadSegments(
			JsonElement segmentsElement,
			List<SharedStringSnapshotSegmentDto> segments,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			if (segmentsElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"SharedString snapshot {path} must be a JSON array.");
			}

			int index = 0;
			foreach (JsonElement segmentElement in segmentsElement.EnumerateArray())
			{
				segments.Add(ReadSegment(segmentElement, $"{path}[{index}]", registry));
				index++;
			}
		}

		private static SharedStringSnapshotHeaderMetadata? TryBuildLegacyHeaderMetadata(
			JsonElement element,
			MergeTreeChunkV1Dto chunk,
			string path)
		{
			int? totalLength = ReadOptionalIntProperty(element, "totalLengthChars", path);
			int? totalSegmentCount = ReadOptionalIntProperty(element, "totalSegmentCount", path);
			long? sequenceNumber = ReadOptionalLongProperty(element, "chunkSequenceNumber", path);
			long? minSequenceNumber = ReadOptionalLongProperty(element, "chunkMinSequenceNumber", path);
			if (totalLength is null
				&& totalSegmentCount is null
				&& sequenceNumber is null
				&& minSequenceNumber is null)
			{
				return null;
			}

			if (totalLength is null
				|| totalSegmentCount is null
				|| sequenceNumber is null)
			{
				throw InvalidSnapshot($"Legacy SharedString snapshot header metadata at {path} is incomplete.");
			}

			SharedStringSnapshotHeaderMetadata metadata = new()
			{
				MinSequenceNumber = minSequenceNumber ?? sequenceNumber.Value,
				SequenceNumber = sequenceNumber.Value,
				TotalLength = totalLength.Value,
				TotalSegmentCount = totalSegmentCount.Value,
			};
			metadata.OrderedChunkMetadata.Add(new MergeTreeHeaderChunkMetadataDto()
			{
				Id = "header",
			});
			if (chunk.Length < totalLength.Value)
			{
				metadata.OrderedChunkMetadata.Add(new MergeTreeHeaderChunkMetadataDto()
				{
					Id = "body",
				});
			}

			return metadata;
		}

		private static void LoadAdditionalChunks(
			SharedStringSnapshotDto snapshot,
			Func<string, string>? blobResolver,
			IFluidDataObjectRegistry? registry)
		{
			SharedStringSnapshotHeaderMetadata? headerMetadata = snapshot.HeaderMetadata;
			if (headerMetadata is null)
			{
				return;
			}

			if (headerMetadata.OrderedChunkMetadata.Count <= 1)
			{
				return;
			}

			if (blobResolver is null)
			{
				throw InvalidSnapshot("SharedString multi-chunk snapshot requires a blobResolver.");
			}

			if (headerMetadata.OrderedChunkMetadata.Count < 2)
			{
				throw InvalidSnapshot("SharedString multi-chunk snapshot requires headerMetadata.orderedChunkMetadata.");
			}

			int totalSegmentCount = snapshot.SegmentCount;
			int totalLength = snapshot.Length;
			for (int chunkIndex = 1; chunkIndex < headerMetadata.OrderedChunkMetadata.Count; chunkIndex++)
			{
				string chunkName = headerMetadata.OrderedChunkMetadata[chunkIndex].Id;
				string chunkJson = ResolveBlob(blobResolver, chunkName);
				using JsonDocument chunkDocument = JsonDocument.Parse(chunkJson);
				MergeTreeChunkV1Dto chunk = ReadChunk(chunkDocument.RootElement, $"blob '{chunkName}'", registry);
				if (chunk.StartIndex != totalSegmentCount)
				{
					throw InvalidSnapshot($"SharedString chunk '{chunkName}' startIndex does not match the accumulated segment count.");
				}

				foreach (SharedStringSnapshotSegmentDto segment in chunk.Segments)
				{
					snapshot.Segments.Add(segment);
				}

				checked
				{
					totalSegmentCount += chunk.SegmentCount;
					totalLength += chunk.Length;
				}
			}

			if (totalSegmentCount != headerMetadata.TotalSegmentCount
				|| totalLength != headerMetadata.TotalLength)
			{
				throw InvalidSnapshot("SharedString multi-chunk snapshot totals do not match headerMetadata.");
			}

			snapshot.SegmentCount = totalSegmentCount;
			snapshot.Length = totalLength;
			ValidateChunkShape(snapshot);
		}

		private static void LoadCatchupOps(SharedStringSnapshotDto snapshot, Func<string, string>? blobResolver)
		{
			SharedStringSnapshotHeaderMetadata? headerMetadata = snapshot.HeaderMetadata;
			if (headerMetadata is null)
			{
				return;
			}

			if (headerMetadata.CatchupOpsBlobNames.Count > 0)
			{
				if (blobResolver is null)
				{
					throw InvalidSnapshot("SharedString catchup ops snapshot requires a blobResolver.");
				}

				foreach (string blobName in headerMetadata.CatchupOpsBlobNames)
				{
					string blobJson = ResolveBlob(blobResolver, blobName);
					CatchupOpsBlobDto catchupOpsBlob = ReadCatchupOpsBlob(blobJson, $"blob '{blobName}'");
					foreach (CatchupOpDto catchupOp in catchupOpsBlob.Ops)
					{
						snapshot.CatchupOps.Add(catchupOp);
					}
				}

				return;
			}

			// TS ref: merge-tree/src/snapshotLoader.ts loadBodyAndCatchupOps —
			// TS's canonical legacy format does not carry catchupOpsBlobNames;
			// TS discovers the catchup blob by listing storage and finding the
			// unnamed extra blob (default name "catchupOps" per snapshotlegacy.ts).
			// The port has no storage-list API, so probe the well-known name
			// via the caller's resolver. The probe stays silent: if the
			// resolver returns null / throws / returns a non-catchup shape,
			// treat it as "no catchup blob" and continue.
			if (blobResolver is null)
			{
				return;
			}

			string? probedJson;
			try
			{
				probedJson = blobResolver(LegacyCatchupOpsBlobName);
			}
			catch
			{
				return;
			}

			if (string.IsNullOrEmpty(probedJson) || !TryReadCatchupOpsBlob(probedJson!, out CatchupOpsBlobDto? probedBlob))
			{
				return;
			}

			foreach (CatchupOpDto catchupOp in probedBlob!.Ops)
			{
				snapshot.CatchupOps.Add(catchupOp);
			}
		}

		// TS ref: merge-tree/src/snapshotlegacy.ts SnapshotLegacy.catchupOps —
		// well-known blob name TS uses for the pre-name-list catchup ops layout.
		private const string LegacyCatchupOpsBlobName = "catchupOps";

		private static bool TryReadCatchupOpsBlob(string json, out CatchupOpsBlobDto? blob)
		{
			blob = null;
			try
			{
				using JsonDocument document = JsonDocument.Parse(json);
				JsonElement opsElement = document.RootElement;
				if (opsElement.ValueKind == JsonValueKind.Object
					&& opsElement.TryGetProperty("ops", out JsonElement nestedOpsElement))
				{
					opsElement = nestedOpsElement;
				}

				if (opsElement.ValueKind != JsonValueKind.Array)
				{
					return false;
				}

				var dto = new CatchupOpsBlobDto();
				int index = 0;
				foreach (JsonElement opElement in opsElement.EnumerateArray())
				{
					// Each entry must be an object with the required catchup fields;
					// otherwise this isn't a catchup blob.
					if (opElement.ValueKind != JsonValueKind.Object)
					{
						return false;
					}

					dto.Ops.Add(ReadCatchupOp(opElement, $"probe[{index}]"));
					index++;
				}

				blob = dto;
				return true;
			}
			catch
			{
				return false;
			}
		}

		private static CatchupOpsBlobDto ReadCatchupOpsBlob(string json, string path)
		{
			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement opsElement = document.RootElement;
			if (opsElement.ValueKind == JsonValueKind.Object
				&& opsElement.TryGetProperty("ops", out JsonElement nestedOpsElement))
			{
				opsElement = nestedOpsElement;
			}

			if (opsElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"SharedString catchup ops at {path} must be a JSON array.");
			}

			var dto = new CatchupOpsBlobDto();
			int index = 0;
			foreach (JsonElement opElement in opsElement.EnumerateArray())
			{
				dto.Ops.Add(ReadCatchupOp(opElement, $"{path}[{index}]"));
				index++;
			}

			return dto;
		}

		private static CatchupOpDto ReadCatchupOp(JsonElement element, string path)
		{
			// Each catchup entry must be an object with full sequence-numbering + clientId + contents.
			if (element.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must be a JSON object (an ISequencedDocumentMessage).");
			}

			long? sequenceNumber = ReadOptionalLongProperty(element, "sequenceNumber", path);
			if (sequenceNumber is null)
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must contain 'sequenceNumber'.");
			}

			long? referenceSequenceNumber = ReadOptionalLongProperty(element, "referenceSequenceNumber", path);
			if (referenceSequenceNumber is null)
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must contain 'referenceSequenceNumber'.");
			}

			long? minimumSequenceNumber = ReadOptionalLongProperty(element, "minimumSequenceNumber", path);
			if (minimumSequenceNumber is null)
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must contain 'minimumSequenceNumber'.");
			}

			if (!element.TryGetProperty("clientId", out JsonElement clientIdElement))
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must contain 'clientId'.");
			}

			if (clientIdElement.ValueKind != JsonValueKind.String)
			{
				throw InvalidSnapshot($"SharedString catchup op at {path}.clientId must be a string.");
			}

			string clientId = clientIdElement.GetString()!;

			if (!element.TryGetProperty("contents", out JsonElement contentsElement))
			{
				throw InvalidSnapshot($"SharedString catchup op at {path} must contain 'contents' (the merge-tree op).");
			}

			var dto = new CatchupOpDto()
			{
				SequenceNumber = sequenceNumber,
				ReferenceSequenceNumber = referenceSequenceNumber,
				MinimumSequenceNumber = minimumSequenceNumber,
				ClientId = clientId,
				OpJson = ReadOpJson(contentsElement, $"{path}.contents"),
			};

			if (string.IsNullOrEmpty(dto.OpJson))
			{
				throw InvalidSnapshot($"SharedString catchup op at {path}.contents must not be empty.");
			}

			return dto;
		}

		private static string ReadOpJson(JsonElement element, string path)
		{
			if (element.ValueKind == JsonValueKind.String)
			{
				return element.GetString() ?? string.Empty;
			}

			if (element.ValueKind == JsonValueKind.Object)
			{
				return element.GetRawText();
			}

			throw InvalidSnapshot($"SharedString catchup op payload at {path} must be a string or object.");
		}

		private static string ResolveBlob(Func<string, string> blobResolver, string blobName)
		{
			if (string.IsNullOrEmpty(blobName))
			{
				throw InvalidSnapshot("SharedString snapshot blob name must not be empty.");
			}

			string? blobJson = blobResolver(blobName);
			if (blobJson is null)
			{
				throw InvalidSnapshot($"Blob resolver returned null for blob '{blobName}'.");
			}

			return blobJson;
		}

		private static SharedStringSnapshotHeaderMetadata CopyHeaderMetadata(MergeTreeHeaderMetadataDto source)
		{
			var metadata = new SharedStringSnapshotHeaderMetadata()
			{
				MinSequenceNumber = source.MinSequenceNumber,
				SequenceNumber = source.SequenceNumber,
				TotalLength = source.TotalLength,
				TotalSegmentCount = source.TotalSegmentCount,
			};

			foreach (MergeTreeHeaderChunkMetadataDto chunkMetadata in source.OrderedChunkMetadata)
			{
				metadata.OrderedChunkMetadata.Add(new MergeTreeHeaderChunkMetadataDto()
				{
					Id = chunkMetadata.Id,
				});
			}

			foreach (string catchupOpsBlobName in source.CatchupOpsBlobNames)
			{
				metadata.CatchupOpsBlobNames.Add(catchupOpsBlobName);
			}

			return metadata;
		}

		private static SharedStringSnapshotSegmentDto ReadSegment(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			var dto = new SharedStringSnapshotSegmentDto();
			JsonElement payloadElement = element;

			if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty("json", out JsonElement wrappedPayloadElement))
			{
				payloadElement = wrappedPayloadElement;
				dto.Seq = ReadOptionalLongProperty(element, "seq", path);
				dto.ClientId = ReadOptionalStringProperty(element, "client", path)
					?? ReadOptionalStringProperty(element, "clientId", path);
				dto.RemovedSeq = ReadOptionalLongProperty(element, "removedSeq", path);
				dto.RemovedClientId = ReadOptionalRemovedClientId(element, path);
				dto.ObliteratedSeq = ReadOptionalLongProperty(element, "movedSeq", path);
				dto.ObliteratedClientId = ReadOptionalMovedClientId(element, path);
				ReadRemoveStamps(element, dto, path);
			}

			ReadTextSegmentPayload(payloadElement, dto, path, registry);
			return dto;
		}

		private static void ReadTextSegmentPayload(
			JsonElement element,
			SharedStringSnapshotSegmentDto dto,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.String:
					dto.Text = element.GetString() ?? string.Empty;
					return;

				case JsonValueKind.Object:
					ReadSegmentObject(element, dto, path, registry);
					return;

				default:
					dto.IsUnknown = true;
					dto.Text = string.Empty;
					return;
			}
		}

		private static void ReadSegmentObject(
			JsonElement element,
			SharedStringSnapshotSegmentDto dto,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			if (element.TryGetProperty("text", out JsonElement textElement))
			{
				if (textElement.ValueKind != JsonValueKind.String)
				{
					throw InvalidSnapshot($"Segment at {path} must contain a string 'text' field.");
				}

				dto.Text = textElement.GetString() ?? string.Empty;
				if (element.TryGetProperty("props", out JsonElement propsElement)
					&& propsElement.ValueKind != JsonValueKind.Null
					&& propsElement.ValueKind != JsonValueKind.Undefined)
				{
					dto.Props = ReadPropertySet(propsElement, $"{path}.props", registry);
				}

				return;
			}

			if (element.TryGetProperty("marker", out JsonElement markerElement))
			{
				ReadMarkerSegmentObject(element, markerElement, dto, path, registry);
				return;
			}

			dto.IsUnknown = true;
			dto.Text = string.Empty;
		}

		private static void ReadMarkerSegmentObject(
			JsonElement element,
			JsonElement markerElement,
			SharedStringSnapshotSegmentDto dto,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			dto.IsMarker = true;
			dto.Text = null;
			MergeTree.PropertySet? nestedProps = null;
			if (markerElement.ValueKind == JsonValueKind.Object)
			{
				if (markerElement.TryGetProperty("refType", out JsonElement refTypeElement))
				{
					dto.MarkerRefType = (MergeTree.ReferenceType)ReadInt(refTypeElement, $"{path}.marker.refType");
				}

				if (markerElement.TryGetProperty("props", out JsonElement nestedPropsElement)
					&& nestedPropsElement.ValueKind != JsonValueKind.Null
					&& nestedPropsElement.ValueKind != JsonValueKind.Undefined)
				{
					nestedProps = ReadPropertySet(nestedPropsElement, $"{path}.marker.props", registry);
				}
			}
			else if (markerElement.ValueKind != JsonValueKind.Null
				&& markerElement.ValueKind != JsonValueKind.Undefined)
			{
				throw InvalidSnapshot($"Segment at {path}.marker must be an object or null.");
			}

			dto.Props = nestedProps;
			if (element.TryGetProperty("props", out JsonElement propsElement)
				&& propsElement.ValueKind != JsonValueKind.Null
				&& propsElement.ValueKind != JsonValueKind.Undefined)
			{
				dto.Props = ReadPropertySet(propsElement, $"{path}.props", registry);
			}
		}

		private static SharedStringSnapshotHeaderMetadata ReadHeaderMetadata(JsonElement element, string path)
		{
			if (element.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Header metadata at {path} must be a JSON object.");
			}

			long sequenceNumber = ReadRequiredLongProperty(element, "sequenceNumber", path);
			var metadata = new SharedStringSnapshotHeaderMetadata()
			{
				MinSequenceNumber = ReadOptionalLongProperty(element, "minSequenceNumber", path) ?? sequenceNumber,
				SequenceNumber = sequenceNumber,
				TotalLength = ReadRequiredIntProperty(element, "totalLength", path),
				TotalSegmentCount = ReadRequiredIntProperty(element, "totalSegmentCount", path),
			};

			if (!element.TryGetProperty("orderedChunkMetadata", out JsonElement orderedChunkMetadataElement)
				|| orderedChunkMetadataElement.ValueKind == JsonValueKind.Null
				|| orderedChunkMetadataElement.ValueKind == JsonValueKind.Undefined)
			{
				throw InvalidSnapshot($"JSON property {path}.orderedChunkMetadata is required.");
			}

			ReadOrderedChunkMetadata(orderedChunkMetadataElement, metadata.OrderedChunkMetadata, $"{path}.orderedChunkMetadata");

			ReadCatchupOpsBlobNames(element, metadata.CatchupOpsBlobNames, path);
			return metadata;
		}

		private static void ReadOrderedChunkMetadata(
			JsonElement element,
			List<MergeTreeHeaderChunkMetadataDto> orderedChunkMetadata,
			string path)
		{
			if (element.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path} must be an array.");
			}

			int index = 0;
			foreach (JsonElement chunkMetadataElement in element.EnumerateArray())
			{
				string id;
				if (chunkMetadataElement.ValueKind == JsonValueKind.String)
				{
					id = chunkMetadataElement.GetString() ?? string.Empty;
				}
				else if (chunkMetadataElement.ValueKind == JsonValueKind.Object)
				{
					id = ReadRequiredStringProperty(chunkMetadataElement, "id", $"{path}[{index}]");
				}
				else
				{
					throw InvalidSnapshot($"Chunk metadata at {path}[{index}] must be an object or string.");
				}

				orderedChunkMetadata.Add(new MergeTreeHeaderChunkMetadataDto()
				{
					Id = id,
				});
				index++;
			}
		}

		private static void ReadCatchupOpsBlobNames(JsonElement element, List<string> blobNames, string path)
		{
			ReadOptionalBlobNameProperty(element, "catchupOps", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchUpOps", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchupOpsBlobNames", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchupOpsBlobName", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchupOpsBlobs", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchupOpsBlobIds", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchupOpsBlob", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchUpBlobNames", blobNames, path);
			ReadOptionalBlobNameProperty(element, "catchUpBlobName", blobNames, path);
		}

		private static void ReadOptionalBlobNameProperty(
			JsonElement element,
			string propertyName,
			List<string> blobNames,
			string path)
		{
			if (!element.TryGetProperty(propertyName, out JsonElement propertyElement)
				|| propertyElement.ValueKind == JsonValueKind.Null
				|| propertyElement.ValueKind == JsonValueKind.Undefined)
			{
				return;
			}

			if (propertyElement.ValueKind == JsonValueKind.String)
			{
				blobNames.Add(propertyElement.GetString() ?? string.Empty);
				return;
			}

			if (propertyElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path}.{propertyName} must be a string or array.");
			}

			int index = 0;
			foreach (JsonElement blobNameElement in propertyElement.EnumerateArray())
			{
				if (blobNameElement.ValueKind == JsonValueKind.String)
				{
					blobNames.Add(blobNameElement.GetString() ?? string.Empty);
				}
				else if (blobNameElement.ValueKind == JsonValueKind.Object)
				{
					blobNames.Add(ReadRequiredStringProperty(blobNameElement, "id", $"{path}.{propertyName}[{index}]"));
				}
				else
				{
					throw InvalidSnapshot($"Blob name at {path}.{propertyName}[{index}] must be a string or object.");
				}

				index++;
			}
		}

		private static MergeTree.PropertySet ReadPropertySet(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			if (element.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Property set at {path} must be a JSON object.");
			}

			MergeTree.PropertySet properties = new();
			foreach (JsonProperty property in element.EnumerateObject())
			{
				properties[property.Name] = JsonElementToObject(property.Value, $"{path}.{property.Name}", registry);
			}

			return properties;
		}

		private static object? JsonElementToObject(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			switch (element.ValueKind)
			{
				case JsonValueKind.Null:
				case JsonValueKind.Undefined:
					return null;

				case JsonValueKind.String:
					return element.GetString();

				case JsonValueKind.Number:
					if (element.TryGetInt64(out long longValue))
					{
						return longValue;
					}

					return element.GetDouble();

				case JsonValueKind.True:
					return true;

				case JsonValueKind.False:
					return false;

				case JsonValueKind.Object:
					if (HandleWireFormat.IsHandleShape(element))
					{
						return HandleWireFormat.ReadHandleFromShape(element, registry);
					}

					return ReadPropertySet(element, path, registry);

				case JsonValueKind.Array:
					List<object?> values = new();
					foreach (JsonElement item in element.EnumerateArray())
					{
						values.Add(JsonElementToObject(item, path, registry));
					}

					return values;

				default:
					throw InvalidSnapshot($"Unsupported JSON value in property set at {path}.");
			}
		}

		private static JsonElement ReadRequiredProperty(JsonElement element, string propertyName, string path)
		{
			if (!element.TryGetProperty(propertyName, out JsonElement propertyElement))
			{
				throw InvalidSnapshot($"JSON object at {path} must contain '{propertyName}'.");
			}

			return propertyElement;
		}

		private static string ReadRequiredStringProperty(JsonElement element, string propertyName, string path)
		{
			JsonElement propertyElement = ReadRequiredProperty(element, propertyName, path);
			if (propertyElement.ValueKind != JsonValueKind.String)
			{
				throw InvalidSnapshot($"JSON property {path}.{propertyName} must be a string.");
			}

			return propertyElement.GetString() ?? string.Empty;
		}

		private static int ReadRequiredIntProperty(JsonElement element, string propertyName, string path)
		{
			return ReadInt(ReadRequiredProperty(element, propertyName, path), $"{path}.{propertyName}");
		}

		private static long ReadRequiredLongProperty(JsonElement element, string propertyName, string path)
		{
			return ReadLong(ReadRequiredProperty(element, propertyName, path), $"{path}.{propertyName}");
		}

		private static long? ReadOptionalLongProperty(JsonElement element, string propertyName, string path)
		{
			if (!element.TryGetProperty(propertyName, out JsonElement propertyElement)
				|| propertyElement.ValueKind == JsonValueKind.Null
				|| propertyElement.ValueKind == JsonValueKind.Undefined)
			{
				return null;
			}

			return ReadLong(propertyElement, $"{path}.{propertyName}");
		}

		private static int? ReadOptionalIntProperty(JsonElement element, string propertyName, string path)
		{
			if (!element.TryGetProperty(propertyName, out JsonElement propertyElement)
				|| propertyElement.ValueKind == JsonValueKind.Null
				|| propertyElement.ValueKind == JsonValueKind.Undefined)
			{
				return null;
			}

			return ReadInt(propertyElement, $"{path}.{propertyName}");
		}

		private static string? ReadOptionalStringProperty(JsonElement element, string propertyName, string path)
		{
			if (!element.TryGetProperty(propertyName, out JsonElement propertyElement)
				|| propertyElement.ValueKind == JsonValueKind.Null
				|| propertyElement.ValueKind == JsonValueKind.Undefined)
			{
				return null;
			}

			return ReadStringLike(propertyElement, $"{path}.{propertyName}");
		}

		private static string? ReadOptionalRemovedClientId(JsonElement element, string path)
		{
			string? removedClientId = ReadOptionalStringProperty(element, "removedClientId", path)
				?? ReadOptionalStringProperty(element, "removedClient", path);
			if (removedClientId is not null)
			{
				return removedClientId;
			}

			if (!element.TryGetProperty("removedClientIds", out JsonElement removedClientIdsElement)
				|| removedClientIdsElement.ValueKind == JsonValueKind.Null
				|| removedClientIdsElement.ValueKind == JsonValueKind.Undefined)
			{
				return null;
			}

			if (removedClientIdsElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path}.removedClientIds must be an array.");
			}

			foreach (JsonElement item in removedClientIdsElement.EnumerateArray())
			{
				return ReadStringLike(item, $"{path}.removedClientIds[0]");
			}

			return null;
		}

		private static string? ReadOptionalMovedClientId(JsonElement element, string path)
		{
			string? movedClientId = ReadOptionalStringProperty(element, "movedClientId", path)
				?? ReadOptionalStringProperty(element, "movedClient", path);
			if (movedClientId is not null)
			{
				return movedClientId;
			}

			if (!element.TryGetProperty("movedClientIds", out JsonElement movedClientIdsElement)
				|| movedClientIdsElement.ValueKind == JsonValueKind.Null
				|| movedClientIdsElement.ValueKind == JsonValueKind.Undefined)
			{
				return null;
			}

			if (movedClientIdsElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path}.movedClientIds must be an array.");
			}

			foreach (JsonElement item in movedClientIdsElement.EnumerateArray())
			{
				return ReadStringLike(item, $"{path}.movedClientIds[0]");
			}

			return null;
		}

		private static void ReadRemoveStamps(JsonElement element, SharedStringSnapshotSegmentDto dto, string path)
		{
			List<string> removedClientIds = ReadOptionalStringArrayProperty(element, "removedClientIds", path);
			List<long> removedSeqs = ReadOptionalLongArrayProperty(element, "removedSeqs", path);
			if (removedClientIds.Count == 0)
			{
				string? removedClientId = ReadOptionalStringProperty(element, "removedClientId", path)
					?? ReadOptionalStringProperty(element, "removedClient", path);
				if (removedClientId is not null)
				{
					removedClientIds.Add(removedClientId);
				}
			}

			if (dto.RemovedSeq is long removedSeq && removedSeqs.Count == 0 && removedClientIds.Count > 0)
			{
				for (int i = 0; i < removedClientIds.Count; i++)
				{
					removedSeqs.Add(removedSeq);
				}
			}

			if (removedSeqs.Count > 0 || removedClientIds.Count > 0)
			{
				if (removedClientIds.Count != removedSeqs.Count)
				{
					throw InvalidSnapshot($"JSON properties {path}.removedClientIds and {path}.removedSeqs must have the same length.");
				}

				for (int i = 0; i < removedClientIds.Count; i++)
				{
					dto.RemoveStamps.Add(new SharedStringSnapshotRemoveStampDto()
					{
						Type = "setRemove",
						Seq = removedSeqs[i],
						ClientId = removedClientIds[i],
					});
				}
			}

			List<string> movedClientIds = ReadOptionalStringArrayProperty(element, "movedClientIds", path);
			List<long> movedSeqs = ReadOptionalLongArrayProperty(element, "movedSeqs", path);
			if (movedClientIds.Count == 0)
			{
				string? movedClientId = ReadOptionalStringProperty(element, "movedClientId", path)
					?? ReadOptionalStringProperty(element, "movedClient", path);
				if (movedClientId is not null)
				{
					movedClientIds.Add(movedClientId);
				}
			}

			if (dto.ObliteratedSeq is long movedSeq && movedSeqs.Count == 0 && movedClientIds.Count > 0)
			{
				for (int i = 0; i < movedClientIds.Count; i++)
				{
					movedSeqs.Add(movedSeq);
				}
			}

			if (movedSeqs.Count > 0 || movedClientIds.Count > 0)
			{
				if (movedClientIds.Count != movedSeqs.Count)
				{
					throw InvalidSnapshot($"JSON properties {path}.movedClientIds and {path}.movedSeqs must have the same length.");
				}

				for (int i = 0; i < movedClientIds.Count; i++)
				{
					dto.RemoveStamps.Add(new SharedStringSnapshotRemoveStampDto()
					{
						Type = "sliceRemove",
						Seq = movedSeqs[i],
						ClientId = movedClientIds[i],
					});
				}
			}

			dto.RemoveStamps.Sort((a, b) => a.Seq.CompareTo(b.Seq));
		}

		private static MergeTree.RemoveOperationStamp CreateSnapshotRemoveStamp(SharedStringSnapshotRemoveStampDto stamp, int clientId)
		{
			if (string.Equals(stamp.Type, "sliceRemove", StringComparison.Ordinal))
			{
				return new MergeTree.SliceRemoveOperationStamp()
				{
					Seq = stamp.Seq,
					ClientId = clientId,
				};
			}

			return new MergeTree.SetRemoveOperationStamp()
			{
				Seq = stamp.Seq,
				ClientId = clientId,
			};
		}

		private static bool HasSnapshotStampType(SharedStringSnapshotSegmentDto dto, string type)
		{
			foreach (SharedStringSnapshotRemoveStampDto stamp in dto.RemoveStamps)
			{
				if (string.Equals(stamp.Type, type, StringComparison.Ordinal))
				{
					return true;
				}
			}

			return false;
		}

		private static List<string> ReadOptionalStringArrayProperty(JsonElement element, string propertyName, string path)
		{
			List<string> values = new();
			if (!element.TryGetProperty(propertyName, out JsonElement arrayElement)
				|| arrayElement.ValueKind == JsonValueKind.Null
				|| arrayElement.ValueKind == JsonValueKind.Undefined)
			{
				return values;
			}

			if (arrayElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path}.{propertyName} must be an array.");
			}

			int index = 0;
			foreach (JsonElement item in arrayElement.EnumerateArray())
			{
				values.Add(ReadStringLike(item, $"{path}.{propertyName}[{index}]"));
				index++;
			}

			return values;
		}

		private static List<long> ReadOptionalLongArrayProperty(JsonElement element, string propertyName, string path)
		{
			List<long> values = new();
			if (!element.TryGetProperty(propertyName, out JsonElement arrayElement)
				|| arrayElement.ValueKind == JsonValueKind.Null
				|| arrayElement.ValueKind == JsonValueKind.Undefined)
			{
				return values;
			}

			if (arrayElement.ValueKind != JsonValueKind.Array)
			{
				throw InvalidSnapshot($"JSON property {path}.{propertyName} must be an array.");
			}

			int index = 0;
			foreach (JsonElement item in arrayElement.EnumerateArray())
			{
				values.Add(ReadLong(item, $"{path}.{propertyName}[{index}]"));
				index++;
			}

			return values;
		}

		private static int ReadInt(JsonElement element, string path)
		{
			long value = ReadLong(element, path);
			if (value < int.MinValue || value > int.MaxValue)
			{
				throw InvalidSnapshot($"JSON number at {path} must fit in Int32.");
			}

			return (int)value;
		}

		private static long ReadLong(JsonElement element, string path)
		{
			if (element.ValueKind == JsonValueKind.Number && element.TryGetInt64(out long longValue))
			{
				return longValue;
			}

			if (element.ValueKind == JsonValueKind.String)
			{
				string? stringValue = element.GetString();
				if (long.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsedValue))
				{
					return parsedValue;
				}
			}

			throw InvalidSnapshot($"JSON value at {path} must be an integer.");
		}

		private static string ReadStringLike(JsonElement element, string path)
		{
			if (element.ValueKind == JsonValueKind.String)
			{
				return element.GetString() ?? string.Empty;
			}

			if (element.ValueKind == JsonValueKind.Number)
			{
				return element.GetRawText();
			}

			throw InvalidSnapshot($"JSON value at {path} must be a string.");
		}

		private static void ValidateChunkShape(SharedStringSnapshotDto snapshot)
		{
			if (snapshot.SegmentCount != snapshot.Segments.Count)
			{
				throw InvalidSnapshot("SharedString snapshot segmentCount does not match segments length.");
			}

			int computedLength = ComputeSequenceLength(snapshot.Segments);
			if (snapshot.Length != computedLength)
			{
				throw InvalidSnapshot("SharedString snapshot length does not match segment length.");
			}
		}

		private static void ValidateChunkShape(MergeTreeChunkV1Dto chunk, string path)
		{
			if (chunk.SegmentCount != chunk.Segments.Count)
			{
				throw InvalidSnapshot($"SharedString snapshot {path}.segmentCount does not match segments length.");
			}

			int computedLength = ComputeSequenceLength(chunk.Segments);
			if (chunk.Length != computedLength)
			{
				throw InvalidSnapshot($"SharedString snapshot {path}.length does not match segment length.");
			}
		}

		private static void ValidateWholeSnapshotForPopulate(SharedStringSnapshotDto snapshot)
		{
			if (!string.Equals(snapshot.Version, "1", StringComparison.Ordinal))
			{
				throw InvalidSnapshot($"Unsupported SharedString snapshot version '{snapshot.Version}'.");
			}

			if (snapshot.HeaderMetadata is null)
			{
				throw InvalidSnapshot("SharedString snapshot headerMetadata is required for population.");
			}

			if (snapshot.StartIndex != 0)
			{
				throw InvalidSnapshot("SharedString snapshot population requires a header chunk with startIndex 0.");
			}

			ValidateChunkShape(snapshot);

			SharedStringSnapshotHeaderMetadata headerMetadata = snapshot.HeaderMetadata;
			if (headerMetadata.SequenceNumber < headerMetadata.MinSequenceNumber)
			{
				throw InvalidSnapshot("SharedString snapshot sequenceNumber must be greater than or equal to minSequenceNumber.");
			}

			if (headerMetadata.TotalSegmentCount != snapshot.SegmentCount
				|| headerMetadata.TotalLength != snapshot.Length)
			{
				throw InvalidSnapshot("SharedString snapshot totals do not match headerMetadata.");
			}
		}

		private static int ComputeSequenceLength(IEnumerable<SharedStringSnapshotSegmentDto> segments)
		{
			int length = 0;
			foreach (SharedStringSnapshotSegmentDto segment in segments)
			{
				if (segment is null)
				{
					throw InvalidSnapshot("Snapshot segments cannot contain null entries.");
				}

				checked
				{
					length += SegmentLength(segment);
				}
			}

			return length;
		}

		private static int SegmentLength(SharedStringSnapshotSegmentDto segment)
		{
			if (segment.IsUnknown)
			{
				return 0;
			}

			if (segment.IsMarker)
			{
				return 1;
			}

			return (segment.Text ?? string.Empty).Length;
		}

		private static OcsException InvalidSnapshot(string message)
		{
			return new OcsException(OcsGateErrorCode.InvalidOperation, message);
		}
	}
}
