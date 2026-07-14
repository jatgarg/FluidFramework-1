// -----------------------------------------------------------------------------
// Parses SharedDirectory snapshot JSON (IDirectoryDataObject in TS) into an
// intermediate DirectorySnapshotDto for later application by Wave 2b.
// Supports the simple IDirectoryDataObject shape and the blob-split
// IDirectoryNewStorageFormat load path.
// Part of the SharedDirectory C# feasibility port — Wave 2a.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Microsoft.Office.Web.Fluid
{
	public sealed class DirectorySnapshotDto
	{
		/// <summary>Key-value pairs at this directory level. Value is ISerializableValue-style snapshot data.</summary>
		public Dictionary<string, SerializedValue> Storage { get; set; } = new();

		/// <summary>Child subdirectories keyed by name. Recursive.</summary>
		public Dictionary<string, DirectorySnapshotDto> Subdirectories { get; set; } = new();

		/// <summary>Subdirectory creation metadata used to identify this incarnation.</summary>
		public DirectoryCreateInfo? CreateInfo { get; set; }
	}

	public sealed class DirectoryCreateInfo
	{
		/// <summary>Sequence number at which this subdirectory was created.</summary>
		public long Csn { get; set; }

		/// <summary>Client IDs that created this subdirectory incarnation.</summary>
		public string[] CcIds { get; set; } = Array.Empty<string>();
	}

	/// <summary>C# counterpart to TS IDirectoryNewStorageFormat.</summary>
	public sealed class IDirectoryNewStorageFormat
	{
		/// <summary>Blob IDs representing larger directory data that was serialized separately.</summary>
		[JsonPropertyName("blobs")]
		public string[] Blobs { get; set; } = Array.Empty<string>();

		/// <summary>Main IDirectoryDataObject-equivalent content that remains in the snapshot blob.</summary>
		[JsonPropertyName("content")]
		public DirectorySnapshotDto Content { get; set; } = new();
	}

	public static class DirectorySnapshotLoader
	{
		/// <summary>Loads a directory snapshot JSON into an intermediate DTO tree.</summary>
		public static DirectorySnapshotDto Load(string json, Func<string, string>? blobResolver = null, IFluidDataObjectRegistry? registry = null)
		{
			return Parse(json, blobResolver, registry);
		}

		/// <summary>Loads a directory snapshot JSON (UTF-8 bytes) into an intermediate DTO tree.</summary>
		public static DirectorySnapshotDto Load(ReadOnlySpan<byte> jsonBytes, Func<string, string>? blobResolver = null, IFluidDataObjectRegistry? registry = null)
		{
			return Parse(jsonBytes, blobResolver, registry);
		}

		/// <summary>Parses a directory snapshot JSON into an intermediate DTO tree.</summary>
		public static DirectorySnapshotDto Parse(string json, Func<string, string>? blobResolver = null, IFluidDataObjectRegistry? registry = null)
		{
			if (json == null)
			{
				throw new ArgumentNullException(nameof(json));
			}

			byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
			return Parse(jsonBytes, blobResolver, registry);
		}

		/// <summary>Parses a directory snapshot JSON (UTF-8 bytes) into an intermediate DTO tree.</summary>
		public static DirectorySnapshotDto Parse(ReadOnlySpan<byte> jsonBytes, Func<string, string>? blobResolver = null, IFluidDataObjectRegistry? registry = null)
		{
			var reader = new Utf8JsonReader(jsonBytes);
			return ReadFrom(ref reader, blobResolver, registry);
		}

		/// <summary>Reads a directory snapshot from an open reader.</summary>
		public static DirectorySnapshotDto ReadFrom(ref Utf8JsonReader reader, Func<string, string>? blobResolver = null, IFluidDataObjectRegistry? registry = null)
		{
			using JsonDocument document = JsonDocument.ParseValue(ref reader);
			if (IsBlobSplitFormat(document.RootElement, out JsonElement blobsElement))
			{
				return ReadBlobSplitFormat(document.RootElement, blobsElement, blobResolver, registry);
			}

			return ReadDirectory(document.RootElement, "root", registry);
		}

		private static DirectorySnapshotDto ReadBlobSplitFormat(
			JsonElement element,
			JsonElement blobsElement,
			Func<string, string>? blobResolver,
			IFluidDataObjectRegistry? registry)
		{
			if (blobResolver == null)
			{
				throw InvalidSnapshot("Blob-split IDirectoryNewStorageFormat snapshots require a blob resolver callback.");
			}

			IDirectoryNewStorageFormat format = ReadNewStorageFormat(element, blobsElement, registry);
			DirectorySnapshotDto snapshot = format.Content;

			foreach (string blobName in format.Blobs)
			{
				string? blobJson = blobResolver(blobName);
				if (blobJson == null)
				{
					throw InvalidSnapshot($"Blob resolver returned null for blob '{blobName}'.");
				}

				DirectorySnapshotDto blobContent = ParseDirectoryDataObjectJson(blobJson, $"blob '{blobName}'", registry);
				MergeSnapshotInto(snapshot, blobContent);
			}

			return snapshot;
		}

		private static IDirectoryNewStorageFormat ReadNewStorageFormat(JsonElement element, JsonElement blobsElement, IFluidDataObjectRegistry? registry)
		{
			if (!element.TryGetProperty("content", out JsonElement contentElement))
			{
				throw InvalidSnapshot("Blob-split IDirectoryNewStorageFormat snapshot must contain a 'content' object.");
			}

			return new IDirectoryNewStorageFormat()
			{
				Blobs = ReadBlobNames(blobsElement),
				Content = ReadDirectory(contentElement, "root.content", registry),
			};
		}

		private static string[] ReadBlobNames(JsonElement blobsElement)
		{
			var blobNames = new List<string>();
			int index = 0;
			foreach (JsonElement blobNameElement in blobsElement.EnumerateArray())
			{
				if (blobNameElement.ValueKind != JsonValueKind.String)
				{
					throw InvalidSnapshot($"Blob name at root.blobs[{index}] must be a string.");
				}

				blobNames.Add(blobNameElement.GetString() ?? string.Empty);
				index++;
			}

			return blobNames.ToArray();
		}

		private static DirectorySnapshotDto ParseDirectoryDataObjectJson(string json, string path, IFluidDataObjectRegistry? registry)
		{
			using JsonDocument document = JsonDocument.Parse(json);
			return ReadDirectory(document.RootElement, path, registry);
		}

		private static void MergeSnapshotInto(DirectorySnapshotDto target, DirectorySnapshotDto source)
		{
			// Matches TS loadCore's sequential populate(...) calls: every fragment merges
			// into the same root and later storage values override earlier ones.
			foreach (KeyValuePair<string, SerializedValue> kvp in source.Storage)
			{
				target.Storage[kvp.Key] = kvp.Value;
			}

			foreach (KeyValuePair<string, DirectorySnapshotDto> kvp in source.Subdirectories)
			{
				if (target.Subdirectories.TryGetValue(kvp.Key, out DirectorySnapshotDto? targetChild))
				{
					MergeSnapshotInto(targetChild, kvp.Value);
				}
				else
				{
					target.Subdirectories[kvp.Key] = kvp.Value;
				}
			}
		}

		private static bool IsBlobSplitFormat(JsonElement element, out JsonElement blobsElement)
		{
			if (element.ValueKind == JsonValueKind.Object
				&& element.TryGetProperty("blobs", out blobsElement)
				&& blobsElement.ValueKind == JsonValueKind.Array)
			{
				return true;
			}

			blobsElement = default;
			return false;
		}

		private static DirectorySnapshotDto ReadDirectory(JsonElement element, string path, IFluidDataObjectRegistry? registry)
		{
			if (element.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Directory snapshot at {path} must be a JSON object.");
			}

			var dto = new DirectorySnapshotDto();

			if (element.TryGetProperty("storage", out JsonElement storageElement))
			{
				ReadStorage(storageElement, dto.Storage, path, registry);
			}

			if (element.TryGetProperty("subdirectories", out JsonElement subdirectoriesElement))
			{
				ReadSubdirectories(subdirectoriesElement, dto.Subdirectories, path, registry);
			}

			if (element.TryGetProperty("ci", out JsonElement createInfoElement))
			{
				dto.CreateInfo = ReadCreateInfo(createInfoElement, path);
			}

			return dto;
		}

		private static DirectoryCreateInfo ReadCreateInfo(JsonElement createInfoElement, string path)
		{
			if (createInfoElement.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Create info at {path}.ci must be a JSON object.");
			}

			if (!createInfoElement.TryGetProperty("csn", out JsonElement csnElement)
				|| csnElement.ValueKind != JsonValueKind.Number
				|| !csnElement.TryGetInt64(out long csn))
			{
				throw InvalidSnapshot($"Create info at {path}.ci must contain a numeric 'csn' field.");
			}

			string[] ccIds = Array.Empty<string>();
			if (createInfoElement.TryGetProperty("ccIds", out JsonElement ccIdsElement))
			{
				if (ccIdsElement.ValueKind != JsonValueKind.Array)
				{
					throw InvalidSnapshot($"Create info at {path}.ci.ccIds must be an array.");
				}

				var clientIds = new List<string>();
				int index = 0;
				foreach (JsonElement clientIdElement in ccIdsElement.EnumerateArray())
				{
					if (clientIdElement.ValueKind != JsonValueKind.String)
					{
						throw InvalidSnapshot($"Create info client id at {path}.ci.ccIds[{index}] must be a string.");
					}

					clientIds.Add(clientIdElement.GetString() ?? string.Empty);
					index++;
				}

				ccIds = clientIds.ToArray();
			}

			return new DirectoryCreateInfo()
			{
				Csn = csn,
				CcIds = ccIds,
			};
		}

		private static void ReadStorage(
			JsonElement storageElement,
			Dictionary<string, SerializedValue> storage,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			if (storageElement.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Storage at {path} must be a JSON object.");
			}

			foreach (JsonProperty storageProperty in storageElement.EnumerateObject())
			{
				string key = storageProperty.Name;
				JsonElement serializedElement = storageProperty.Value;
				if (serializedElement.ValueKind != JsonValueKind.Object)
				{
					throw InvalidSnapshot($"Storage value '{key}' at {path} must be an object with 'type' and 'value' fields.");
				}

				if (!serializedElement.TryGetProperty("type", out JsonElement typeElement)
					|| typeElement.ValueKind != JsonValueKind.String)
				{
					throw InvalidSnapshot($"Storage value '{key}' at {path} must contain a string 'type' field.");
				}

				object? value = null;
				if (serializedElement.TryGetProperty("value", out JsonElement valueElement))
				{
					value = DirectoryOpSerializer.ReadJsonValueWithHandles(valueElement, registry);
				}

				storage[key] = new SerializedValue()
				{
					Type = typeElement.GetString() ?? string.Empty,
					Value = value,
				};
			}
		}

		private static void ReadSubdirectories(
			JsonElement subdirectoriesElement,
			Dictionary<string, DirectorySnapshotDto> subdirectories,
			string path,
			IFluidDataObjectRegistry? registry)
		{
			if (subdirectoriesElement.ValueKind != JsonValueKind.Object)
			{
				throw InvalidSnapshot($"Subdirectories at {path} must be a JSON object.");
			}

			foreach (JsonProperty subdirectoryProperty in subdirectoriesElement.EnumerateObject())
			{
				string subdirectoryName = subdirectoryProperty.Name;
				string subdirectoryPath = path == "root"
					? subdirectoryName
					: $"{path}/{subdirectoryName}";
				subdirectories[subdirectoryName] = ReadDirectory(subdirectoryProperty.Value, subdirectoryPath, registry);
			}
		}

		private static OcsException InvalidSnapshot(string message)
		{
			return new OcsException(OcsGateErrorCode.InvalidOperation, message);
		}
	}
}
