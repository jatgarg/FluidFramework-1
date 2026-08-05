// -----------------------------------------------------------------------------
// Wave 2c tests for DirectoryOpSerializer. Structure-inspired by TS
// packages/dds/map/src/test/mocha/directory.spec.ts; test bodies target
// the C# port's Option B pending-change model.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Text.Json;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class DirectoryOpSerializerTests
	{
		[Fact]
		public void SetOp_RoundTripsPlainObjectValue()
		{
			var op = new DirectorySetOperation()
			{
				Path = "/foo",
				Key = "profile",
				Value = new SerializableValue()
				{
					Type = "Plain",
					Value = new Dictionary<string, object?>()
					{
						["name"] = "Ada",
						["count"] = 42,
						["enabled"] = true,
						["tags"] = new[] { "math", "logic" },
					},
				},
			};

			string json = DirectoryOpSerializer.Serialize(op);
			Assert.StartsWith("{\"type\":\"set\"", json);

			using JsonDocument document = JsonDocument.Parse(json);
			JsonElement root = document.RootElement;
			Assert.Equal("set", root.GetProperty("type").GetString());
			Assert.Equal("/foo", root.GetProperty("path").GetString());
			Assert.Equal("profile", root.GetProperty("key").GetString());
			JsonElement wireValue = root.GetProperty("value");
			Assert.Equal("Plain", wireValue.GetProperty("type").GetString());
			Assert.Equal("Ada", wireValue.GetProperty("value").GetProperty("name").GetString());
			Assert.Equal(42, wireValue.GetProperty("value").GetProperty("count").GetInt32());
			Assert.True(wireValue.GetProperty("value").GetProperty("enabled").GetBoolean());

			DirectorySetOperation roundTripped = Assert.IsType<DirectorySetOperation>(DirectoryOpSerializer.Deserialize(json));
			Assert.Equal(DirectoryOpType.Set, roundTripped.Type);
			Assert.Equal("/foo", roundTripped.Path);
			Assert.Equal("profile", roundTripped.Key);
			Assert.Equal("Plain", roundTripped.Value.Type);
			// TS-parity: JSON values now materialize to native types (Finding 18).
			Dictionary<string, object?> roundTrippedValue = Assert.IsType<Dictionary<string, object?>>(roundTripped.Value.Value);
			Assert.Equal("Ada", Assert.IsType<string>(roundTrippedValue["name"]));
			Assert.Equal(42d, Assert.IsType<double>(roundTrippedValue["count"]));
			Assert.True(Assert.IsType<bool>(roundTrippedValue["enabled"]));
			List<object?> tags = Assert.IsType<List<object?>>(roundTrippedValue["tags"]);
			Assert.Equal("logic", Assert.IsType<string>(tags[1]));
		}

		[Fact]
		public void DeleteOp_RoundTripsFields()
		{
			var op = new DirectoryDeleteOperation()
			{
				Path = "/foo",
				Key = "profile",
			};

			string json = DirectoryOpSerializer.Serialize(op);
			DirectoryDeleteOperation roundTripped = Assert.IsType<DirectoryDeleteOperation>(DirectoryOpSerializer.Deserialize(json));

			Assert.Equal(DirectoryOpType.Delete, roundTripped.Type);
			Assert.Equal("/foo", roundTripped.Path);
			Assert.Equal("profile", roundTripped.Key);
		}

		[Fact]
		public void ClearOp_RoundTripsFields()
		{
			var op = new DirectoryClearOperation()
			{
				Path = "/foo/bar",
			};

			string json = DirectoryOpSerializer.Serialize(op);
			DirectoryClearOperation roundTripped = Assert.IsType<DirectoryClearOperation>(DirectoryOpSerializer.Deserialize(json));

			Assert.Equal(DirectoryOpType.Clear, roundTripped.Type);
			Assert.Equal("/foo/bar", roundTripped.Path);
		}

		[Fact]
		public void CreateSubDirectoryOp_RoundTripsFields()
		{
			var op = new DirectoryCreateSubDirectoryOperation()
			{
				Path = "/foo",
				SubdirName = "bar",
			};

			string json = DirectoryOpSerializer.Serialize(op);
			DirectoryCreateSubDirectoryOperation roundTripped = Assert.IsType<DirectoryCreateSubDirectoryOperation>(DirectoryOpSerializer.Deserialize(json));

			Assert.Equal(DirectoryOpType.CreateSubDirectory, roundTripped.Type);
			Assert.Equal("/foo", roundTripped.Path);
			Assert.Equal("bar", roundTripped.SubdirName);
		}

		[Fact]
		public void DeleteSubDirectoryOp_RoundTripsFields()
		{
			var op = new DirectoryDeleteSubDirectoryOperation()
			{
				Path = "/foo",
				SubdirName = "bar",
			};

			string json = DirectoryOpSerializer.Serialize(op);
			DirectoryDeleteSubDirectoryOperation roundTripped = Assert.IsType<DirectoryDeleteSubDirectoryOperation>(DirectoryOpSerializer.Deserialize(json));

			Assert.Equal(DirectoryOpType.DeleteSubDirectory, roundTripped.Type);
			Assert.Equal("/foo", roundTripped.Path);
			Assert.Equal("bar", roundTripped.SubdirName);
		}

		[Fact]
		public void Deserialize_UnknownOpType_ThrowsUnknownOp()
		{
			OcsException exception = Assert.Throws<OcsException>(
				() => DirectoryOpSerializer.Deserialize("{\"type\":\"whatever\",\"path\":\"/\"}"));

			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		// -----------------------------------------------------------------
		// Strict-deserialization regression tests.
		// TS ref: packages/dds/map/src/directory.ts — IDirectorySetOperation,
		// IDirectoryDeleteOperation, IDirectoryClearOperation,
		// IDirectoryCreateSubDirectoryOperation, IDirectoryDeleteSubDirectoryOperation.
		// All fields on those interfaces are required; the C# deserializer must
		// reject wire ops missing them rather than defaulting to "".
		// -----------------------------------------------------------------

		[Fact]
		public void Deserialize_MissingType_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"path\":\"/\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":1}}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_MissingPath_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"set\",\"key\":\"k\",\"value\":{\"type\":\"Plain\",\"value\":1}}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_SetOp_MissingKey_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"set\",\"path\":\"/\",\"value\":{\"type\":\"Plain\",\"value\":1}}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_SetOp_MissingValue_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\"}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_SetOp_ValueMissingType_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"set\",\"path\":\"/\",\"key\":\"k\",\"value\":{\"value\":1}}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_DeleteOp_MissingKey_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"delete\",\"path\":\"/\"}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_CreateSubDirectoryOp_MissingSubdirName_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"createSubDirectory\",\"path\":\"/\"}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_DeleteSubDirectoryOp_MissingSubdirName_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"deleteSubDirectory\",\"path\":\"/\"}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}

		[Fact]
		public void Deserialize_ClearOp_MissingPath_Throws()
		{
			OcsException exception = Assert.Throws<OcsException>(() => DirectoryOpSerializer.Deserialize("{\"type\":\"clear\"}"));
			Assert.Equal(OcsGateErrorCode.UnknownOp, exception.ErrorCode);
		}
	}
}
