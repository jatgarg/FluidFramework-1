// -----------------------------------------------------------------------------
// Wave 8 tests for SharedString POC.
// -----------------------------------------------------------------------------

using System.Collections.Generic;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public class SharedStringBasicsTests
	{
		[Fact]
		public void Ctor_AssignsId_WhenNotProvided()
		{
			var sharedString = new SharedString();

			Assert.False(string.IsNullOrEmpty(sharedString.Id));
		}

		[Fact]
		public void Ctor_AcceptsExplicitId()
		{
			var sharedString = new SharedString("test-id");

			Assert.Equal("test-id", sharedString.Id);
		}

		[Fact]
		public void EmptyString_HasZeroLength()
		{
			var sharedString = new SharedString();

			Assert.Equal(0, sharedString.GetLength());
			Assert.Equal(string.Empty, sharedString.GetText());
		}

		[Fact]
		public void InsertText_At0_AppendsAndReadsBack()
		{
			var sharedString = new SharedString();

			sharedString.InsertText(0, "hello");

			Assert.Equal("hello", sharedString.GetText());
			Assert.Equal(5, sharedString.GetLength());
		}

		[Fact]
		public void InsertText_AtMiddle_SplitsCorrectly()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.InsertText(5, " beautiful");

			Assert.Equal("hello beautiful world", sharedString.GetText());
		}

		[Fact]
		public void InsertText_AtEnd_Appends()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");

			sharedString.InsertText(sharedString.GetLength(), " world");

			Assert.Equal("hello world", sharedString.GetText());
		}

		[Fact]
		public void GetText_WithRange_ReturnsSubstring()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			Assert.Equal("hello", sharedString.GetText(0, 5));
			Assert.Equal("world", sharedString.GetText(6, 11));
		}

		[Fact]
		public void DeleteText_RemovesRange()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");

			sharedString.DeleteText(5, 6);

			Assert.Equal("helloworld", sharedString.GetText());
			Assert.Equal(10, sharedString.GetLength());
		}

		[Fact]
		public void DeleteText_FullContent_LeavesEmpty()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello");

			sharedString.DeleteText(0, sharedString.GetLength());

			Assert.Equal(string.Empty, sharedString.GetText());
			Assert.Equal(0, sharedString.GetLength());
		}

		[Fact]
		public void InsertText_MultipleSequentially_Concatenates()
		{
			var sharedString = new SharedString();

			sharedString.InsertText(0, "one");
			sharedString.InsertText(3, " two");
			sharedString.InsertText(7, " three");

			Assert.Equal("one two three", sharedString.GetText());
		}

		[Fact]
		public void OnSequenceDelta_FiresOnLocalInsert_WithLocalTrue()
		{
			var sharedString = new SharedString();
			SequenceDeltaEventArgs? captured = null;
			sharedString.OnSequenceDelta += (s, e) => captured = e;

			sharedString.InsertText(0, "hello");

			Assert.NotNull(captured);
			Assert.True(captured!.Local);
			Assert.Equal("insert", captured.OpType);
			Assert.Equal(0, captured.Position);
			Assert.Equal(5, captured.Length);
			Assert.Equal("hello", captured.Text);
		}

		[Fact]
		public void OnSequenceDelta_FiresOnLocalDelete_WithLocalTrue()
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, "hello world");
			var events = new List<SequenceDeltaEventArgs>();
			sharedString.OnSequenceDelta += (s, e) => events.Add(e);

			sharedString.DeleteText(5, 6);

			SequenceDeltaEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Equal("remove", captured.OpType);
			Assert.Equal(5, captured.Position);
			Assert.Equal(1, captured.Length);
			Assert.Null(captured.Text);
		}

		[Fact]
		public void InsertText_WithSender_EmitsOp()
		{
			var sender = new FakeFluidDataObjectSender();
			var sharedString = new SharedString("shared-string", sender);

			sharedString.InsertText(0, "hello");

			var sent = Assert.Single(sender.Sent);
			Assert.Equal("shared-string", sent.Address);
			Assert.Equal("insert", sent.OpTypeName);
			Assert.False(string.IsNullOrEmpty(sent.OpJson));
			Assert.Equal(1, sent.ClientSeq);
		}

		[Fact]
		public void InsertText_WithoutSender_DoesNotEmit()
		{
			var sharedString = new SharedString(sender: null);

			sharedString.InsertText(0, "hello");

			Assert.Equal("hello", sharedString.GetText());
		}
	}
}
