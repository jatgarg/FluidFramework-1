// -----------------------------------------------------------------------------
// Regression tests for named telemetry events emitted by SharedString.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Linq;

using Microsoft.Office.Web.Fluid;
using Microsoft.Office.Web.Fluid.MergeTree;

using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringTelemetryTests
	{
		public SharedStringTelemetryTests()
		{
			// Reset the process-lifetime rate-limit counter so tests don't
			// interfere with each other.
			SharedString.ResetLocalOpReentryLogCountForTests();
		}

		[Fact]
		public void LocalOpReentry_FiresOnceOnReentrancy_WithDepthPayload()
		{
			var logger = new RecordingLogger();
			var sharedString = new SharedString(sender: null, logger: logger);
			sharedString.InsertText(0, "hello");
			sharedString.OnSequenceDelta += (_, _) =>
			{
				try { sharedString.InsertText(0, "X"); } catch (LoggingError) { }
			};

			// Trigger the reentrancy path.
			try { sharedString.InsertText(0, "Y"); } catch (LoggingError) { }

			var reentryEvent = Assert.Single(logger.Telemetry, e => e.EventName == "LocalOpReentry");
			Assert.NotNull(reentryEvent.Properties);
			Assert.True(reentryEvent.Properties!.ContainsKey("depth"));
			Assert.Equal(1, Assert.IsType<int>(reentryEvent.Properties["depth"]));
			Assert.NotNull(reentryEvent.Error);
			Assert.IsType<LoggingError>(reentryEvent.Error);
		}

		[Fact]
		public void LocalOpReentry_IsRateLimitedToThreePerProcess()
		{
			var logger = new RecordingLogger();
			var sharedString = new SharedString(sender: null, logger: logger);
			sharedString.InsertText(0, "hello");
			sharedString.OnSequenceDelta += (_, _) =>
			{
				try { sharedString.InsertText(0, "X"); } catch (LoggingError) { }
			};

			// Trigger reentrancy 5 times.
			for (int i = 0; i < 5; i++)
			{
				try { sharedString.InsertText(0, "Y"); } catch (LoggingError) { }
			}

			int reentryEventCount = logger.Telemetry.Count(e => e.EventName == "LocalOpReentry");
			Assert.Equal(3, reentryEventCount);
		}

		[Fact]
		public void SequenceLoadFailed_FiresOnSnapshotLoadException()
		{
			var logger = new RecordingLogger();
			var sharedString = new SharedString(sender: null, logger: logger);
			// Non-empty target — PopulateFromSnapshot rejects with LoggingError.
			sharedString.InsertText(0, "existing");
			SharedStringSnapshotDto snapshot = new()
			{
				Version = "1",
				SegmentCount = 0,
				Length = 0,
				StartIndex = 0,
				Segments = new(),
				HeaderMetadata = new()
				{
					SequenceNumber = 0,
					MinSequenceNumber = 0,
				},
			};

			LoggingError ex = Assert.Throws<LoggingError>(() => sharedString.LoadFromSnapshot(snapshot));

			var failedEvent = Assert.Single(logger.Errors, e => e.EventName == "SequenceLoadFailed");
			Assert.Same(ex, failedEvent.Error);
		}

		[Fact]
		public void CatchupOpsLoadFailure_FiresWhenCatchupOpMalformed()
		{
			var logger = new RecordingLogger();
			var sharedString = new SharedString(sender: null, logger: logger);
			// Snapshot with a malformed catchup op (missing sequenceNumber).
			SharedStringSnapshotDto snapshot = new()
			{
				Version = "1",
				SegmentCount = 0,
				Length = 0,
				StartIndex = 0,
				Segments = new(),
				HeaderMetadata = new()
				{
					SequenceNumber = 1,
					MinSequenceNumber = 0,
				},
				CatchupOps = new()
				{
					new CatchupOpDto()
					{
						// missing SequenceNumber → throws in ApplyCatchupOps
						ReferenceSequenceNumber = 1,
						MinimumSequenceNumber = 1,
						ClientId = "c1",
						OpJson = "{}",
					},
				},
			};

			Assert.Throws<LoggingError>(() => sharedString.LoadFromSnapshot(snapshot));

			// SequenceLoadFailed fires from the SharedString.LoadFromSnapshot outer
			// catch, and CatchupOpsLoadFailure fires from the inner ApplyCatchupOps
			// catch. Both should be present.
			Assert.Contains(logger.Errors, e => e.EventName == "CatchupOpsLoadFailure");
			Assert.Contains(logger.Errors, e => e.EventName == "SequenceLoadFailed");
		}
	}
}
