## ✅ Real Fluid fixtures

These fixtures are produced by real `@fluidframework/sequence` `SharedString.summarize()` running
against the `MockFluidDataStoreRuntime` from `@fluidframework/test-runtime-utils`. Wave 5's C#
snapshot loader reads the `.snapshot.bin` files and validates that it can parse Fluid's V1
merge-tree snapshot format and reconstruct the expected text described by each fixture.

## Regenerating

The generator lives at `generate.mjs` in this folder. It must run **from the sequence package
directory** so that `@fluidframework/sequence/internal` and
`@fluidframework/test-runtime-utils/internal` resolve:

```shell
cd /Users/jatingarg/Desktop/work/transpiledir/packages/dds/sequence
node src/csharp-port/SharedString.Tests/Fixtures/generate.mjs
```

Prerequisite: the Fluid TS workspace must be built. Run `pnpm install` from the repo root
followed by `pnpm run --filter=@fluidframework/sequence... build` first.

The generator has a fallback path (synthetic V1-lite blobs) that fires only when the Fluid
packages can't be resolved. For simple text-only inserts (no props on segments), the fallback
happens to produce the same output as real Fluid — segments are serialized as plain strings
when they carry no metadata. Trace lines (`[real Fluid]` vs `[fallback]`) at the top of the
generator's stdout show which path was taken.

## Format (real Fluid V1)

The `.snapshot.bin` files are raw UTF-8 bytes from the SharedString summary's `content/header`
blob, matching the `MergeTreeChunkV1` shape from
`packages/dds/merge-tree/src/snapshotChunks.ts`:

- `version: "1"`
- `segments: JsonSegmentSpecs[]` — each segment is either a plain string (for TextSegment
  with no properties) OR a JSON object like `{ text: "...", props: {...} }` for annotated
  segments, OR an `IJSONSegmentWithMergeInfo` wrapper with per-segment seq / clientId metadata
  when concurrent-edit state is present. Current fixtures use only the plain-string form
  because the ops applied are simple no-props text edits.
- `length` / `segmentCount`: chunk-local text length and segment count
- `startIndex`: the chunk's first segment index
- `headerMetadata`: `MergeTreeHeaderMetadata` with `minSequenceNumber`,
  `sequenceNumber`, `orderedChunkMetadata`, `totalLength`, `totalSegmentCount`

The C# loader handles all three segment forms (see `SharedStringSnapshotLoader.cs` Wave 5
implementation). The `.snapshot.json.txt` files contain the same bytes decoded and
pretty-printed only to make PR review easier.

