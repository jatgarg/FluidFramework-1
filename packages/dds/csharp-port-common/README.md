# csharp-port-common

> Shared throwaway shims used by every DDS C# port in this repo
> (SharedDirectory, SharedString, ...). Deleted on transfer to the host repo.

Types the host repo already provides (from
`waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/FluidDispatcher.cs`
and friends):

- `IFluidDataObject`
- `IFluidDataObjectMessageHandler`
- `IFluidDataObjectSender`
- `IFluidDataObjectRegistry`
- `SequenceNumber`
- `OpOrigin`
- `SequencedDocumentMessageDescriptor`
- `FluidObjectId`
- `OcsException` + `OcsGateErrorCode`
- `fluidDataStoreMessageAttach` (attach-message shim)

## Why extract these

The SharedDirectory port originally had these in its own `cs-shims/` folder.
When we started SharedString, we duplicated them there too — bad. Extracted to
this shared library so both ports (and future DDS ports) reference the same
shim definitions.

## Visibility choices

Fields on `SequenceNumber` and `SequencedDocumentMessageDescriptor` are `public`
(vs `internal` in the host repo). Reason: shim consumers span multiple
assemblies (`SharedDirectory`, `SharedString`, their test projects). Public
avoids `[InternalsVisibleTo]` proliferation. When the host takes the code, the
host's `internal` visibility on real types takes over — no code impact.

## Transfer

On transfer to the host repo:
1. Delete this entire project.
2. The host repo's real definitions (in `FluidDispatcher.cs`, `FluidObjectId.cs`,
   etc.) take over. Consumer code compiles unchanged because we matched the
   host's public interface shapes exactly.

## Build

Referenced via `<ProjectReference>` from each DDS port csproj:

```xml
<ProjectReference Include="..\..\..\..\..\csharp-port-common\CsharpPortCommon.csproj" />
```

(from `packages/dds/<pkg>/src/csharp-port/<Port>/<Port>.csproj`)
