# csharp-port-common

> Shared library used by every DDS C# port in this repo (SharedDirectory,
> SharedString, ...). Mostly throwaway shims — but two files are shared real
> deliverables that carry over on transfer.

## Contents

### Throwaway shims (deleted on transfer)

These mirror types the host repo already provides
(`waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/FluidDispatcher.cs`
and friends). Removed at transfer time; waccobalt's real types take over at
the same names.

- `IFluidDataObject`
- `IFluidDataObjectMessageHandler`
- `IFluidDataObjectSender` (with a `LocalClientId` addition — see note below)
- `IFluidDataObjectRegistry`
- `SequenceNumber`
- `OpOrigin`
- `SequencedDocumentMessageDescriptor` (extended with `RefSeq` + `ClientId`)
- `FluidObjectId`
- `OcsException` + `OcsGateErrorCode`
- `fluidDataStoreMessageAttach` (attach-message shim)

### Real deliverables (carry over on transfer)

Shared across every DDS port; used to detect + emit the Fluid handle wire
shape (`{"type":"__fluid_handle__","url":"…"}`).

- **`SerializedFluidHandle.cs`** — tiny POCO used on ingress when a handle URL
  can't be resolved to a live `IFluidDataObject` (e.g. no registry available).
- **`HandleWireFormat.cs`** — value-walker helpers: `IsHandleShape`,
  `ReadHandleFromShape`, `WriteHandleShape`. Used by SharedDirectory's
  `DirectoryOpSerializer` + SharedString's `SharedStringOpSerializer`.

Both types are referenced by DDS-port code that carries over verbatim. On
transfer they should be moved into whichever waccobalt assembly makes sense
(or replaced by real implementations if waccobalt already has equivalents —
check first before assuming).

## Notes for reviewers coming from Fluid TS

- **`IFluidDataObjectRegistry`** is waccobalt's name for what Fluid TS calls
  `IFluidHandleContext` — URL ↔ object resolution. Fluid TS uses "registry"
  for a different concept (factory registration). See the docstring on the
  interface for the full note.
- **`IFluidDataObjectSender.LocalClientId`** is a nullable string added so
  local subdirectory creates can seed the creator client id at pending-time
  (matches TS `runtime.clientId ?? "detached"` behavior). Waccobalt may need
  a similar exposure on its real sender.

## Why extract these

The SharedDirectory port originally had these in its own `cs-shims/` folder.
When we started SharedString, we duplicated them there too — bad. Extracted
to this shared library so both ports (and future DDS ports) reference the
same definitions.

## Visibility choices

Fields on `SequenceNumber` and `SequencedDocumentMessageDescriptor` are `public`
(vs `internal` in the host repo). Reason: shim consumers span multiple
assemblies (`SharedDirectory`, `SharedString`, their test projects). Public
avoids `[InternalsVisibleTo]` proliferation. When the host takes the code, the
host's `internal` visibility on real types takes over — no code impact.

## Transfer

On transfer to the host repo:
1. Delete the throwaway shim files (list above).
2. Move `SerializedFluidHandle.cs` + `HandleWireFormat.cs` to a shared
   waccobalt location — or replace with waccobalt equivalents if any exist.
3. The host repo's real definitions for the removed types (in `FluidDispatcher.cs`,
   `FluidObjectId.cs`, etc.) take over. Consumer code compiles unchanged
   because we matched the host's public interface shapes exactly.

## Build

Referenced via `<ProjectReference>` from each DDS port csproj:

```xml
<ProjectReference Include="..\..\..\..\..\csharp-port-common\CsharpPortCommon.csproj" />
```

(from `packages/dds/<pkg>/src/csharp-port/<Port>/<Port>.csproj`)
