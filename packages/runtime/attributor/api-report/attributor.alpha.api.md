## Alpha API Report File for "@fluidframework/attributor"

> Do not edit this file. It is a report generated by [API Extractor](https://api-extractor.com/).

```ts

import { AttributionInfo } from '@fluidframework/runtime-definitions/internal';
import { AttributionKey } from '@fluidframework/runtime-definitions/internal';
import { IDeltaManager } from '@fluidframework/container-definitions/internal';
import { IDocumentMessage } from '@fluidframework/driver-definitions/internal';
import { IQuorumClients } from '@fluidframework/driver-definitions/internal';
import { ISequencedDocumentMessage } from '@fluidframework/driver-definitions/internal';
import { ISnapshotTree } from '@fluidframework/driver-definitions/internal';
import { ISummaryTreeWithStats } from '@fluidframework/runtime-definitions/internal';
import type { Jsonable } from '@fluidframework/datastore-definitions/internal';

// @alpha (undocumented)
export const enableOnNewFileKey = "Fluid.Attribution.EnableOnNewFile";

// @alpha (undocumented)
export interface IProvideRuntimeAttributor {
    // (undocumented)
    readonly IRuntimeAttributor: IRuntimeAttributor;
}

// @alpha (undocumented)
export const IRuntimeAttributor: keyof IProvideRuntimeAttributor;

// @alpha
export interface IRuntimeAttributor extends IProvideRuntimeAttributor {
    // (undocumented)
    get(key: AttributionKey): AttributionInfo;
    // (undocumented)
    has(key: AttributionKey): boolean;
    // (undocumented)
    readonly isEnabled: boolean;
}

// (No @packageDocumentation comment for this package)

```