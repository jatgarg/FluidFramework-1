/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { bufferToString, stringToBuffer } from "@fluid-internal/client-utils";
import type { Jsonable } from "@fluidframework/datastore-definitions/internal";
import { compress, decompress } from "lz4js";

import { type Encoder } from "./encoders.js";

// TODO: document this
/**
 * @internal
 */
export function makeLZ4Encoder<T>(): Encoder<Jsonable<T>, string> {
	return {
		encode: (decoded: Jsonable<T>): string => {
			const uncompressed = new TextEncoder().encode(JSON.stringify(decoded));
			const compressed = compress(uncompressed) as ArrayBufferLike;
			return bufferToString(compressed, "base64");
		},
		decode: (serializedSummary: string): Jsonable<T> => {
			const compressed = new Uint8Array(stringToBuffer(serializedSummary, "base64"));
			const uncompressed = decompress(compressed) as BufferSource;
			const decoded = JSON.parse(new TextDecoder().decode(uncompressed)) as Jsonable<T>;
			return decoded;
		},
	};
}