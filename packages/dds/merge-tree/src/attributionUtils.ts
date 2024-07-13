/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { assert } from "@fluidframework/core-utils/internal";
import type { CustomAttributionKey } from "@fluidframework/runtime-definitions/internal";

import type { PropertySet } from "./properties.js";

/**
 * @legacy
 * @alpha
 */
export const customAttributionKeysPropName: string =
	"CAK-cf9b6fe4-4c50-4a5d-9045-eb73b886f740";

/**
 * @legacy
 * @alpha
 */
export interface ICustomAttributionKeyList {
	type: "custom";
	keys: { offset: number; key: CustomAttributionKey }[];
}

export function isCustomAttributionKeyList(obj: unknown): obj is ICustomAttributionKeyList {
	return (
		obj !== null &&
		typeof obj === "object" &&
		"type" in obj &&
		"keys" in obj &&
		obj.type === "custom" &&
		Array.isArray(obj.keys)
	);
}

/**
 * Utility api to add custom key prop in the property set.
 * @param props - property set into which the  custom attribution keys will be inserted.
 * @param offsets - list of offset at which corresponding keys are provided.
 * @param keys - list of keys generated by the attributor.
 * @legacy
 * @alpha
 */
export function insertCustomAttributionPropInPropertySet(
	props: PropertySet,
	offsets: number[],
	keys: CustomAttributionKey[],
) {
	if (keys.length === 0) {
		return;
	}
	assert(offsets.length === keys.length, "offsets and keys length should be same");
	const attributionKeyList: ICustomAttributionKeyList = { type: "custom", keys: [] };
	for (let i = 0; i < offsets.length; i++) {
		// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
		attributionKeyList.keys.push({ offset: offsets[i]!, key: keys[i]! });
	}
	props[customAttributionKeysPropName] = attributionKeyList;
}
