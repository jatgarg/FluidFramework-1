/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { type IFluidLoadable } from "@fluidframework/core-interfaces";
import { type IChannelFactory } from "@fluidframework/datastore-definitions";
import {
	type IFluidDataStoreFactory,
	type NamedFluidDataStoreRegistryEntry,
} from "@fluidframework/runtime-definitions";
import type { ISharedObjectKind } from "@fluidframework/shared-object-base";
import { UsageError } from "@fluidframework/telemetry-utils";

import { type ContainerSchema, type DataObjectClass, type LoadableObjectClass } from "./types.js";

/**
 * An internal type used by the internal type guard isDataObjectClass to cast a
 * DataObjectClass to a type that is strongly coupled to IFluidDataStoreFactory.
 * Unlike the external and exported type DataObjectClass which is
 * weakly coupled to the IFluidDataStoreFactory to prevent leaking internals.
 */
export type InternalDataObjectClass<T extends IFluidLoadable> = DataObjectClass<T> &
	Record<"factory", IFluidDataStoreFactory>;

/**
 * Runtime check to determine if a class is a DataObject type.
 */
export function isDataObjectClass<T extends IFluidLoadable>(
	obj: LoadableObjectClass<T>,
): obj is InternalDataObjectClass<T>;

/**
 * Runtime check to determine if a class is a DataObject type.
 */
export function isDataObjectClass(
	obj: LoadableObjectClass,
): obj is InternalDataObjectClass<IFluidLoadable>;

/**
 * Runtime check to determine if a class is a DataObject type.
 */
export function isDataObjectClass(
	obj: LoadableObjectClass,
): obj is InternalDataObjectClass<IFluidLoadable> {
	const maybe = obj as Partial<InternalDataObjectClass<IFluidLoadable>> | undefined;
	const isDataObject =
		maybe?.factory?.IFluidDataStoreFactory !== undefined &&
		maybe.factory.IFluidDataStoreFactory === maybe.factory;

	if (
		isDataObject ===
		((obj as Partial<ISharedObjectKind<IFluidLoadable>>).getFactory !== undefined)
	) {
		// TODO: Currently nothing in the types or docs requires an actual DataObjectClass to not have a member called "getFactory" so there is a risk of this being a false positive.
		// Refactoring the use of LoadableObjectClass such that explicit down casting is not required (for example by having a single factory API shared by both cases) could avoid problems like this.
		throw new UsageError("Invalid LoadableObjectClass");
	}

	return isDataObject;
}

/**
 * Runtime check to determine if a class is a SharedObject type
 */
export function isSharedObjectKind(
	obj: LoadableObjectClass,
): obj is ISharedObjectKind<IFluidLoadable> {
	return !isDataObjectClass(obj);
}

/**
 * The ContainerSchema consists of initialObjects and dynamicObjectTypes. These types can be
 * of both SharedObject or DataObject. This function separates the two and returns a registry
 * of DataObject types and an array of SharedObjects.
 */
export const parseDataObjectsFromSharedObjects = (
	schema: ContainerSchema,
): [NamedFluidDataStoreRegistryEntry[], IChannelFactory[]] => {
	const registryEntries = new Set<NamedFluidDataStoreRegistryEntry>();
	const sharedObjects = new Set<IChannelFactory>();

	const tryAddObject = (obj: LoadableObjectClass): void => {
		if (isSharedObjectKind(obj)) {
			sharedObjects.add(obj.getFactory());
		} else if (isDataObjectClass(obj)) {
			registryEntries.add([obj.factory.type, Promise.resolve(obj.factory)]);
		} else {
			throw new Error(`Entry is neither a DataObject or a SharedObject`);
		}
	};

	// Add the object types that will be initialized
	const dedupedObjects = new Set([
		...Object.values(schema.initialObjects),
		...(schema.dynamicObjectTypes ?? []),
	]);
	for (const obj of dedupedObjects) {
		tryAddObject(obj);
	}

	if (registryEntries.size === 0 && sharedObjects.size === 0) {
		throw new Error("Container cannot be initialized without any DataTypes");
	}

	return [[...registryEntries], [...sharedObjects]];
};
