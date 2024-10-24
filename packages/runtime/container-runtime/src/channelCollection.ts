/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { AttachState } from "@fluidframework/container-definitions";
import {
	FluidObject,
	IDisposable,
	IFluidHandle,
	IRequest,
	IResponse,
	ITelemetryBaseLogger,
} from "@fluidframework/core-interfaces";
import { assert, Lazy, LazyPromise } from "@fluidframework/core-utils";
import { FluidObjectHandle } from "@fluidframework/datastore";
import { buildSnapshotTree } from "@fluidframework/driver-utils";
import { ISequencedDocumentMessage, ISnapshotTree } from "@fluidframework/protocol-definitions";
import {
	AliasResult,
	CreateSummarizerNodeSource,
	IAttachMessage,
	IEnvelope,
	IFluidDataStoreChannel,
	IFluidDataStoreContext,
	IFluidDataStoreContextDetached,
	IFluidDataStoreFactory,
	IFluidDataStoreRegistry,
	IFluidParentContext,
	IGarbageCollectionData,
	IInboundSignalMessage,
	ISummarizeResult,
	ISummaryTreeWithStats,
	ITelemetryContext,
	InboundAttachMessage,
	NamedFluidDataStoreRegistryEntries,
	channelsTreeName,
} from "@fluidframework/runtime-definitions";
import {
	GCDataBuilder,
	RequestParser,
	SummaryTreeBuilder,
	convertSnapshotTreeToSummaryTree,
	convertSummaryTreeToITree,
	create404Response,
	createResponseError,
	encodeCompactIdToString,
	isSerializedHandle,
	processAttachMessageGCData,
	responseToException,
	unpackChildNodesUsedRoutes,
} from "@fluidframework/runtime-utils";
import {
	DataCorruptionError,
	DataProcessingError,
	LoggingError,
	MonitoringContext,
	createChildLogger,
	createChildMonitoringContext,
	extractSafePropertiesFromMessage,
	tagCodeArtifacts,
} from "@fluidframework/telemetry-utils";

import { RuntimeHeaderData, defaultRuntimeHeaderData } from "./containerRuntime.js";
import {
	IDataStoreAliasMessage,
	channelToDataStore,
	isDataStoreAliasMessage,
} from "./dataStore.js";
import {
	FluidDataStoreContext,
	IFluidDataStoreContextInternal,
	ILocalDetachedFluidDataStoreContextProps,
	LocalDetachedFluidDataStoreContext,
	LocalFluidDataStoreContext,
	RemoteFluidDataStoreContext,
	createAttributesBlob,
} from "./dataStoreContext.js";
import { DataStoreContexts } from "./dataStoreContexts.js";
import { FluidDataStoreRegistry } from "./dataStoreRegistry.js";
import {
	GCNodeType,
	detectOutboundRoutesViaDDSKey,
	trimLeadingAndTrailingSlashes,
} from "./gc/index.js";
import { ContainerMessageType, LocalContainerRuntimeMessage } from "./messageTypes.js";
import { StorageServiceWithAttachBlobs } from "./storageServiceWithAttachBlobs.js";
import {
	IContainerRuntimeMetadata,
	nonDataStorePaths,
	rootHasIsolatedChannels,
} from "./summary/index.js";

/**
 * Accepted header keys for requests coming to the runtime.
 * @internal
 */
export enum RuntimeHeaders {
	/** True to wait for a data store to be created and loaded before returning it. */
	wait = "wait",
	/** True if the request is coming from an IFluidHandle. */
	viaHandle = "viaHandle",
}

/** True if a tombstoned object should be returned without erroring
 * @alpha
 */
export const AllowTombstoneRequestHeaderKey = "allowTombstone"; // Belongs in the enum above, but avoiding the breaking change
/**
 * [IRRELEVANT IF throwOnInactiveLoad OPTION NOT SET] True if an inactive object should be returned without erroring
 * @internal
 */
export const AllowInactiveRequestHeaderKey = "allowInactive"; // Belongs in the enum above, but avoiding the breaking change

type PendingAliasResolve = (success: boolean) => void;

interface FluidDataStoreMessage {
	content: any;
	type: string;
}

/**
 * Creates a shallow wrapper of {@link IFluidParentContext}. The wrapper can then have its methods overwritten as needed
 */
export function wrapContext(context: IFluidParentContext): IFluidParentContext {
	return {
		get IFluidDataStoreRegistry() {
			return context.IFluidDataStoreRegistry;
		},
		IFluidHandleContext: context.IFluidHandleContext,
		options: context.options,
		get clientId() {
			return context.clientId;
		},
		get connected() {
			return context.connected;
		},
		deltaManager: context.deltaManager,
		storage: context.storage,
		logger: context.logger,
		get clientDetails() {
			return context.clientDetails;
		},
		get idCompressor() {
			return context.idCompressor;
		},
		loadingGroupId: context.loadingGroupId,
		get attachState() {
			return context.attachState;
		},
		containerRuntime: context.containerRuntime,
		scope: context.scope,
		gcThrowOnTombstoneUsage: context.gcThrowOnTombstoneUsage,
		gcTombstoneEnforcementAllowed: context.gcTombstoneEnforcementAllowed,
		getAbsoluteUrl: async (...args) => {
			return context.getAbsoluteUrl(...args);
		},
		getQuorum: (...args) => {
			return context.getQuorum(...args);
		},
		getAudience: (...args) => {
			return context.getAudience(...args);
		},
		ensureNoDataModelChanges: (...args) => {
			return context.ensureNoDataModelChanges(...args);
		},
		submitMessage: (...args) => {
			return context.submitMessage(...args);
		},
		submitSignal: (...args) => {
			return context.submitSignal(...args);
		},
		makeLocallyVisible: (...args) => {
			return context.makeLocallyVisible(...args);
		},
		uploadBlob: async (...args) => {
			return context.uploadBlob(...args);
		},
		addedGCOutboundReference: (...args) => {
			return context.addedGCOutboundReference?.(...args);
		},
		getCreateChildSummarizerNodeFn: (...args) => {
			return context.getCreateChildSummarizerNodeFn?.(...args);
		},
		deleteChildSummarizerNode: (...args) => {
			return context.deleteChildSummarizerNode(...args);
		},
		setChannelDirty: (address: string) => {
			return context.setChannelDirty(address);
		},
	};
}

/**
 * Creates a wrapper of a {@link IFluidParentContext} to be provided to the inner datastore channels.
 * The wrapper will have the submit methods overwritten with the appropriate id as the destination address.
 *
 * @param id - the id of the channel
 * @param parentContext - the {@link IFluidParentContext} to wrap
 * @returns A wrapped {@link IFluidParentContext}
 */
export function wrapContextForInnerChannel(
	id: string,
	parentContext: IFluidParentContext,
): IFluidParentContext {
	const context = wrapContext(parentContext);

	context.submitMessage = (type: string, content: any, localOpMetadata: unknown) => {
		const fluidDataStoreContent: FluidDataStoreMessage = {
			content,
			type,
		};
		const envelope: IEnvelope = {
			address: id,
			contents: fluidDataStoreContent,
		};
		parentContext.submitMessage(
			ContainerMessageType.FluidDataStoreOp,
			envelope,
			localOpMetadata,
		);
	};

	context.submitSignal = (type: string, contents: any, targetClientId?: string) => {
		const envelope: IEnvelope = {
			address: id,
			contents,
		};
		parentContext.submitSignal(type, envelope, targetClientId);
	};

	return context;
}

/**
 * This class encapsulates data store handling. Currently it is only used by the container runtime,
 * but eventually could be hosted on any channel once we formalize the channel api boundary.
 * @internal
 */
export class ChannelCollection implements IFluidDataStoreChannel, IDisposable {
	// Stores tracked by the Domain
	private readonly pendingAttach = new Map<string, IAttachMessage>();
	// 0.24 back-compat attachingBeforeSummary
	public readonly attachOpFiredForDataStore = new Set<string>();

	protected readonly mc: MonitoringContext;

	private readonly disposeOnce = new Lazy<void>(() => this.contexts.dispose());

	public readonly entryPoint: IFluidHandle<FluidObject>;

	public readonly containerLoadStats: {
		// number of dataStores during loadContainer
		readonly containerLoadDataStoreCount: number;
		// number of unreferenced dataStores during loadContainer
		readonly referencedDataStoreCount: number;
	};

	// Stores the ids of new data stores between two GC runs. This is used to notify the garbage collector of new
	// root data stores that are added.
	private dataStoresSinceLastGC: string[] = [];
	// The handle to the container runtime. This is used mainly for GC purposes to represent outbound reference from
	// the container runtime to other nodes.
	private readonly containerRuntimeHandle: IFluidHandle;
	private readonly pendingAliasMap: Map<string, Promise<AliasResult>> = new Map<
		string,
		Promise<AliasResult>
	>();

	protected readonly contexts: DataStoreContexts;

	constructor(
		protected readonly baseSnapshot: ISnapshotTree | undefined,
		public readonly parentContext: IFluidParentContext,
		baseLogger: ITelemetryBaseLogger,
		private readonly gcNodeUpdated: (
			nodePath: string,
			reason: "Loaded" | "Changed",
			timestampMs?: number,
			packagePath?: readonly string[],
			request?: IRequest,
			headerData?: RuntimeHeaderData,
		) => void,
		private readonly isDataStoreDeleted: (nodePath: string) => boolean,
		private readonly aliasMap: Map<string, string>,
		provideEntryPoint: (runtime: ChannelCollection) => Promise<FluidObject>,
	) {
		this.mc = createChildMonitoringContext({ logger: baseLogger });
		this.contexts = new DataStoreContexts(baseLogger);
		this.containerRuntimeHandle = new FluidObjectHandle(
			this.parentContext,
			"/",
			this.parentContext.IFluidHandleContext,
		);
		this.entryPoint = new FluidObjectHandle<FluidObject>(
			new LazyPromise(async () => provideEntryPoint(this)),
			"",
			this.parentContext.IFluidHandleContext,
		);

		// Extract stores stored inside the snapshot
		const fluidDataStores = new Map<string, ISnapshotTree>();
		if (baseSnapshot) {
			for (const [key, value] of Object.entries(baseSnapshot.trees)) {
				fluidDataStores.set(key, value);
			}
		}

		let unreferencedDataStoreCount = 0;
		// Create a context for each of them
		for (const [key, value] of fluidDataStores) {
			let dataStoreContext: FluidDataStoreContext;

			// counting number of unreferenced data stores
			if (value.unreferenced) {
				unreferencedDataStoreCount++;
			}
			// If we have a detached container, then create local data store contexts.
			if (this.parentContext.attachState !== AttachState.Detached) {
				dataStoreContext = new RemoteFluidDataStoreContext({
					id: key,
					snapshotTree: value,
					parentContext: this.wrapContextForInnerChannel(key),
					storage: this.parentContext.storage,
					scope: this.parentContext.scope,
					createSummarizerNodeFn: this.parentContext.getCreateChildSummarizerNodeFn(key, {
						type: CreateSummarizerNodeSource.FromSummary,
					}),
					loadingGroupId: value.groupId,
				});
			} else {
				if (typeof value !== "object") {
					throw new LoggingError("Snapshot should be there to load from!!");
				}
				const snapshotTree = value;
				dataStoreContext = new LocalFluidDataStoreContext({
					id: key,
					pkg: undefined,
					parentContext: this.wrapContextForInnerChannel(key),
					storage: this.parentContext.storage,
					scope: this.parentContext.scope,
					createSummarizerNodeFn: this.parentContext.getCreateChildSummarizerNodeFn(key, {
						type: CreateSummarizerNodeSource.FromSummary,
					}),
					makeLocallyVisibleFn: () => this.makeDataStoreLocallyVisible(key),
					snapshotTree,
				});
			}
			this.contexts.addBoundOrRemoted(dataStoreContext);
		}
		this.containerLoadStats = {
			containerLoadDataStoreCount: fluidDataStores.size,
			referencedDataStoreCount: fluidDataStores.size - unreferencedDataStoreCount,
		};
	}

	public get aliases(): ReadonlyMap<string, string> {
		return this.aliasMap;
	}

	public get pendingAliases(): Map<string, Promise<AliasResult>> {
		return this.pendingAliasMap;
	}

	public async waitIfPendingAlias(maybeAlias: string): Promise<AliasResult> {
		const pendingAliasPromise = this.pendingAliases.get(maybeAlias);
		return pendingAliasPromise ?? "Success";
	}

	/** For sampling. Only log once per container */
	private shouldSendAttachLog = true;

	protected wrapContextForInnerChannel(id: string): IFluidParentContext {
		return wrapContextForInnerChannel(id, this.parentContext);
	}

	/**
	 * IFluidDataStoreChannel.makeVisibleAndAttachGraph implementation
	 * Not clear when it would be called and what it should do.
	 * Currently this API is called by context only for root data stores.
	 */
	public makeVisibleAndAttachGraph() {
		this.parentContext.makeLocallyVisible();
	}

	private processAttachMessage(message: ISequencedDocumentMessage, local: boolean) {
		const attachMessage = message.contents as InboundAttachMessage;

		this.dataStoresSinceLastGC.push(attachMessage.id);

		// We need to process the GC Data for both local and remote attach messages
		const foundGCData = processAttachMessageGCData(attachMessage.snapshot, (nodeId, toPath) => {
			// nodeId is the relative path under the node being attached. Always starts with "/", but no trailing "/" after an id
			const fromPath = `/${attachMessage.id}${nodeId === "/" ? "" : nodeId}`;
			this.parentContext.addedGCOutboundReference?.(
				{ absolutePath: fromPath },
				{ absolutePath: toPath },
			);
		});

		// Only log once per container to avoid noise/cost.
		// Allows longitudinal tracking of various state (e.g. foundGCData), and some sampled details
		if (this.shouldSendAttachLog) {
			this.shouldSendAttachLog = false;
			this.mc.logger.sendTelemetryEvent({
				eventName: "dataStoreAttachMessage_sampled",
				...tagCodeArtifacts({ id: attachMessage.id, pkg: attachMessage.type }),
				details: {
					local,
					snapshot: !!attachMessage.snapshot,
					foundGCData,
				},
				...extractSafePropertiesFromMessage(message),
			});
		}

		// The local object has already been attached
		if (local) {
			assert(
				this.pendingAttach.has(attachMessage.id),
				0x15e /* "Local object does not have matching attach message id" */,
			);
			this.contexts.get(attachMessage.id)?.setAttachState(AttachState.Attached);
			this.pendingAttach.delete(attachMessage.id);
			return;
		}

		// If a non-local operation then go and create the object, otherwise mark it as officially attached.
		if (this.alreadyProcessed(attachMessage.id)) {
			// TODO: dataStoreId may require a different tag from PackageData #7488
			const error = new DataCorruptionError(
				// pre-0.58 error message: duplicateDataStoreCreatedWithExistingId
				"Duplicate DataStore created with existing id",
				{
					...extractSafePropertiesFromMessage(message),
					...tagCodeArtifacts({ dataStoreId: attachMessage.id }),
				},
			);
			throw error;
		}

		const flatAttachBlobs = new Map<string, ArrayBufferLike>();
		let snapshotTree: ISnapshotTree | undefined;
		if (attachMessage.snapshot) {
			snapshotTree = buildSnapshotTree(attachMessage.snapshot.entries, flatAttachBlobs);
		}

		// Include the type of attach message which is the pkg of the store to be
		// used by RemoteFluidDataStoreContext in case it is not in the snapshot.
		const pkg = [attachMessage.type];
		const remoteFluidDataStoreContext = new RemoteFluidDataStoreContext({
			id: attachMessage.id,
			snapshotTree,
			parentContext: this.wrapContextForInnerChannel(attachMessage.id),
			storage: new StorageServiceWithAttachBlobs(this.parentContext.storage, flatAttachBlobs),
			scope: this.parentContext.scope,
			loadingGroupId: attachMessage.snapshot?.groupId,
			createSummarizerNodeFn: this.parentContext.getCreateChildSummarizerNodeFn(
				attachMessage.id,
				{
					type: CreateSummarizerNodeSource.FromAttach,
					sequenceNumber: message.sequenceNumber,
					snapshot: attachMessage.snapshot ?? {
						entries: [createAttributesBlob(pkg, true /* isRootDataStore */)],
					},
				},
			),
			pkg,
		});

		this.contexts.addBoundOrRemoted(remoteFluidDataStoreContext);
	}

	private processAliasMessage(
		message: ISequencedDocumentMessage,
		localOpMetadata: unknown,
		local: boolean,
	): void {
		const aliasMessage = message.contents as IDataStoreAliasMessage;
		if (!isDataStoreAliasMessage(aliasMessage)) {
			throw new DataCorruptionError("malformedDataStoreAliasMessage", {
				...extractSafePropertiesFromMessage(message),
			});
		}

		const resolve = localOpMetadata as PendingAliasResolve;
		const aliasResult = this.processAliasMessageCore(
			aliasMessage.internalId,
			aliasMessage.alias,
		);
		if (local) {
			resolve(aliasResult);
		}
	}

	public processAliasMessageCore(internalId: string, alias: string): boolean {
		if (this.alreadyProcessed(alias)) {
			return false;
		}

		const context = this.contexts.get(internalId);
		// If the data store has been deleted, log an error and ignore this message. This helps prevent document
		// corruption in case a deleted data store accidentally submitted a signal.
		if (this.checkAndLogIfDeleted(internalId, context, "Changed", "processAliasMessageCore")) {
			return false;
		}

		if (context === undefined) {
			this.mc.logger.sendErrorEvent({
				eventName: "AliasFluidDataStoreNotFound",
				fluidDataStoreId: internalId,
			});
			return false;
		}

		const handle = new FluidObjectHandle(
			context,
			internalId,
			this.parentContext.IFluidHandleContext,
		);
		this.parentContext.addedGCOutboundReference?.(this.containerRuntimeHandle, handle);

		this.aliasMap.set(alias, context.id);
		context.setInMemoryRoot();
		return true;
	}

	private alreadyProcessed(id: string): boolean {
		return this.aliasMap.get(id) !== undefined || this.contexts.get(id) !== undefined;
	}

	/** Package up the context's attach summary etc into an IAttachMessage */
	private generateAttachMessage(localContext: IFluidDataStoreContextInternal): IAttachMessage {
		const { attachSummary } = localContext.getAttachData(/* includeGCData: */ true);
		const type = localContext.packagePath[localContext.packagePath.length - 1];

		// Attach message needs the summary in ITree format. Convert the ISummaryTree into an ITree.
		const snapshot = convertSummaryTreeToITree(attachSummary.summary);

		return {
			id: localContext.id,
			snapshot,
			type,
		} satisfies IAttachMessage;
	}

	/**
	 * Make the data store locally visible in the container graph by moving the data store context from unbound to
	 * bound list and submitting the attach message. This data store can now be reached from the root.
	 * @param id - The id of the data store context to make visible.
	 */
	private makeDataStoreLocallyVisible(id: string): void {
		const localContext = this.contexts.getUnbound(id);
		assert(!!localContext, 0x15f /* "Could not find unbound context to bind" */);

		/**
		 * If the container is not detached, it is globally visible to all clients. This data store should also be
		 * globally visible. Move it to attaching state and send an "attach" op for it.
		 * If the container is detached, this data store will be part of the summary that makes the container attached.
		 */
		if (this.parentContext.attachState !== AttachState.Detached) {
			localContext.setAttachState(AttachState.Attaching);
			this.submitAttachChannelOp(localContext);
		}

		this.contexts.bind(id);
	}

	protected submitAttachChannelOp(localContext: LocalFluidDataStoreContext) {
		const message = this.generateAttachMessage(localContext);
		this.pendingAttach.set(localContext.id, message);
		this.parentContext.submitMessage(ContainerMessageType.Attach, message, undefined);
		this.attachOpFiredForDataStore.add(localContext.id);
	}

	/**
	 * Generate compact internal DataStore ID.
	 *
	 * A note about namespace and name collisions:
	 * This code assumes that that's the only way to generate internal IDs, and that it's Ok for this namespace to overlap with
	 * user-provided alias names namespace.
	 * There are two scenarios where it could cause trouble:
	 * 1) Old files, where (already removed) CreateRoot*DataStore*() API was used, and thus internal name of data store
	 * was provided by user. Such files may experience name collision with future data stores that receive a name generated
	 * by this function.
	 * 2) Much less likely, but if it happen that internal ID (generated by this function) is exactly the same as alias name
	 * that user might use in the future, them ContainerRuntime.getAliasedDataStoreEntryPoint() or
	 * ContainerRuntime.getDataStoreFromRequest() could return a data store with internalID matching user request, even though
	 * user expected some other data store (that would receive alias later).
	 * Please note that above mentioned functions have the implementation they have (allowing #2) due to #1.
	 */
	protected createDataStoreId(): string {
		// We use three non-overlapping namespaces:
		// - detached state: even numbers
		// - attached state: odd numbers
		// - uuids
		// In first two cases we will encode result as strings in more compact form.
		if (this.parentContext.attachState === AttachState.Detached) {
			// container is detached, only one client observes content,  no way to hit collisions with other clients.
			return encodeCompactIdToString(2 * this.contexts.size);
		}
		const id = this.parentContext.containerRuntime.generateDocumentUniqueId();
		if (typeof id === "number") {
			return encodeCompactIdToString(2 * id + 1);
		}
		return id;
	}

	public createDetachedDataStore(
		pkg: Readonly<string[]>,
		loadingGroupId?: string,
	): IFluidDataStoreContextDetached {
		return this.createContext(
			this.createDataStoreId(),
			pkg,
			LocalDetachedFluidDataStoreContext,
			undefined, // props
			loadingGroupId,
		);
	}

	public createDataStoreContext(
		pkg: Readonly<string[]>,
		props?: any,
		loadingGroupId?: string,
	): IFluidDataStoreContextInternal {
		return this.createContext(
			this.createDataStoreId(),
			pkg,
			LocalFluidDataStoreContext,
			props,
			loadingGroupId,
		);
	}

	protected createContext<T extends LocalFluidDataStoreContext>(
		id: string,
		pkg: Readonly<string[]>,
		contextCtor: new (props: ILocalDetachedFluidDataStoreContextProps) => T,
		createProps?: any,
		loadingGroupId?: string,
	) {
		const context = new contextCtor({
			id,
			pkg,
			parentContext: this.wrapContextForInnerChannel(id),
			storage: this.parentContext.storage,
			scope: this.parentContext.scope,
			createSummarizerNodeFn: this.parentContext.getCreateChildSummarizerNodeFn(id, {
				type: CreateSummarizerNodeSource.Local,
			}),
			makeLocallyVisibleFn: () => this.makeDataStoreLocallyVisible(id),
			snapshotTree: undefined,
			createProps,
			loadingGroupId,
			channelToDataStoreFn: (channel: IFluidDataStoreChannel) =>
				channelToDataStore(
					channel,
					id,
					this,
					createChildLogger({ logger: this.parentContext.logger }),
				),
		});

		this.contexts.addUnbound(context);
		return context;
	}

	public get disposed() {
		return this.disposeOnce.evaluated;
	}
	public readonly dispose = () => this.disposeOnce.value;

	public reSubmit(type: string, content: any, localOpMetadata: unknown) {
		switch (type) {
			case ContainerMessageType.Attach:
			case ContainerMessageType.Alias:
				this.parentContext.submitMessage(type, content, localOpMetadata);
				return;
			case ContainerMessageType.FluidDataStoreOp:
				return this.reSubmitChannelOp(type, content, localOpMetadata);
			default:
				assert(false, 0x907 /* unknown op type */);
		}
	}

	protected reSubmitChannelOp(type: string, content: any, localOpMetadata: unknown) {
		const envelope = content as IEnvelope;
		const context = this.contexts.get(envelope.address);
		// If the data store has been deleted, log an error and throw an error. If there are local changes for a
		// deleted data store, it can otherwise lead to inconsistent state when compared to other clients.
		if (
			this.checkAndLogIfDeleted(envelope.address, context, "Changed", "resubmitDataStoreOp")
		) {
			throw new DataCorruptionError("Context is deleted!", {
				callSite: "resubmitDataStoreOp",
				...tagCodeArtifacts({ id: envelope.address }),
			});
		}
		assert(!!context, 0x160 /* "There should be a store context for the op" */);
		const innerContents = envelope.contents as FluidDataStoreMessage;
		context.reSubmit(innerContents.type, innerContents.content, localOpMetadata);
	}

	public rollback(type: string, content: any, localOpMetadata: unknown) {
		assert(type === ContainerMessageType.FluidDataStoreOp, 0x8e8 /* type */);
		const envelope = content as IEnvelope;
		const context = this.contexts.get(envelope.address);
		// If the data store has been deleted, log an error and throw an error. If there are local changes for a
		// deleted data store, it can otherwise lead to inconsistent state when compared to other clients.
		if (
			this.checkAndLogIfDeleted(envelope.address, context, "Changed", "rollbackDataStoreOp")
		) {
			throw new DataCorruptionError("Context is deleted!", {
				callSite: "rollbackDataStoreOp",
				...tagCodeArtifacts({ id: envelope.address }),
			});
		}
		assert(!!context, 0x2e8 /* "There should be a store context for the op" */);
		const innerContents = envelope.contents as FluidDataStoreMessage;
		context.rollback(innerContents.type, innerContents.content, localOpMetadata);
	}

	public async applyStashedOp(content: unknown): Promise<unknown> {
		const opContents = content as LocalContainerRuntimeMessage;
		switch (opContents.type) {
			case ContainerMessageType.Attach:
				return this.applyStashedAttachOp(opContents.contents);
			case ContainerMessageType.Alias:
				return;
			case ContainerMessageType.FluidDataStoreOp:
				return this.applyStashedChannelChannelOp(opContents.contents);
			default:
				assert(false, 0x908 /* unknon type of op */);
		}
	}

	protected async applyStashedChannelChannelOp(envelope: IEnvelope) {
		const context = this.contexts.get(envelope.address);
		// If the data store has been deleted, log an error and ignore this message. This helps prevent document
		// corruption in case the data store that stashed the op is deleted.
		if (this.checkAndLogIfDeleted(envelope.address, context, "Changed", "applyStashedOp")) {
			return undefined;
		}
		assert(!!context, 0x161 /* "There should be a store context for the op" */);
		return context.applyStashedOp(envelope.contents);
	}

	private async applyStashedAttachOp(message: IAttachMessage) {
		const { id, snapshot } = message;

		// build the snapshot from the summary in the attach message
		const flatAttachBlobs = new Map<string, ArrayBufferLike>();
		const snapshotTree = buildSnapshotTree(snapshot.entries, flatAttachBlobs);
		const storage = new StorageServiceWithAttachBlobs(
			this.parentContext.storage,
			flatAttachBlobs,
		);

		// create a local datastore context for the data store context,
		// which this message represents. All newly created data store
		// contexts start as a local context on the client that created
		// them, and for stashed ops, the client that applies it plays
		// the role of creating client.
		const dataStoreContext = new LocalFluidDataStoreContext({
			id,
			pkg: undefined,
			parentContext: this.wrapContextForInnerChannel(id),
			storage,
			scope: this.parentContext.scope,
			createSummarizerNodeFn: this.parentContext.getCreateChildSummarizerNodeFn(id, {
				type: CreateSummarizerNodeSource.FromSummary,
			}),
			makeLocallyVisibleFn: () => this.makeDataStoreLocallyVisible(id),
			snapshotTree,
		});

		// realize the local context, as local contexts shouldn't be delay
		// loaded, as this client is playing the role of creating client,
		// and creating clients always create realized data store contexts.
		const channel = await dataStoreContext.realize();
		await channel.entryPoint.get();

		// add to the list of bound or remoted, as this context must be bound
		// to had an attach message sent, and is the non-detached case is remoted.
		this.contexts.addBoundOrRemoted(dataStoreContext);
		if (this.parentContext.attachState !== AttachState.Detached) {
			// if the client is not detached put in the pending attach list
			// so that on ack of the stashed op, the context is found.
			// detached client don't send ops, so should not expect and ack.
			this.pendingAttach.set(message.id, message);
		}
	}

	public process(
		message: ISequencedDocumentMessage,
		local: boolean,
		localMessageMetadata: unknown,
		addedOutboundReference?: (fromNodePath: string, toNodePath: string) => void,
	) {
		switch (message.type) {
			case ContainerMessageType.Attach:
				this.processAttachMessage(message, local);
				return;
			case ContainerMessageType.Alias:
				this.processAliasMessage(message, localMessageMetadata, local);
				return;
			case ContainerMessageType.FluidDataStoreOp: {
				const envelope = message.contents as IEnvelope;
				const innerContents = envelope.contents as FluidDataStoreMessage;
				const transformed = {
					...message,
					type: innerContents.type,
					contents: innerContents.content,
				};

				this.processChannelOp(envelope.address, transformed, local, localMessageMetadata);

				// By default, we use the new behavior of detecting outbound routes here.
				// If this setting is true, then DataStoreContext would be notifying GC instead.
				if (
					this.mc.config.getBoolean(detectOutboundRoutesViaDDSKey) !== true &&
					addedOutboundReference !== undefined
				) {
					// Notify GC of any outbound references that were added by this op.
					detectOutboundReferences(
						envelope.address,
						transformed.contents,
						addedOutboundReference,
					);
				}
				break;
			}
			default:
				assert(false, 0x8e9 /* unreached */);
		}
	}

	protected processChannelOp(
		address: string,
		message: ISequencedDocumentMessage,
		local: boolean,
		localMessageMetadata: unknown,
	) {
		const context = this.contexts.get(address);

		// If the data store has been deleted, log an error and ignore this message. This helps prevent document
		// corruption in case a deleted data store accidentally submitted an op.
		if (this.checkAndLogIfDeleted(address, context, "Changed", "processFluidDataStoreOp")) {
			return;
		}

		if (context === undefined) {
			// Former assert 0x162
			throw DataProcessingError.create(
				"No context for op",
				"processFluidDataStoreOp",
				message,
				{
					local,
					messageDetails: JSON.stringify({
						type: message.type,
						contentType: typeof message.contents,
					}),
					...tagCodeArtifacts({ address }),
				},
			);
		}

		context.process(message, local, localMessageMetadata);

		// Notify that a GC node for the data store changed. This is used to detect if a deleted data store is
		// being used.
		this.gcNodeUpdated(
			`/${address}`,
			"Changed",
			message.timestamp,
			context.isLoaded ? context.packagePath : undefined,
		);
	}

	public async getDataStore(
		id: string,
		requestHeaderData: RuntimeHeaderData,
	): Promise<IFluidDataStoreContextInternal> {
		const headerData = { ...defaultRuntimeHeaderData, ...requestHeaderData };
		if (
			this.checkAndLogIfDeleted(
				id,
				this.contexts.get(id),
				"Requested",
				"getDataStore",
				requestHeaderData,
			)
		) {
			// The requested data store has been deleted by gc. Create a 404 response exception.
			const request: IRequest = { url: id };
			throw responseToException(
				createResponseError(404, "DataStore was deleted", request),
				request,
			);
		}

		const context = await this.contexts.getBoundOrRemoted(id, headerData.wait);
		if (context === undefined) {
			// The requested data store does not exits. Throw a 404 response exception.
			const request: IRequest = { url: id };
			throw responseToException(create404Response(request), request);
		}
		return context;
	}

	/**
	 * Returns the data store requested with the given id if available. Otherwise, returns undefined.
	 */
	public async getDataStoreIfAvailable(
		id: string,
		requestHeaderData: RuntimeHeaderData,
	): Promise<IFluidDataStoreContextInternal | undefined> {
		// If the data store has been deleted, log an error and return undefined.
		if (
			this.checkAndLogIfDeleted(
				id,
				this.contexts.get(id),
				"Requested",
				"getDataStoreIfAvailable",
				requestHeaderData,
			)
		) {
			return undefined;
		}
		const headerData = { ...defaultRuntimeHeaderData, ...requestHeaderData };
		const context = await this.contexts.getBoundOrRemoted(id, headerData.wait);
		if (context === undefined) {
			return undefined;
		}
		return context;
	}

	/**
	 * Checks if the data store has been deleted by GC. If so, log an error.
	 * @param id - The data store's id.
	 * @param context - The data store context.
	 * @param callSite - The function name this is called from.
	 * @param requestHeaderData - The request header information to log if the data store is deleted.
	 * @returns true if the data store is deleted. Otherwise, returns false.
	 */
	private checkAndLogIfDeleted(
		id: string,
		context: IFluidDataStoreContext | undefined,
		deletedLogSuffix: string,
		callSite: string,
		requestHeaderData?: RuntimeHeaderData,
	) {
		const dataStoreNodePath = `/${id}`;
		if (!this.isDataStoreDeleted(dataStoreNodePath)) {
			return false;
		}

		this.mc.logger.sendErrorEvent({
			eventName: `GC_Deleted_DataStore_${deletedLogSuffix}`,
			...tagCodeArtifacts({ id }),
			callSite,
			headers: JSON.stringify(requestHeaderData),
			exists: context !== undefined,
		});
		return true;
	}

	public processSignal(messageArg: IInboundSignalMessage, local: boolean) {
		const envelope = messageArg.content as IEnvelope;
		const fluidDataStoreId = envelope.address;
		const message = { ...messageArg, content: envelope.contents };
		const context = this.contexts.get(fluidDataStoreId);
		// If the data store has been deleted, log an error and ignore this message. This helps prevent document
		// corruption in case a deleted data store accidentally submitted a signal.
		if (this.checkAndLogIfDeleted(fluidDataStoreId, context, "Changed", "processSignal")) {
			return;
		}

		if (!context) {
			// Attach message may not have been processed yet
			assert(!local, 0x163 /* "Missing datastore for local signal" */);
			this.mc.logger.sendTelemetryEvent({
				eventName: "SignalFluidDataStoreNotFound",
				...tagCodeArtifacts({
					fluidDataStoreId,
				}),
			});
			return;
		}

		context.processSignal(message, local);
	}

	public setConnectionState(connected: boolean, clientId?: string) {
		for (const [fluidDataStoreId, context] of this.contexts) {
			try {
				context.setConnectionState(connected, clientId);
			} catch (error) {
				this.mc.logger.sendErrorEvent(
					{
						eventName: "SetConnectionStateError",
						clientId,
						...tagCodeArtifacts({
							fluidDataStoreId,
						}),
						details: JSON.stringify({
							runtimeConnected: this.parentContext.connected,
							connected,
						}),
					},
					error,
				);
			}
		}
	}

	public setAttachState(attachState: AttachState.Attaching | AttachState.Attached): void {
		for (const [, context] of this.contexts) {
			// Fire only for bounded stores.
			if (!this.contexts.isNotBound(context.id)) {
				context.setAttachState(attachState);
			}
		}
	}

	public get size(): number {
		return this.contexts.size;
	}

	public async summarize(
		fullTree: boolean,
		trackState: boolean,
		telemetryContext?: ITelemetryContext,
	): Promise<ISummaryTreeWithStats> {
		const summaryBuilder = new SummaryTreeBuilder();

		// Iterate over each store and ask it to snapshot
		await Promise.all(
			Array.from(this.contexts)
				.filter(([_, context]) => {
					// Summarizer works only with clients with no local changes. A data store in attaching
					// state indicates an op was sent to attach a local data store, and the the attach op
					// had not yet round tripped back to the client.
					if (context.attachState === AttachState.Attaching) {
						// Formerly assert 0x588
						const error = DataProcessingError.create(
							"Local data store detected in attaching state during summarize",
							"summarize",
						);
						throw error;
					}
					return context.attachState === AttachState.Attached;
				})
				.map(async ([contextId, context]) => {
					const contextSummary = await context.summarize(
						fullTree,
						trackState,
						telemetryContext,
					);
					summaryBuilder.addWithStats(contextId, contextSummary);
				}),
		);

		return summaryBuilder.getSummaryTree();
	}

	/**
	 * Create a summary. Used when attaching or serializing a detached container.
	 */
	public getAttachSummary(telemetryContext?: ITelemetryContext): ISummaryTreeWithStats {
		const builder = new SummaryTreeBuilder();
		// Attaching graph of some stores can cause other stores to get bound too.
		// So keep taking summary until no new stores get bound.
		let notBoundContextsLength: number;
		do {
			const builderTree = builder.summary.tree;
			notBoundContextsLength = this.contexts.notBoundLength();
			// Iterate over each data store and ask it to snapshot
			Array.from(this.contexts)
				.filter(
					([key, _]) =>
						// Take summary of bounded data stores only, make sure we haven't summarized them already
						// and no attach op has been fired for that data store because for loader versions <= 0.24
						// we set attach state as "attaching" before taking createNew summary.
						!(
							this.contexts.isNotBound(key) ||
							builderTree[key] ||
							this.attachOpFiredForDataStore.has(key)
						),
				)
				.map(([key, value]) => {
					let dataStoreSummary: ISummarizeResult;
					if (value.isLoaded) {
						dataStoreSummary = value.getAttachData(
							/* includeGCCData: */ false,
							telemetryContext,
						).attachSummary;
					} else {
						// If this data store is not yet loaded, then there should be no changes in the snapshot from
						// which it was created as it is detached container. So just use the previous snapshot.
						assert(
							!!this.baseSnapshot,
							0x166 /* "BaseSnapshot should be there as detached container loaded from snapshot" */,
						);
						dataStoreSummary = convertSnapshotTreeToSummaryTree(
							this.baseSnapshot.trees[key],
						);
					}
					builder.addWithStats(key, dataStoreSummary);
				});
		} while (notBoundContextsLength !== this.contexts.notBoundLength());

		return builder.getSummaryTree();
	}

	/**
	 * Before GC runs, called by the garbage collector to update any pending GC state.
	 * The garbage collector needs to know all outbound references that are added. Since root data stores are not
	 * explicitly marked as referenced, notify GC of new root data stores that were added since the last GC run.
	 */
	public async updateStateBeforeGC(): Promise<void> {
		for (const id of this.dataStoresSinceLastGC) {
			const context = this.contexts.get(id);
			assert(context !== undefined, 0x2b6 /* Missing data store context */);
			if (await context.isRoot()) {
				// A root data store is basically a reference from the container runtime to the data store.
				const handle = new FluidObjectHandle(
					context,
					id,
					this.parentContext.IFluidHandleContext,
				);
				this.parentContext.addedGCOutboundReference?.(this.containerRuntimeHandle, handle);
			}
		}
		this.dataStoresSinceLastGC = [];
	}

	/**
	 * Generates data used for garbage collection. It does the following:
	 *
	 * 1. Calls into each child data store context to get its GC data.
	 *
	 * 2. Prefixes the child context's id to the GC nodes in the child's GC data. This makes sure that the node can be
	 * identified as belonging to the child.
	 *
	 * 3. Adds a GC node for this channel to the nodes received from the children. All these nodes together represent
	 * the GC data of this channel.
	 *
	 * @param fullGC - true to bypass optimizations and force full generation of GC data.
	 */
	public async getGCData(fullGC: boolean = false): Promise<IGarbageCollectionData> {
		const builder = new GCDataBuilder();
		// Iterate over each store and get their GC data.
		await Promise.all(
			Array.from(this.contexts)
				.filter(([_, context]) => {
					// Summarizer client and hence GC works only with clients with no local changes. A data store in
					// attaching state indicates an op was sent to attach a local data store, and the the attach op
					// had not yet round tripped back to the client.
					// Formerly assert 0x589
					if (context.attachState === AttachState.Attaching) {
						const error = DataProcessingError.create(
							"Local data store detected in attaching state while running GC",
							"getGCData",
						);
						throw error;
					}

					return context.attachState === AttachState.Attached;
				})
				.map(async ([contextId, context]) => {
					const contextGCData = await context.getGCData(fullGC);
					// Prefix the child's id to the ids of its GC nodes so they can be identified as belonging to the child.
					// This also gradually builds the id of each node to be a path from the root.
					builder.prefixAndAddNodes(contextId, contextGCData.gcNodes);
				}),
		);

		// Get the outbound routes and add a GC node for this channel.
		builder.addNode("/", await this.getOutboundRoutes());
		return builder.getGCData();
	}

	/**
	 * After GC has run, called to notify this Container's data stores of routes that are used in it.
	 * @param usedRoutes - The routes that are used in all data stores in this Container.
	 */
	public updateUsedRoutes(usedRoutes: readonly string[]) {
		// Get a map of data store ids to routes used in it.
		const usedDataStoreRoutes = unpackChildNodesUsedRoutes(usedRoutes);

		// Verify that the used routes are correct.
		for (const [id] of usedDataStoreRoutes) {
			assert(
				this.contexts.has(id),
				0x167 /* "Used route does not belong to any known data store" */,
			);
		}

		// Update the used routes in each data store. Used routes is empty for unused data stores.
		for (const [contextId, context] of this.contexts) {
			context.updateUsedRoutes(usedDataStoreRoutes.get(contextId) ?? []);
		}
	}

	public deleteChild(dataStoreId: string) {
		const dataStoreContext = this.contexts.get(dataStoreId);
		assert(dataStoreContext !== undefined, 0x2d7 /* No data store with specified id */);

		dataStoreContext.delete();
		// Delete the contexts of unused data stores.
		this.contexts.delete(dataStoreId);
		// Delete the summarizer node of the unused data stores.
		this.parentContext.deleteChildSummarizerNode(dataStoreId);
	}

	/**
	 * Delete data stores and its objects that are sweep ready.
	 * @param sweepReadyDataStoreRoutes - The routes of data stores and its objects that are sweep ready and should
	 * be deleted.
	 * @returns The routes of data stores and its objects that were deleted.
	 */
	public deleteSweepReadyNodes(sweepReadyDataStoreRoutes: readonly string[]): readonly string[] {
		for (const route of sweepReadyDataStoreRoutes) {
			const pathParts = route.split("/");
			const dataStoreId = pathParts[1];

			// Ignore sub-data store routes because a data store and its sub-routes are deleted together, so, we only
			// need to delete the data store.
			// These routes will still be returned below as among the deleted routes
			if (pathParts.length > 2) {
				continue;
			}

			const dataStoreContext = this.contexts.get(dataStoreId);
			if (dataStoreContext === undefined) {
				// If the data store hasn't already been deleted, log an error because this should never happen.
				// If the data store has already been deleted, log a telemetry event. This can happen because multiple GC
				// sweep ops can contain the same data store. It would be interesting to track how often this happens.
				const alreadyDeleted = this.isDataStoreDeleted(`/${dataStoreId}`);
				this.mc.logger.sendTelemetryEvent({
					eventName: "DeletedDataStoreNotFound",
					category: alreadyDeleted ? "generic" : "error",
					...tagCodeArtifacts({ id: dataStoreId }),
					details: { alreadyDeleted },
				});
				continue;
			}

			this.deleteChild(dataStoreId);
		}
		return Array.from(sweepReadyDataStoreRoutes);
	}

	/**
	 * This is called to update objects whose routes are tombstones.
	 *
	 * A Tombstoned object has been unreferenced long enough that GC knows it won't be referenced again.
	 * Tombstoned objects are eventually deleted by GC.
	 *
	 * @param tombstonedRoutes - The routes that are tombstones in all data stores in this Container.
	 */
	public updateTombstonedRoutes(tombstonedRoutes: readonly string[]) {
		const tombstonedDataStoresSet: Set<string> = new Set();
		for (const route of tombstonedRoutes) {
			const pathParts = route.split("/");
			// Tombstone data store only if its route (/datastoreId) is directly in tombstoneRoutes.
			if (pathParts.length > 2) {
				continue;
			}
			const dataStoreId = pathParts[1];
			assert(this.contexts.has(dataStoreId), 0x510 /* No data store with specified id */);
			tombstonedDataStoresSet.add(dataStoreId);
		}

		// Update the used routes in each data store. Used routes is empty for unused data stores.
		for (const [contextId, context] of this.contexts) {
			context.setTombstone(tombstonedDataStoresSet.has(contextId));
		}
	}

	/**
	 * Returns the outbound routes of this channel. Only root data stores are considered referenced and their paths are
	 * part of outbound routes.
	 */
	private async getOutboundRoutes(): Promise<string[]> {
		const outboundRoutes: string[] = [];
		for (const [contextId, context] of this.contexts) {
			const isRootDataStore = await context.isRoot();
			if (isRootDataStore) {
				outboundRoutes.push(`/${contextId}`);
			}
		}
		return outboundRoutes;
	}

	/**
	 * Called by GC to retrieve the package path of a data store node with the given path.
	 */
	public async getDataStorePackagePath(nodePath: string): Promise<readonly string[] | undefined> {
		// If the node belongs to a data store, return its package path. For DDSes, we return the package path of the
		// data store that contains it.
		const context = this.contexts.get(nodePath.split("/")[1]);
		return (await context?.getInitialSnapshotDetails())?.pkg;
	}

	/**
	 * Called by GC to determine if a node is for a data store or for an object within a data store (for e.g. DDS).
	 * @returns the GC node type if the node belongs to a data store or object within data store, undefined otherwise.
	 */
	public getGCNodeType(nodePath: string): GCNodeType | undefined {
		const pathParts = nodePath.split("/");
		if (!this.contexts.has(pathParts[1])) {
			return undefined;
		}

		// Data stores paths are of the format "/dataStoreId".
		// Sub data store paths are of the format "/dataStoreId/subPath/...".
		if (pathParts.length === 2) {
			return GCNodeType.DataStore;
		}
		return GCNodeType.SubDataStore;
	}

	public internalId(maybeAlias: string): string {
		return this.aliases.get(maybeAlias) ?? maybeAlias;
	}

	public async request(request: IRequest): Promise<IResponse> {
		const requestParser = RequestParser.create(request);
		const id = requestParser.pathParts[0];

		// Differentiate between requesting the dataStore directly, or one of its children
		const requestForChild = !requestParser.isLeaf(1);

		const headerData: RuntimeHeaderData = {};
		if (typeof request.headers?.[RuntimeHeaders.wait] === "boolean") {
			headerData.wait = request.headers[RuntimeHeaders.wait];
		}
		if (typeof request.headers?.[RuntimeHeaders.viaHandle] === "boolean") {
			headerData.viaHandle = request.headers[RuntimeHeaders.viaHandle];
		}
		if (typeof request.headers?.[AllowTombstoneRequestHeaderKey] === "boolean") {
			headerData.allowTombstone = request.headers[AllowTombstoneRequestHeaderKey];
		}
		if (typeof request.headers?.[AllowInactiveRequestHeaderKey] === "boolean") {
			headerData.allowInactive = request.headers[AllowInactiveRequestHeaderKey];
		}

		// We allow Tombstone requests for sub-DataStore objects
		if (requestForChild) {
			headerData.allowTombstone = true;
		}

		await this.waitIfPendingAlias(id);
		const internalId = this.internalId(id);
		const dataStoreContext = await this.getDataStore(internalId, headerData);

		// Remove query params, leading and trailing slashes from the url. This is done to make sure the format is
		// the same as GC nodes id.
		const urlWithoutQuery = trimLeadingAndTrailingSlashes(request.url.split("?")[0]);
		// Get the initial snapshot details which contain the data store package path.
		const details = await dataStoreContext.getInitialSnapshotDetails();

		// Note that this will throw if the data store is inactive or tombstoned and throwing on incorrect usage
		// is configured.
		this.gcNodeUpdated(
			`/${urlWithoutQuery}`,
			"Loaded",
			undefined /* timestampMs */,
			details.pkg,
			request,
			headerData,
		);
		const dataStore = await dataStoreContext.realize();

		const subRequest = requestParser.createSubRequest(1);
		// We always expect createSubRequest to include a leading slash, but asserting here to protect against
		// unintentionally modifying the url if that changes.
		assert(
			subRequest.url.startsWith("/"),
			0x126 /* "Expected createSubRequest url to include a leading slash" */,
		);

		return dataStore.request(subRequest);
	}
}

export function getSummaryForDatastores(
	snapshot: ISnapshotTree | undefined,
	metadata?: IContainerRuntimeMetadata,
): ISnapshotTree | undefined {
	if (!snapshot) {
		return undefined;
	}

	if (rootHasIsolatedChannels(metadata)) {
		const datastoresSnapshot = snapshot.trees[channelsTreeName];
		assert(!!datastoresSnapshot, 0x168 /* Expected tree in snapshot not found */);
		return datastoresSnapshot;
	} else {
		// back-compat: strip out all non-datastore paths before giving to DataStores object.
		const datastoresTrees: ISnapshotTree["trees"] = {};
		for (const [key, value] of Object.entries(snapshot.trees)) {
			if (!nonDataStorePaths.includes(key)) {
				datastoresTrees[key] = value;
			}
		}
		return {
			...snapshot,
			trees: datastoresTrees,
		};
	}
}

/**
 * Traverse this op's contents and detect any outbound routes that were added by this op.
 *
 * @internal
 */
export function detectOutboundReferences(
	address: string,
	contents: unknown,
	addedOutboundReference: (fromNodePath: string, toNodePath: string) => void,
): void {
	// These will be built up as we traverse the envelope contents
	const outboundPaths: string[] = [];
	let ddsAddress: string | undefined;

	function recursivelyFindHandles(obj: unknown) {
		if (typeof obj === "object" && obj !== null) {
			for (const [key, value] of Object.entries(obj)) {
				// If 'value' is a serialized IFluidHandle, it represents a new outbound route.
				if (isSerializedHandle(value)) {
					outboundPaths.push(value.url);
				}

				// NOTE: This is taking a hard dependency on the fact that in our DataStore implementation,
				// the address of the DDS is stored in a property called "address".  This is not ideal.
				// An alternative would be for the op envelope to include the absolute path (built up as it is submitted)
				if (key === "address" && ddsAddress === undefined) {
					ddsAddress = value;
				}

				recursivelyFindHandles(value);
			}
		}
	}

	recursivelyFindHandles(contents);

	// GC node paths are all absolute paths, hence the "" prefix.
	// e.g. this will yield "/dataStoreId/ddsId"
	const fromPath = ["", address, ddsAddress].join("/");
	outboundPaths.forEach((toPath) => addedOutboundReference(fromPath, toPath));
}

/** @internal */
export class ChannelCollectionFactory<T extends ChannelCollection = ChannelCollection>
	implements IFluidDataStoreFactory
{
	public readonly type = "ChannelCollectionChannel";

	public IFluidDataStoreRegistry: IFluidDataStoreRegistry;

	constructor(
		registryEntries: NamedFluidDataStoreRegistryEntries,
		// ADO:7302 We need a better type here
		private readonly provideEntryPoint: (
			runtime: IFluidDataStoreChannel,
		) => Promise<FluidObject>,
		private readonly ctor: (...args: ConstructorParameters<typeof ChannelCollection>) => T,
	) {
		this.IFluidDataStoreRegistry = new FluidDataStoreRegistry(registryEntries);
	}

	public get IFluidDataStoreFactory() {
		return this;
	}

	public async instantiateDataStore(
		context: IFluidDataStoreContext,
		_existing: boolean,
	): Promise<IFluidDataStoreChannel> {
		const runtime = this.ctor(
			context.baseSnapshot,
			context, // parentContext
			context.logger,
			() => {}, // gcNodeUpdated
			(_nodePath: string) => false, // isDataStoreDeleted
			new Map(), // aliasMap
			this.provideEntryPoint,
		);

		return runtime;
	}
}
