#!/usr/bin/env node
/*
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { Buffer } from "node:buffer";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const fixturesDir = dirname(fileURLToPath(import.meta.url));

const fixtures = [
	{
		name: "simple-hello",
		text: "Hello, world!",
		ops: [{ kind: "insert", position: 0, text: "Hello, world!" }],
	},
	{
		name: "two-insertions",
		text: "Hello beautiful world!",
		ops: [
			{ kind: "insert", position: 0, text: "Hello world!" },
			{ kind: "insert", position: 6, text: "beautiful " },
		],
	},
	{
		name: "insert-then-delete",
		text: "Hell world!",
		ops: [
			{ kind: "insert", position: 0, text: "Hello world!" },
			{ kind: "delete", start: 4, end: 5, text: "o" },
		],
	},
];

function applyOps(sharedString, ops) {
	for (const op of ops) {
		switch (op.kind) {
			case "insert": {
				sharedString.insertText(op.position, op.text);
				break;
			}
			case "delete": {
				sharedString.removeText(op.start, op.end);
				break;
			}
			default: {
				throw new Error(`Unsupported op: ${JSON.stringify(op)}`);
			}
		}
	}
}

function asBuffer(content) {
	if (typeof content === "string") {
		return Buffer.from(content, "utf8");
	}
	if (content instanceof Uint8Array) {
		return Buffer.from(content);
	}
	throw new Error(`Unsupported summary blob content type: ${typeof content}`);
}

function extractHeaderBlob(summary) {
	const content = summary?.tree?.content;
	const header = content?.tree?.header;
	if (header?.content === undefined) {
		throw new Error(
			`Unable to find SharedString content/header blob in summary: ${JSON.stringify(summary)}`,
		);
	}
	return asBuffer(header.content);
}

function prettyPrintBlob(blob) {
	const text = blob.toString("utf8");
	try {
		return `${JSON.stringify(JSON.parse(text), undefined, 2)}\n`;
	} catch {
		return `${text}\n`;
	}
}

function operationsMarkdown(ops) {
	return ops
		.map((op, index) => {
			if (op.kind === "insert") {
				return `${index + 1}. Insert \`${JSON.stringify(op.text).slice(1, -1)}\` at position ${op.position}.`;
			}
			return `${index + 1}. Delete positions ${op.start}-${op.end} (removes \`${op.text}\`).`;
		})
		.join("\n");
}

function descriptionMarkdown(fixture) {
	return `# ${fixture.name}\n\n` +
		`Golden SharedString snapshot fixture for the C# snapshot loader POC.\n\n` +
		`## Text content\n\n` +
		`\`${JSON.stringify(fixture.text).slice(1, -1)}\`\n\n` +
		`## Operations performed\n\n` +
		`${operationsMarkdown(fixture.ops)}\n\n` +
		`## Expected length\n\n` +
		`${fixture.text.length} UTF-16 code units.\n`;
}

async function loadFluidPackages() {
	const [{ SharedString, SharedStringFactory }, runtimeUtils] = await Promise.all([
		import("@fluidframework/sequence/internal"),
		import("@fluidframework/test-runtime-utils/internal"),
	]);

	const factory =
		SharedStringFactory === undefined ? SharedString.getFactory() : new SharedStringFactory();

	return {
		factory,
		MockContainerRuntimeFactory: runtimeUtils.MockContainerRuntimeFactory,
		MockFluidDataStoreRuntime: runtimeUtils.MockFluidDataStoreRuntime,
		MockStorage: runtimeUtils.MockStorage,
	};
}

async function summarizeWithSharedString(fixture, fluid) {
	const dataStoreRuntime = new fluid.MockFluidDataStoreRuntime({
		clientId: `${fixture.name}-client`,
		id: `${fixture.name}-datastore`,
	});
	dataStoreRuntime.options.newMergeTreeSnapshotFormat = true;

	const containerRuntimeFactory = new fluid.MockContainerRuntimeFactory();
	containerRuntimeFactory.createContainerRuntime(dataStoreRuntime);

	const sharedString = fluid.factory.create(dataStoreRuntime, `${fixture.name}-shared-string`);
	sharedString.connect({
		deltaConnection: dataStoreRuntime.createDeltaConnection(),
		objectStorage: new fluid.MockStorage(),
	});

	applyOps(sharedString, fixture.ops);
	containerRuntimeFactory.processAllMessages();

	// Advance MSN so these stable golden snapshots contain only the final visible segments.
	dataStoreRuntime.deltaManagerInternal.minimumSequenceNumber =
		dataStoreRuntime.deltaManagerInternal.lastSequenceNumber;

	if (sharedString.getText() !== fixture.text) {
		throw new Error(
			`${fixture.name} text mismatch: expected ${JSON.stringify(fixture.text)}, got ${JSON.stringify(
				sharedString.getText(),
			)}`,
		);
	}

	const { summary } = await sharedString.summarize();
	return extractHeaderBlob(summary);
}

function fallbackSnapshotBlob(fixture) {
	const sequenceNumber = fixture.ops.length;
	return Buffer.from(
		JSON.stringify({
			version: "1",
			segmentCount: 1,
			length: fixture.text.length,
			segments: [fixture.text],
			startIndex: 0,
			headerMetadata: {
				minSequenceNumber: sequenceNumber,
				sequenceNumber,
				orderedChunkMetadata: [{ id: "header" }],
				totalLength: fixture.text.length,
				totalSegmentCount: 1,
			},
		}),
		"utf8",
	);
}

async function getSnapshotBlob(fixture) {
	try {
		const result = await summarizeWithSharedString(fixture, await loadFluidPackages());
		console.log(`[real Fluid] ${fixture.name}: ${result.length} bytes`);
		return result;
	} catch (error) {
		console.log(`[fallback] ${fixture.name}: ${error?.code ?? "unknown"} — ${error?.message?.slice(0, 100) ?? ""}`);
		if (error?.code !== "ERR_MODULE_NOT_FOUND") {
			throw error;
		}
		return fallbackSnapshotBlob(fixture);
	}
}

await mkdir(fixturesDir, { recursive: true });

let generated = 0;
for (const fixture of fixtures) {
	const snapshotBlob = await getSnapshotBlob(fixture);
	await writeFile(join(fixturesDir, `${fixture.name}.snapshot.bin`), snapshotBlob);
	await writeFile(join(fixturesDir, `${fixture.name}.snapshot.json.txt`), prettyPrintBlob(snapshotBlob));
	await writeFile(join(fixturesDir, `${fixture.name}.description.md`), descriptionMarkdown(fixture));
	generated++;
}

console.log(`Generated ${generated} fixtures in ${fixturesDir}`);
