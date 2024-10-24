/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import path from "path";

import { IIdCompressor } from "@fluidframework/id-compressor";

import { ChangesetLocalId, RevisionTagCodec } from "../../../core/index.js";
import {
	OptionalChangeset,
	makeOptionalFieldCodecFamily,
	// eslint-disable-next-line import/no-internal-modules
} from "../../../feature-libraries/optional-field/index.js";
import { brand } from "../../../util/index.js";
import { takeJsonSnapshot, useSnapshotDirectory } from "../../snapshots/index.js";
// eslint-disable-next-line import/no-internal-modules
import { createSnapshotCompressor } from "../../snapshots/testTrees.js";
import { TestChange } from "../../testChange.js";

import { Change } from "./optionalFieldUtils.js";

function generateTestChangesets(
	idCompressor: IIdCompressor,
): { name: string; change: OptionalChangeset<TestChange> }[] {
	const revision = idCompressor.generateCompressedId();
	const localId: ChangesetLocalId = brand(42);
	const childChange = TestChange.mint([], 1);
	return [
		{
			name: "empty",
			change: Change.empty(),
		},
		{
			name: "change with moves",
			change: Change.atOnce(
				Change.move({ revision, localId }, "self"),
				Change.clear("self", { revision, localId }),
				Change.move(localId, localId),
			),
		},
		{
			name: "with child change",
			change: Change.atOnce(
				Change.childAt({ revision, localId }, childChange),
				Change.childAt(localId, childChange),
				Change.child(childChange),
			),
		},
		{
			name: "with reserved detach",
			change: Change.reserve("self", { revision, localId }),
		},
		{
			name: "pin",
			change: Change.pin({ revision, localId }),
		},
	];
}

export function testSnapshots() {
	describe("Snapshots", () => {
		const snapshotCompressor = createSnapshotCompressor();
		const changesets = generateTestChangesets(snapshotCompressor);
		const family = makeOptionalFieldCodecFamily(
			TestChange.codec,
			new RevisionTagCodec(snapshotCompressor),
		);

		for (const version of family.getSupportedFormats()) {
			describe(`version ${version}`, () => {
				const dir = path.join("optional-field", `V${version}`);
				useSnapshotDirectory(dir);
				const codec = family.resolve(version);
				for (const { name, change } of changesets) {
					it(name, () => {
						const encoded = codec.json.encode(change, {
							originatorId: snapshotCompressor.localSessionId,
						});
						takeJsonSnapshot(encoded);
					});
				}
			});
		}
	});
}
