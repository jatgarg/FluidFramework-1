/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import type { Linter } from "eslint";
import { recommended } from "@fluidframework/eslint-config-fluid/flat.mts";

const config: Linter.Config[] = [
	...recommended,
	{
		// C# port work lives under src/csharp-port/ and is excluded from TS build
		// (see tsconfig.json). Also exclude it from ESLint so shim / fixture-generator
		// files (which use Node builtins, non-TS syntax, or bin/obj build artifacts)
		// don't fail lint.
		ignores: ["src/csharp-port/**"],
	},
	{
		rules: {
			"@fluid-internal/fluid/no-unchecked-record-access": "warn",
			"@typescript-eslint/explicit-function-return-type": "warn",
			"@typescript-eslint/explicit-module-boundary-types": "off",
			"@typescript-eslint/no-explicit-any": "off",
			"@typescript-eslint/no-unsafe-argument": "off",
			"@typescript-eslint/no-unsafe-assignment": "off",
			"@typescript-eslint/no-unsafe-call": "off",
			"@typescript-eslint/no-unsafe-member-access": "off",
			"@typescript-eslint/strict-boolean-expressions": "off",
			"no-void": "off",
			"unicorn/no-array-for-each": "off",
			"unicorn/no-array-reduce": "off",
			"unicorn/no-null": "off",
			"unicorn/no-object-as-default-parameter": "off",
			"unicorn/prefer-code-point": "off",
			"unicorn/prefer-spread": "off",
			"unicorn/prefer-string-slice": "off",
		},
	},
];

export default config;
