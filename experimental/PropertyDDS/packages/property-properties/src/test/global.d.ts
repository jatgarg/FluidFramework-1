/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

export { };

declare global {
    // The generated d.ts files from the .js files in src contain typescript types that are not properly imported.
    // These types are generated by the doc comments, but imports for them are not generated.
    // As these files are currently only consumed by other .js files, these d.ts stubs are enough to make it all work.
    type property = unknown;
    type BaseProperty = unknown;
    type NamedProperty = unknown;
    type NamedNodeProperty = unknown;
}
