/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

const fluidRoute = require("@fluid-example/webpack-fluid-loader");
const path = require("path");
const webpack = require("webpack");
const { merge } = require("webpack-merge");

module.exports = (env) =>
	merge(
		{
			entry: {
				main: "./src/index.ts",
			},
			resolve: {
				extensionAlias: {
					".js": [".ts", ".tsx", ".js"],
					".mjs": [".mts", ".mjs"],
				},
			},
			module: {
				rules: [
					{
						test: /\.tsx?$/,
						loader: "ts-loader",
					},
					{
						test: /\.m?js$/,
						use: [require.resolve("source-map-loader")],
						enforce: "pre",
					},
				],
			},
			output: {
				filename: "[name].bundle.js",
				path: path.resolve(__dirname, "dist"),
				library: { name: "[name]", type: "umd" },
			},
			watchOptions: {
				ignored: "**/node_modules/**",
			},
		},
		env?.production
			? {
					mode: "production",
					devtool: "source-map",
			  }
			: {
					mode: "development",
					devtool: "inline-source-map",
			  },
		fluidRoute.devServerConfig(__dirname, env),
	);
