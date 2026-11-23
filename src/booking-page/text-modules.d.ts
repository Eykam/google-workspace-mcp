/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

// Bundled as strings by esbuild (see the loader config in esbuild.config.js).
declare module '*.gs' {
  const source: string;
  export default source;
}

declare module '*.html' {
  const source: string;
  export default source;
}
