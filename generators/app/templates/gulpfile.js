'use strict';

const build = require('@microsoft/sp-build-web');
const envInject = require('./scripts/env-inject');

// ─── Disable tslint if not present (SPFx 1.14+) ──────────────────────────────
if (build.tslintCmd && build.tslintCmd.enabled) {
  build.tslintCmd.enabled = false;
}

// ─── Register env-inject as a pre-build task ─────────────────────────────────
// This runs BEFORE any compile/bundle step, ensuring all config files
// have been patched with the correct environment values.
build.rig.addPreBuildTask(envInject);

build.initialize(require('gulp'));
