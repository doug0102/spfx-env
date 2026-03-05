'use strict';

/**
 * env-inject.js
 * ─────────────────────────────────────────────────────────────────────────────
 * SPFx Gulp pre-build task that reads an environment variable file and injects
 * values into:
 *
 *   config/package-solution.json
 *     solution.name                              ← solutionName
 *     solution.id                                ← solutionId
 *     solution.webApiPermissionRequests[*].resource ← webApiPermissions[*].webApiPermName
 *     paths.zippedPackage                        ← solutionPackageZipPath
 *
 *   .yo-rc.json
 *     @microsoft/generator-sharepoint.libraryId  ← libraryId
 *
 *   src/**‌/*-manifest.schema.json  (one per manifest, matched by array index)
 *     id                                         ← manifests[n].manifestId
 *     preconfiguredEntries[*].groupId            ← manifests[n].webGroupId
 *     preconfiguredEntries[*].title              ← manifests[n].webTitle
 *
 * Usage (via gulp):
 *   gulp serve              --env dev
 *   gulp bundle --ship      --env uat
 *   gulp package-solution   --env prod
 *
 * The --env flag is read from process.argv. Falls back to "dev".
 * ─────────────────────────────────────────────────────────────────────────────
 */

const fs    = require('fs');
const path  = require('path');
const build = require('@microsoft/sp-build-web');

// ─── Resolve --env argument ───────────────────────────────────────────────────

function resolveEnvironment() {
  const args   = process.argv;
  const envIdx = args.indexOf('--env');
  const env    = envIdx !== -1 ? args[envIdx + 1] : null;

  const valid = ['dev', 'uat', 'prod'];

  if (!env) {
    return 'dev';
  }

  if (!valid.includes(env.toLowerCase())) {
    throw new Error(
      `[env-inject] Invalid --env value: "${env}". Must be one of: ${valid.join(', ')}`
    );
  }

  return env.toLowerCase();
}

// ─── Logging helpers ──────────────────────────────────────────────────────────

function logInfo(msg)  { console.log(`[env-inject] ℹ️  ${msg}`); }
function logOk(msg)    { console.log(`[env-inject] ✔  ${msg}`); }
function logWarn(msg)  { console.warn(`[env-inject] ⚠️  ${msg}`); }
function logError(msg) { console.error(`[env-inject] ✖  ${msg}`); }

// ─── Safe JSON helpers ────────────────────────────────────────────────────────

function stripJsonComments(str) {
  return str
    .replace(/^\uFEFF/, '')                // strip BOM
    .replace(/\/\*[\s\S]*?\*\//g, '')      // strip /* */ block comments
    .replace(/\/\/[^\n]*/g, '')            // strip // line comments
    .replace(/,(\s*[}\]])/g, '$1');        // strip trailing commas
}

function readJson(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  try {
    return JSON.parse(stripJsonComments(fs.readFileSync(filePath, 'utf8')));
  } catch (e) {
    throw new Error(`Failed to parse JSON at ${filePath}: ${e.message}`);
  }
}

function writeJson(filePath, data) {
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2) + '\n', 'utf8');
}

function backupJson(filePath) {
  const bakPath = filePath.replace(/\.json$/, '.bak.json');
  if (!fs.existsSync(bakPath)) {
    fs.copyFileSync(filePath, bakPath);
    logInfo(`Backup created: ${path.basename(bakPath)}`);
  }
}

// ─── Deep set by dot-path ─────────────────────────────────────────────────────
// e.g. setDeep(obj, 'solution.name', 'MyApp')

function setDeep(obj, dotPath, value) {
  const parts = dotPath.split('.');
  let cursor = obj;
  for (let i = 0; i < parts.length - 1; i++) {
    const key = parts[i];
    if (cursor[key] === undefined || cursor[key] === null) {
      throw new Error(`Path segment "${key}" not found while traversing "${dotPath}"`);
    }
    cursor = cursor[key];
  }
  cursor[parts[parts.length - 1]] = value;
}

// ─── Validate env file ────────────────────────────────────────────────────────

function validateEnvFile(env) {
  const required = [
    'solutionName',
    'solutionId',
    'solutionPackageZipPath',
    'webApiPermissions',
    'libraryId',
    'manifests'
  ];

  // Keys must exist, but null values are allowed — nulls are skipped at inject time
  const missing = required.filter(k => !(k in env));
  if (missing.length > 0) {
    throw new Error(
      `env file is missing required keys: ${missing.join(', ')}\n` +
      `  Check your env/*.env.json file.`
    );
  }

  if (env.webApiPermissions !== null && !Array.isArray(env.webApiPermissions)) {
    throw new Error('"webApiPermissions" must be an array or null.');
  }

  if (env.manifests !== null && !Array.isArray(env.manifests)) {
    throw new Error('"manifests" must be an array or null.');
  }
}

// ─── Null guard helper ────────────────────────────────────────────────────────

function hasValue(val) {
  return val !== null && val !== undefined && val !== '';
}

// ─── Injectors ───────────────────────────────────────────────────────────────

function injectPackageSolution(envVars, projectRoot) {
  const filePath = path.join(projectRoot, 'config', 'package-solution.json');
  logInfo(`Injecting → config/package-solution.json`);

  const json = readJson(filePath);
  backupJson(filePath);

  // solution.name
  if (hasValue(envVars.solutionName)) {
    setDeep(json, 'solution.name', envVars.solutionName);
    logOk(`  solution.name = "${envVars.solutionName}"`);
  } else {
    logWarn(`  solution.name is null — skipping, existing value preserved`);
  }

  // solution.id
  if (hasValue(envVars.solutionId)) {
    setDeep(json, 'solution.id', envVars.solutionId);
    logOk(`  solution.id = "${envVars.solutionId}"`);
  } else {
    logWarn(`  solution.id is null — skipping, existing value preserved`);
  }

  // paths.zippedPackage
  if (hasValue(envVars.solutionPackageZipPath)) {
    setDeep(json, 'paths.zippedPackage', envVars.solutionPackageZipPath);
    logOk(`  paths.zippedPackage = "${envVars.solutionPackageZipPath}"`);
  } else {
    logWarn(`  solutionPackageZipPath is null — skipping, existing value preserved`);
  }

  // solution.webApiPermissionRequests[*].resource
  const perms = json.solution.webApiPermissionRequests;
  if (!Array.isArray(perms)) {
    logWarn('  solution.webApiPermissionRequests is not an array — skipping resource injection');
  } else {
    const envPerms = envVars.webApiPermissions;

    if (perms.length !== envPerms.length) {
      logWarn(
        `  webApiPermissionRequests length mismatch: ` +
        `file has ${perms.length}, env has ${envPerms.length}. ` +
        `Injecting up to min(${Math.min(perms.length, envPerms.length)}).`
      );
    }

    const limit = Math.min(perms.length, envPerms.length);
    for (let i = 0; i < limit; i++) {
      if (hasValue(envPerms[i].webApiPermName)) {
        perms[i].resource = envPerms[i].webApiPermName;
        logOk(`  webApiPermissionRequests[${i}].resource = "${envPerms[i].webApiPermName}"`);
      } else {
        logWarn(`  webApiPermissions[${i}].webApiPermName is null — skipping, existing value preserved`);
      }
    }
  }

  writeJson(filePath, json);
}

function injectYoRc(envVars, projectRoot) {
  const filePath = path.join(projectRoot, '.yo-rc.json');
  logInfo(`Injecting → .yo-rc.json`);

  const json = readJson(filePath);
  backupJson(filePath);

  const ns = '@microsoft/generator-sharepoint';
  if (!json[ns]) {
    throw new Error(`.yo-rc.json does not contain "${ns}" namespace.`);
  }

  if (hasValue(envVars.libraryId)) {
    json[ns].libraryId = envVars.libraryId;
    logOk(`  ${ns}.libraryId = "${envVars.libraryId}"`);
  } else {
    logWarn(`  libraryId is null — skipping, existing value preserved`);
  }

  writeJson(filePath, json);
}

function injectManifests(envVars, projectRoot) {
  // Find all *.manifest.json files under src/
  let globSync;
  try {
    globSync = require('glob').sync;
  } catch(e) {
    throw new Error(
      'The "glob" package is required. Run: npm install --save-dev glob@8'
    );
  }

  const pattern  = path.join(projectRoot, 'src', '**', '*.manifest.json');
  const files    = globSync(pattern.replace(/\\/g, '/'));
  const envMfsts = envVars.manifests;

  if (files.length === 0) {
    logWarn('No *.manifest.json files found under src/');
    return;
  }

  if (files.length !== envMfsts.length) {
    logWarn(
      `Manifest count mismatch: found ${files.length} manifest file(s), ` +
      `env has ${envMfsts.length} entry/entries. ` +
      `Injecting up to min(${Math.min(files.length, envMfsts.length)}).`
    );
  }

  const limit = Math.min(files.length, envMfsts.length);

  for (let i = 0; i < limit; i++) {
    const filePath = files[i];
    const envM     = envMfsts[i];
    const relPath  = path.relative(projectRoot, filePath);
    logInfo(`Injecting → ${relPath}`);

    const json = readJson(filePath);
    backupJson(filePath);

    // id
    if (hasValue(envM.manifestId)) {
      json.id = envM.manifestId;
      logOk(`  id = "${envM.manifestId}"`);
    } else {
      logWarn(`  manifests[${i}].manifestId is null — skipping, existing value preserved`);
    }

    // preconfiguredEntries[*].groupId and .title
    if (!Array.isArray(json.preconfiguredEntries) || json.preconfiguredEntries.length === 0) {
      logWarn(`  preconfiguredEntries not found or empty in ${relPath}`);
    } else {
      json.preconfiguredEntries.forEach((entry, idx) => {
        if (hasValue(envM.webGroupId)) {
          entry.groupId = envM.webGroupId;
          logOk(`  preconfiguredEntries[${idx}].groupId = "${envM.webGroupId}"`);
        } else {
          logWarn(`  manifests[${i}].webGroupId is null — skipping, existing value preserved`);
        }

        if (hasValue(envM.webTitle)) {
          entry.title = typeof entry.title === 'object'
            ? { ...entry.title, default: envM.webTitle }
            : envM.webTitle;
          logOk(`  preconfiguredEntries[${idx}].title = "${envM.webTitle}"`);
        } else {
          logWarn(`  manifests[${i}].webTitle is null — skipping, existing value preserved`);
        }
      });
    }

    writeJson(filePath, json);
  }
}

// ─── TypeScript scaffolding ───────────────────────────────────────────────────

function scaffoldTypeDeclaration(envVars, projectRoot) {
  const typesDir  = path.join(projectRoot, 'src', 'types');
  const filePath  = path.join(typesDir, 'env.d.ts');

  // Build property lines from all keys in the env file, skipping private _ keys
  const knownKeys = new Set([
    '_environment', '_description', '$schema',
    'solutionName', 'solutionId', 'solutionPackageZipPath',
    'webApiPermissions', 'libraryId', 'manifests'
  ]);

  const customProps = Object.keys(envVars)
    .filter(k => !knownKeys.has(k))
    .map(k => {
      const val  = envVars[k];
      const type = Array.isArray(val) ? 'unknown[]'
                 : val === null       ? 'string | null'
                 : typeof val === 'number'  ? 'number'
                 : typeof val === 'boolean' ? 'boolean'
                 : 'string';
      return `    ${k}: ${type};`;
    })
    .join('\n');

  const content =
`// Auto-generated by env-inject.js — do not edit manually.
// Re-generated on every gulp serve / bundle run.
declare module '*/env/*.env.json' {
  const value: {
    _environment: string;
    solutionName: string | null;
    solutionId: string | null;
    solutionPackageZipPath: string | null;
    webApiPermissions: { webApiPermName: string }[] | null;
    libraryId: string | null;
    manifests: {
      manifestId: string | null;
      webGroupId: string | null;
      webTitle: string | null;
    }[] | null;
${customProps ? customProps + '\n' : ''}  };
  export default value;
}
`;

  if (!fs.existsSync(typesDir)) {
    fs.mkdirSync(typesDir, { recursive: true });
    logInfo('Created src/types/');
  }

  fs.writeFileSync(filePath, content, 'utf8');
  logOk('Scaffolded src/types/env.d.ts');
}

function scaffoldTsConfig(projectRoot) {
  const filePath = path.join(projectRoot, 'tsconfig.json');

  if (!fs.existsSync(filePath)) {
    logWarn('tsconfig.json not found — skipping resolveJsonModule patch');
    return;
  }

  const json = readJson(filePath);

  if (json.compilerOptions?.resolveJsonModule === true) {
    logInfo('tsconfig.json already has resolveJsonModule: true — skipping');
    return;
  }

  if (!json.compilerOptions) json.compilerOptions = {};
  json.compilerOptions.resolveJsonModule = true;

  writeJson(filePath, json);
  logOk('Patched tsconfig.json → compilerOptions.resolveJsonModule = true');
}

function scaffoldEnvConfig(envName, projectRoot) {
  const configDir = path.join(projectRoot, 'src', 'config');
  const filePath  = path.join(configDir, 'envConfig.ts');

  const content =
`// Auto-generated by env-inject.js — do not edit manually.
// Re-generated on every gulp serve / bundle run to point at the active environment.
import env from '../../env/${envName}.env.json';
export default env;
`;

  if (!fs.existsSync(configDir)) {
    fs.mkdirSync(configDir, { recursive: true });
    logInfo('Created src/config/');
  }

  fs.writeFileSync(filePath, content, 'utf8');
  logOk(`Scaffolded src/config/envConfig.ts → imports ${envName}.env.json`);
}

// ─── Build task definition ────────────────────────────────────────────────────

const envInjectTask = build.subTask('env-inject', function(gulp, buildOptions, done) {
  try {
    const projectRoot = buildOptions.rootPath || process.cwd();
    const envName     = resolveEnvironment();
    const envFilePath = path.join(projectRoot, 'env', `${envName}.env.json`);

    logInfo(`Environment  : ${envName.toUpperCase()}`);
    logInfo(`Env file     : ${envFilePath}`);
    logInfo(`Project root : ${projectRoot}`);
    console.log('');

    if (!fs.existsSync(envFilePath)) {
      throw new Error(
        `Env file not found: ${envFilePath}\n` +
        `  Create it by copying env/dev.env.json and filling in your values.`
      );
    }

    const envVars = readJson(envFilePath);
    validateEnvFile(envVars);

    injectPackageSolution(envVars, projectRoot);
    console.log('');
    injectYoRc(envVars, projectRoot);
    console.log('');
    injectManifests(envVars, projectRoot);
    console.log('');

    logInfo('Scaffolding TypeScript config references...');
    scaffoldTypeDeclaration(envVars, projectRoot);
    scaffoldTsConfig(projectRoot);
    scaffoldEnvConfig(envName, projectRoot);
    console.log('');

    logInfo('─────────────────────────────────────────────');
    logOk(`Injection complete for environment: ${envName.toUpperCase()}`);
    logInfo('─────────────────────────────────────────────\n');

    done();
  } catch (err) {
    logError(err.message);
    done(new Error(err.message));
  }
});

module.exports = envInjectTask;
