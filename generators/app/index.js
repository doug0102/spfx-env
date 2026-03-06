'use strict';

const Generator = require('yeoman-generator');
const path = require('path');
const fs = require('fs');
const chalk = require('chalk');

module.exports = class extends Generator {
  constructor(args, opts) {
    super(args, opts);
    this.log(chalk.cyan('\n🔧 SPFx Environment Config Injector - Yeoman Generator\n'));
  }

  // ─── STEP 1: Prompts ────────────────────────────────────────────────────────

  async prompting() {
    const cwd = this.destinationRoot();

    // Detect if this looks like an SPFx project
    const hasSolution = fs.existsSync(path.join(cwd, 'config', 'package-solution.json'));
    const hasYoRc     = fs.existsSync(path.join(cwd, '.yo-rc.json'));

    if (!hasSolution || !hasYoRc) {
      this.log(chalk.yellow(
        '⚠️  Warning: This does not appear to be an SPFx project root.\n' +
        '   Expected: config/package-solution.json and .yo-rc.json\n' +
        '   Continuing anyway — you may need to adjust paths manually.\n'
      ));
    }

    this.answers = await this.prompt([
      {
        type: 'confirm',
        name: 'overwriteEnvFiles',
        message: 'env/ files already exist (if any). Overwrite environment files?',
        default: false
      }
    ]);
  }

  // ─── STEP 2: Write files ─────────────────────────────────────────────────────

  writing() {
    this.projectValues = this._readProjectValues();
    this._writeEnvFiles();
    this._writeInjectorScript();
    this._writeGulpfile();
    this._scaffoldTypeDeclaration();
    this._scaffoldTsConfig();
    this._scaffoldEnvConfig();
    this._patchGitignore();
  }

  // ─── READ EXISTING PROJECT VALUES ────────────────────────────────────────────

  _readProjectValues() {
    const cwd     = this.destinationRoot();
    const values  = {};
    const missing = [];

    // ── package-solution.json ──────────────────────────────────────────────────
    const solutionPath = path.join(cwd, 'config', 'package-solution.json');
    if (fs.existsSync(solutionPath)) {
      try {
        const sol = JSON.parse(fs.readFileSync(solutionPath, 'utf8'));

        values.solutionName           = sol.solution?.name          || null;
        values.solutionId             = sol.solution?.id            || null;
        values.solutionPackageZipPath = sol.paths?.zippedPackage    || null;

        const perms = sol.solution?.webApiPermissionRequests;
        values.webApiPermissions = Array.isArray(perms) && perms.length > 0
          ? perms.map(p => ({ webApiPermName: p.resource || '' }))
          : null;

      } catch (e) {
        this.log(chalk.yellow(`  ⚠️  Could not parse config/package-solution.json: ${e.message}`));
      }
    } else {
      missing.push('config/package-solution.json');
    }

    // ── .yo-rc.json ────────────────────────────────────────────────────────────
    const yorcPath = path.join(cwd, '.yo-rc.json');
    if (fs.existsSync(yorcPath)) {
      try {
        const yorc = JSON.parse(fs.readFileSync(yorcPath, 'utf8'));
        values.libraryId = yorc['@microsoft/generator-sharepoint']?.libraryId || null;
      } catch (e) {
        this.log(chalk.yellow(`  ⚠️  Could not parse .yo-rc.json: ${e.message}`));
      }
    } else {
      missing.push('.yo-rc.json');
    }

    // ── *-manifest.schema.json files ──────────────────────────────────────────
    const srcDir = path.join(cwd, 'src');
    if (fs.existsSync(srcDir)) {
      try {
        const manifestFiles = this._findManifestFiles(srcDir);

        values.manifests = manifestFiles.length > 0
          ? manifestFiles.map(filePath => {
              try {
                const raw     = fs.readFileSync(filePath, 'utf8');
                const cleaned = this._stripJsonComments(raw);
                const m       = JSON.parse(cleaned);

                this.log(chalk.gray(`     manifest file     : ${path.relative(cwd, filePath)}`));
                this.log(chalk.gray(`     m.id              : ${m.id}`));
                this.log(chalk.gray(`     preconfiguredEntries exists : ${!!m.preconfiguredEntries}`));
                this.log(chalk.gray(`     preconfiguredEntries length : ${m.preconfiguredEntries?.length}`));

                const firstEntry = m.preconfiguredEntries?.[0];
                this.log(chalk.gray(`     firstEntry keys   : ${firstEntry ? Object.keys(firstEntry).join(', ') : 'none'}`));
                this.log(chalk.gray(`     firstEntry.groupId: ${firstEntry?.groupId}`));
                this.log(chalk.gray(`     firstEntry.title  : ${JSON.stringify(firstEntry?.title)}`));

                const rawTitle = firstEntry?.title;
                const title    = typeof rawTitle === 'object' ? rawTitle?.default : rawTitle;
                this.log(chalk.gray(`     resolved title    : ${title}`));

                return {
                  manifestId : m.id                || null,
                  webGroupId : firstEntry?.groupId || null,
                  webTitle   : title               || null
                };
              } catch (e) {
                this.log(chalk.yellow(`  ⚠️  Could not parse ${path.relative(cwd, filePath)}: ${e.message}`));
                return { manifestId: null, webGroupId: null, webTitle: null };
              }
            })
          : null;
      } catch (e) {
        this.log(chalk.yellow(`  ⚠️  Error scanning src/ for manifests: ${e.message}`));
      }
    } else {
      missing.push('src/');
    }

    if (missing.length > 0) {
      this.log(chalk.yellow(`  ⚠️  Could not read values from: ${missing.join(', ')} — placeholders will be used.`));
    }

    // Log what was found
    this.log('');
    this.log(chalk.cyan('  📋 Values read from existing project files:'));
    this.log(chalk.gray(`     solutionName           : ${values.solutionName           ?? '(not found)'}`));
    this.log(chalk.gray(`     solutionId             : ${values.solutionId             ?? '(not found)'}`));
    this.log(chalk.gray(`     solutionPackageZipPath : ${values.solutionPackageZipPath ?? '(not found)'}`));
    this.log(chalk.gray(`     libraryId              : ${values.libraryId              ?? '(not found)'}`));
    this.log(chalk.gray(`     webApiPermissions      : ${values.webApiPermissions      ? values.webApiPermissions.map(p => p.webApiPermName).join(', ') : '(not found)'}`));
    this.log(chalk.gray(`     manifests              : ${values.manifests              ? values.manifests.length + ' found' : '(not found)'}`));
    this.log('');

    return values;
  }

  _stripJsonComments(str) {
    return str
      .replace(/^\uFEFF/, '')                 // strip BOM
      .replace(/\/\*[\s\S]*?\*\//g, '')       // strip /* */ block comments
      .replace(/(^|\s)\/\/[^\n]*/g, '$1')     // strip // comments but not URLs
      .replace(/,(\s*[}\]])/g, '$1');         // strip trailing commas
  }

  _findManifestFiles(dir) {
    let results = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        results = results.concat(this._findManifestFiles(fullPath));
      } else if (entry.isFile() && entry.name.endsWith('.manifest.json')) {
        results.push(fullPath);
      }
    }
    return results.sort(); // alphabetical — same order as env-inject.js uses
  }

  // ─── ENV FILES ───────────────────────────────────────────────────────────────

  _writeEnvFiles() {
    const envDir = this.destinationPath('env');
    if (!fs.existsSync(envDir)) fs.mkdirSync(envDir, { recursive: true });

    const pv   = this.projectValues;
    const envs = ['dev', 'uat', 'prod'];
    const templateFile = this.templatePath('env/env.template.json');

    // Build the template context — fall back to placeholder strings if a value
    // could not be read from the project so the file is still valid JSON.
    const templateContext = {
      solutionName:           pv.solutionName           ?? 'my-solution',
      solutionId:             pv.solutionId             ?? '00000000-0000-0000-0000-000000000000',
      solutionPackageZipPath: pv.solutionPackageZipPath ?? 'my-solution.sppkg',
      webApiPermissions:      pv.webApiPermissions      ?? [{ webApiPermName: 'https://graph.microsoft.com' }],
      libraryId:              pv.libraryId              ?? '00000000-0000-0000-0000-000000000000',
      manifests:              pv.manifests              ?? [{ manifestId: '00000000-0000-0000-0000-000000000000', webGroupId: '00000000-0000-0000-0000-000000000000', webTitle: 'My Web Part' }]
    };

    envs.forEach(env => {
      const dest = this.destinationPath(`env/${env}.env.json`);
      if (!fs.existsSync(dest) || this.answers.overwriteEnvFiles) {
        this.fs.copyTpl(templateFile, dest, { env, ...templateContext });
        this.log(chalk.green(`  ✔ Created env/${env}.env.json`));
      } else {
        this.log(chalk.gray(`  ↷ Skipped env/${env}.env.json (already exists)`));
      }
    });
  }

  // ─── INJECTOR SCRIPT ─────────────────────────────────────────────────────────

  _writeInjectorScript() {
    const dir  = this.destinationPath('scripts');
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const dest = this.destinationPath('scripts/env-inject.js');
    this.fs.copy(this.templatePath('scripts/env-inject.js'), dest);
    this.log(chalk.green('  ✔ Created scripts/env-inject.js'));
  }

  // ─── GULPFILE ────────────────────────────────────────────────────────────────

  _writeGulpfile() {
    const dest = this.destinationPath('gulpfile.js');
    if (!fs.existsSync(dest)) {
      // No gulpfile at all — write the template as a safe starting point
      this.fs.copy(this.templatePath('gulpfile.js'), dest);
      this.log(chalk.green('  ✔ Created gulpfile.js'));
    } else {
      // Always patch existing gulpfile — never replace it
      this._patchExistingGulpfile(dest);
    }
  }

  _patchExistingGulpfile(dest) {
    let content = this.fs.read(dest, { defaults: '' });

    const requireLine  = `const envInject = require('./scripts/env-inject');`;
    const buildHookSig = `build.rig.addPreBuildTask(envInject);`;

    if (content.includes('env-inject')) {
      this.log(chalk.gray('  ↷ gulpfile.js already references env-inject — skipping patch'));
      return;
    }

    // Append before the last line (build.initialize(require('gulp')))
    const initLine = `build.initialize(require('gulp'));`;
    if (content.includes(initLine)) {
      content = content.replace(
        initLine,
        `${requireLine}\n${buildHookSig}\n\n${initLine}`
      );
      this.fs.write(dest, content);
      this.log(chalk.green('  ✔ Patched existing gulpfile.js with env-inject hook'));
    } else {
      this.log(chalk.yellow(
        '  ⚠️  Could not auto-patch gulpfile.js — please add these lines manually:\n' +
        `     ${requireLine}\n` +
        `     ${buildHookSig}`
      ));
    }
  }

  // ─── SCAFFOLD TYPE DECLARATION ───────────────────────────────────────────────

  _scaffoldTypeDeclaration() {
    const typesDir = this.destinationPath('src/types');
    const dest     = this.destinationPath('src/types/env.d.ts');

    if (!fs.existsSync(typesDir)) fs.mkdirSync(typesDir, { recursive: true });

    const pv = this.projectValues;
    const knownKeys = new Set([
      '_environment', '_description', '$schema',
      'solutionName', 'solutionId', 'solutionPackageZipPath',
      'webApiPermissions', 'libraryId', 'manifests'
    ]);

    const customProps = Object.keys(pv)
      .filter(k => !knownKeys.has(k) && !k.startsWith('_'))
      .map(k => {
        const val  = pv[k];
        const type = Array.isArray(val)          ? 'unknown[]'
                   : val === null                ? 'string | null'
                   : typeof val === 'number'     ? 'number'
                   : typeof val === 'boolean'    ? 'boolean'
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

    this.fs.write(dest, content);
    this.log(chalk.green('  ✔ Created src/types/env.d.ts'));
  }

  // ─── SCAFFOLD TSCONFIG ───────────────────────────────────────────────────────

  _scaffoldTsConfig() {
    const dest = this.destinationPath('tsconfig.json');

    if (!fs.existsSync(dest)) {
      this.log(chalk.yellow('  ⚠️  tsconfig.json not found — skipping resolveJsonModule patch'));
      return;
    }

    const raw     = fs.readFileSync(dest, 'utf8');
    const cleaned = raw
      .replace(/^\uFEFF/, '')
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/\/\/[^\n]*/g, '')
      .replace(/,(\s*[}\]])/g, '$1');

    const json = JSON.parse(cleaned);

    if (json.compilerOptions?.resolveJsonModule === true) {
      this.log(chalk.gray('  ↷ tsconfig.json already has resolveJsonModule: true — skipping'));
      return;
    }

    if (!json.compilerOptions) json.compilerOptions = {};
    json.compilerOptions.resolveJsonModule = true;

    this.fs.write(dest, JSON.stringify(json, null, 2) + '\n');
    this.log(chalk.green('  ✔ Patched tsconfig.json → resolveJsonModule: true'));
  }

  // ─── SCAFFOLD ENV CONFIG ─────────────────────────────────────────────────────

  _scaffoldEnvConfig() {
    const configDir = this.destinationPath('src/config');
    const dest      = this.destinationPath('src/config/envConfig.ts');

    if (!fs.existsSync(configDir)) fs.mkdirSync(configDir, { recursive: true });

    const content =
      `// Auto-generated by env-inject.js — do not edit manually.
      // Re-generated on every gulp serve / bundle run to point at the active environment.
      import * as env from '../../env/dev.env.json';
      export default env;`;

    this.fs.write(dest, content);
    this.log(chalk.green('  ✔ Created src/config/envConfig.ts'));
  }

  // ─── GITIGNORE PATCH ─────────────────────────────────────────────────────────

  _patchGitignore() {
    const dest = this.destinationPath('.gitignore');
    let content = fs.existsSync(dest) ? fs.readFileSync(dest, 'utf8') : '';

    const additions = [];
    if (!content.includes('config/*.bak.json'))       
      additions.push('config/*.bak.json');

    if (!content.includes('.yo-rc.bak.json'))          
      additions.push('.yo-rc.bak.json');

    if (!content.includes('src/**/*.bak.json'))        
      additions.push('src/**/*.bak.json');

    if (additions.length > 0) {
      const block = `\n# SPFx env-inject — build-time generated files\n${additions.join('\n')}\n`;
      this.fs.write(dest, content + block);
      this.log(chalk.green('  ✔ Patched .gitignore'));
    } else {
      this.log(chalk.gray('  ↷ .gitignore already up to date'));
    }
  }

  // ─── STEP 3: Install ─────────────────────────────────────────────────────────

  install() {
    this.log(chalk.cyan('\n📦 Checking for required dependencies...\n'));
    // glob is needed by env-inject.js at runtime in the target project
    this.spawnCommandSync('npm', ['install', '--save-dev', 'glob@8']);
  }

  // ─── STEP 4: End ─────────────────────────────────────────────────────────────

  end() {
    this.log(chalk.cyan('\n✅ SPFx env-inject system installed!\n'));
    this.log(chalk.white('Next steps:'));
    this.log(chalk.white('  1. Fill in values in env/dev.env.json, env/uat.env.json, env/prod.env.json'));
    this.log(chalk.white('  2. Run:  gulp serve                            (defaults to dev)'));
    this.log(chalk.white('       or: gulp serve --env uat'));
    this.log(chalk.white('       or: gulp bundle --ship --env prod && gulp package-solution --ship --env prod\n'));
  }
};
