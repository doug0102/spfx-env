# SPFx env-inject — Environment Files

This folder contains per-environment variable files used by the `env-inject` gulp task
to patch project config files before every build.

## Files

| File | Purpose |
|---|---|
| `dev.env.json`  | Development values |
| `uat.env.json`  | UAT / staging values |
| `prod.env.json` | Production values |
| `env.schema.json` | JSON Schema — provides IDE validation & autocomplete |
| `README.md` | This file |

> ⚠️ **These files contain environment-specific IDs and secrets.**
> They are excluded from source control via `.gitignore`.
> Each developer and each CI/CD pipeline must supply their own copies.

---

## Variable Reference

| Variable | File Patched | JSON Path |
|---|---|---|
| `solutionName` | `config/package-solution.json` | `solution.name` |
| `solutionId` | `config/package-solution.json` | `solution.id` |
| `solutionPackageZipPath` | `config/package-solution.json` | `paths.zippedPackage` |
| `webApiPermissions[n].webApiPermName` | `config/package-solution.json` | `solution.webApiPermissionRequests[n].resource` |
| `libraryId` | `.yo-rc.json` | `@microsoft/generator-sharepoint.libraryId` |
| `manifests[n].manifestId` | `src/**/*-manifest.schema.json` (nth file) | `id` |
| `manifests[n].webGroupId` | `src/**/*-manifest.schema.json` (nth file) | `preconfiguredEntries[*].groupId` |
| `manifests[n].webTitle` | `src/**/*-manifest.schema.json` (nth file) | `preconfiguredEntries[*].title` |

---

## Usage

```bash
# Local development
gulp serve --env dev

# UAT bundle + package
gulp bundle --ship --env uat
gulp package-solution --ship --env uat

# Production bundle + package
gulp bundle --ship --env prod
gulp package-solution --ship --env prod
```

### CI/CD Pipelines

Store env file contents as **pipeline secrets** and write them to disk before running gulp:

```yaml
# Azure DevOps example
- script: echo '$(PROD_ENV_JSON)' > env/prod.env.json
  displayName: 'Write prod env file'

- script: npx gulp bundle --ship --env prod
  displayName: 'Bundle'

- script: npx gulp package-solution --ship --env prod
  displayName: 'Package'
```

---

## Adding a New Manifest / Web Part

1. Add a new entry to the `manifests` array in each `*.env.json`.
2. The injector matches manifest entries **by array index** to the alphabetically-sorted list
   of `*-manifest.schema.json` files found under `src/`.
3. Ensure your new entry has `manifestId`, `webGroupId`, and `webTitle`.

## Backup Files

Before patching, the injector creates `*.bak.json` backups of each file it modifies
(e.g. `config/package-solution.bak.json`). These are also excluded from source control.
