# generator-spfx-env

A **Yeoman generator** that adds a robust multi-environment config injection system to any
existing SPFx (SharePoint Framework) project with zero manual file setup.

---

## What it does

Injects environment-specific values into SPFx config files **at build time**, before
every `gulp serve`, `gulp bundle`, or `gulp package-solution` run:

| Source (env file) | → | Target file | Target path |
|---|---|---|---|
| `solutionName` | → | `config/package-solution.json` | `solution.name` |
| `solutionId` | → | `config/package-solution.json` | `solution.id` |
| `solutionPackageZipPath` | → | `config/package-solution.json` | `paths.zippedPackage` |
| `webApiPermissions[n].webApiPermName` | → | `config/package-solution.json` | `solution.webApiPermissionRequests[n].resource` |
| `libraryId` | → | `.yo-rc.json` | `@microsoft/generator-sharepoint.libraryId` |
| `manifests[n].manifestId` | → | `*-manifest.schema.json` (nth) | `id` |
| `manifests[n].webGroupId` | → | `*-manifest.schema.json` (nth) | `preconfiguredEntries[*].groupId` |
| `manifests[n].webTitle` | → | `*-manifest.schema.json` (nth) | `preconfiguredEntries[*].title` |

---

## Installation

```bash
# 1. Install Yeoman globally (once)
npm install -g yo

# 2. Link this generator globally
cd generator-spfx-env
npm install
npm link

# 3. Run the generator inside your SPFx project root
cd /path/to/your-spfx-project
yo spfx-env
```

---

## What gets added to your project

```
your-spfx-project/
├── env/
│   ├── dev.env.json          ← Fill in dev values
│   ├── uat.env.json          ← Fill in UAT values
│   ├── prod.env.json         ← Fill in prod values
│   ├── env.schema.json       ← IDE validation schema (committed)
│   └── README.md             ← Usage docs (committed)
├── scripts/
│   └── env-inject.js         ← Gulp task (committed)
├── gulpfile.js               ← Patched or replaced (committed)
└── .gitignore                ← Patched to exclude *.env.json (committed)
```

---

## Developer workflow

```bash
gulp serve --env dev
gulp bundle --ship --env uat && gulp package-solution --ship --env uat
gulp bundle --ship --env prod && gulp package-solution --ship --env prod
```

## CI/CD Pipeline (Azure DevOps example)

```yaml
variables:
  - group: spfx-prod-secrets   # variable group containing PROD_ENV_JSON

steps:
  - script: echo '$(PROD_ENV_JSON)' > env/prod.env.json
    displayName: 'Write env file from secret'

  - script: npm ci
    displayName: 'Install dependencies'

  - script: npx gulp bundle --ship --env prod
    displayName: 'Bundle'

  - script: npx gulp package-solution --ship --env prod
    displayName: 'Package'
```

---

## Notes

- `*.env.json` files are **gitignored** automatically — they contain IDs that differ
  per environment and may contain sensitive values.
- `*.bak.json` backup files are also gitignored (created before each injection).
- The `env.schema.json` file **is** committed — it gives developers IDE autocomplete
  and validation when editing their local env files.
- Manifest files are matched to `manifests[]` entries **by sorted alphabetical order**.
  If you have multiple web parts, ensure your array order matches.
