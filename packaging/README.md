# Packaging manifests for `pptcli`

Templates for the two Windows package managers (issue #112). Everything here is a
**template**: the release workflow renders it, and a human submits it.

| Directory | Channel | Submitted to |
|---|---|---|
| `winget/` | winget | PR to [`microsoft/winget-pkgs`](https://github.com/microsoft/winget-pkgs) |
| `chocolatey/` | Chocolatey | push to [community.chocolatey.org](https://community.chocolatey.org) |

## How a release produces them

1. The `build-cli` job publishes a **framework-dependent** `pptcli` and archives it as
   `PptMcp-CLI-<version>-win-x64.zip`.
2. It computes the SHA256 of that archive and writes `PptMcp-CLI-<version>-win-x64.zip.sha256`.
3. The **Render package manifests** step substitutes `{{VERSION}}`, `{{SHA256}}` and
   `{{RELEASE_DATE}}` into every `*.template` file here and drops the results in
   `packaging-rendered/`.
4. Both the ZIP and `pptcli-packaging-manifests-<version>.zip` (the rendered manifests)
   are attached to the GitHub release.

So submitting a new version is copy-paste from the release assets. Nothing is
hand-edited, and in particular **no checksum is ever typed by hand** - a wrong one
produces an install that fails only on a user's machine.

## Submitting (manual)

**winget** - copy the three rendered files into a fork of `microsoft/winget-pkgs` under
`manifests/t/trsdn/pptcli/<version>/` and open a PR. Validate first:

```powershell
winget validate --manifest <folder>
winget install --manifest <folder>   # requires local manifest support enabled
```

**Chocolatey** - from the rendered `chocolatey/` folder:

```powershell
choco pack
choco push pptcli.<version>.nupkg --source https://push.chocolatey.org/
```

## The .NET dependency is load-bearing

Both manifests declare a hard dependency on the **.NET 9 Desktop Runtime**:

| Channel | Package ID |
|---|---|
| winget | `Microsoft.DotNet.DesktopRuntime.9` |
| Chocolatey | `dotnet-9.0-desktopruntime` |

Two things make this easy to get wrong, and both fail the same silent way - the package
installs cleanly and `pptcli.exe` then refuses to start:

- The artifact is **framework-dependent**, chosen deliberately over self-contained for
  download size. The runtime is therefore not in the box.
- `src/PptMcp.CLI/PptMcp.CLI.csproj` sets `UseWindowsForms=true`, so the *base*
  `Microsoft.DotNet.Runtime.9` / `dotnet-9.0-runtime` package is **not sufficient**. It
  must be the Desktop Runtime.

If the CLI ever switches to a self-contained build, drop these dependencies in the same
change - a stale runtime dependency is a pointless extra install, not a broken one, but
the reverse mistake is user-visible.

> Package IDs were verified against the live registries rather than copied from search
> results, which reported a plausible but non-existent `dotnet-desktopruntime-9.0`.

## What these manifests cannot do

`pptcli` needs **Microsoft PowerPoint installed**, and no package manager can supply it.
Every description here leads with that constraint, and `chocolateyinstall.ps1` probes for
the `PowerPoint.Application` COM registration and warns at install time - otherwise the
first command a user runs fails with an error that reads like a bug in the tool.
