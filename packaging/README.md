# Portable Windows Packages

## Public Package

```powershell
python scripts/build_portable_client.py --public
```

The public ZIP includes Python and runtime libraries but no workbook, customer
mapping, uploads or logs. It starts at the existing onboarding screen. Users
upload their own Excel file, review the suggested mapping, validate it and create
a profile. Excel desktop is required for recalculation. No manual `.env` editing
is needed. On later visits, select the saved profile and reselect the workbook.

`--public` rejects all client workbook/profile arguments. Only the clean
`Excel_AI_Controller_Public_*.zip` is intended for GitHub Releases. Never
re-zip a folder that has been used: it can contain user data.

## Private Client Package

The public repository contains packaging code only. Supply your workbook and
mapping locally; they are deliberately not included in Git. Create a mapping
using the app's onboarding workflow before building a client package.

Build on Windows x64 with Python 3.13 and pip:

```powershell
python scripts/build_portable_client.py --model models/base_model.xlsx --profile-map config/profiles/client_model/model_map.yaml --profile client_model
```

The builder downloads a checksum-pinned Python 3.13.7 embeddable runtime and
installs pinned runtime libraries. Internet is required on the build machine.
Client machines do not need Python, pip or Internet for the default local parser.
Microsoft Excel desktop is still required for formula recalculation.

The builder copies only application code, web assets, a clean configuration
template, and the explicitly supplied workbook/mapping. It never copies the
developer's `.env`. Output ZIPs, extracted packages, checksums and diagnostic
reports stay under the ignored `outputs/` directory. Do not upload private client
packages as public GitHub releases: they contain client workbook data.

## Verification

Verify a locally built archive (requires Excel on the verification machine):

```powershell
python scripts/verify_portable_client.py outputs/portable/Excel_AI_Controller_Client_TIMESTAMP.zip
```

The verifier checks archive hashes, configuration, isolated imports, the BAT
launcher, port conflict handling, and an Excel scenario/export without changing
the source workbook. Public packages also undergo empty-state and synthetic
workbook tests: upload, mapping validation, profile creation, recalculation,
download and profile reuse. Synthetic workbooks only live in the verification
directory, never in the release ZIP.
An optional `--smoke-script PATH` runs a local API smoke
script with both Excel engines. Verification reports remain local.

Clients extract the entire ZIP into a writable folder and open
`Excel AI Controller.bat`. The launcher binds only to loopback and uses bundled
Python with user-site packages disabled. This is not a signed Windows installer.

Customer mappings formerly tracked under `config/` are no longer included in
new commits. Removing tracked paths does not erase earlier Git history.
