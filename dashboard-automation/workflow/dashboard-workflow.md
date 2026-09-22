# Multi-Project Dashboard Workflow

## New dashboard

1. Interpret project name from user request.
2. Create `projects/<project>/`.
3. Place or locate source files in `input/`.
4. Read optional `config.json`.
5. Read optional `dashboard-request.md`.
6. Run full dashboard pipeline.
7. Write artifacts to that project's `output/`.

## Existing dashboard

1. Select the requested project.
2. Inspect existing output/dashboard.
3. Inspect new input files.
4. Preserve existing implementation where practical.
5. Re-run affected analysis and QA.
6. Update dashboard and reports.

## Multiple projects

Never use files from another project unless explicitly requested.

## Publish preparation

Verify:
- no raw sensitive data is inside public dashboard assets
- dashboard references correct local/public assets
- build/open instructions are documented
- QA passes

Do not actually publish unless the user explicitly requests publishing and
the required credentials/tools are available.
