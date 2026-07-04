# E2E Tests

## Layout

- `specs/`: Playwright specs, browser test page, and shared contracts.
- `scripts/`: deterministic Office file generators.
- `files/`: committed DOCX/XLSX/PPTX files served by `/e2e/fixtures/[name]`.

## OnlyOffice Factory Coverage

`specs/onlyoffice-factory.spec.ts` runs the same factory API scenario in both modes:

- local resources from the app origin
- CDN resources from `PLAYWRIGHT_CDN_PORT` or `ONLYOFFICE_E2E_CDN_ORIGIN`

The browser-side scenario lives in `specs/onlyoffice-factory.page.tsx`. The app route at
`src/app/e2e/onlyoffice-factory/page.tsx` should stay a thin wrapper around
`specs/onlyoffice-factory.page.tsx`.

## Fixtures

Office files are committed under `files/`. Regenerate them only when the file
set or edge cases change:

```bash
pnpm test:e2e:files
```

`pnpm test:e2e` and `pnpm test:e2e:ui` read the committed files as-is. The
generator uses the project-local ZIP writer; do not add a ZIP dependency for
these files.

## Running

```bash
pnpm test:e2e
```

Local runs are headed by default so the page is visible at startup. CI runs
headless. Override with `PLAYWRIGHT_HEADLESS=true` or
`PLAYWRIGHT_HEADLESS=false`.
