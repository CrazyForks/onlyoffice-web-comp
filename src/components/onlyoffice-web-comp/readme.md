# OnlyOffice Web Comp

> 📖 English | [中文](readme.zh.md)

A **browser-side document editor component library** built on the OnlyOffice static SDK. Supports online editing, read-only preview, export, and x2t conversion for Word, Excel, and PowerPoint. **No self-hosted Document Server required**—only static SDK assets on your site.

> This file is an **entry point**. Full documentation lives in [`docs/`](./docs/00-概述.md). On the demo site, these Markdown files are rendered at `/docs`.

## Documentation

| # | Doc | Description |
|---|-----|-------------|
| 00 | [Overview](./docs/00-概述.md) | Index and reading path |
| 01 | [Quick Start](./docs/01-快速开始.md) | Init, container mount, create editor |
| 02 | [Core API](./docs/02-核心API.md) | `OnlyOfficeManager`, `EditorManager`, multi-instance |
| 03 | [Event System](./docs/03-事件系统.md) | EventBus, event types, listeners |
| 04 | [Full Examples](./docs/04-完整示例.md) | React integration patterns |
| 05 | [API Reference](./docs/05-API参考.md) | Constants and types |
| 06 | [Notes & Formats](./docs/06-注意事项与支持格式.md) | Prerequisites, formats, pitfalls |
| 07 | [Comments, Revisions & Word API](./docs/07-批注修订与-Word-API.md) | Comments, revisions, SDK callbacks |
| 08 | [Single-instance Demo](./docs/08-单实例示例.md) | Single editor demo + source walkthrough |
| 09 | [Multi-instance Demo](./docs/09-多实例示例.md) | Full Tab demo source |

**Suggested paths**

| Scenario | Path |
|----------|------|
| First integration | [01](./docs/01-快速开始.md) → [02](./docs/02-核心API.md) |
| Try live demos | [08](./docs/08-单实例示例.md) · [09](./docs/09-多实例示例.md) |
| React page integration | [04](./docs/04-完整示例.md) |
| Multi-instance / export | [02](./docs/02-核心API.md) · [03](./docs/03-事件系统.md) |

## Package Layout

```
onlyoffice-web-comp/
├── const/       Constants, static paths, file types, themes
├── store/       Document / language state
├── util/        SDK init, x2t conversion, download
├── core/        EditorManager, OnlyOfficeManager, EventBus
├── feature/     Comments, revisions
├── docs/        Full documentation (Markdown source of truth)
└── internal/    Mock server / x2t worker (not exported)
```

## Minimal Example

```typescript
import {
  OnlyOfficeManager,
  ONLYOFFICE_ID,
  FILE_TYPE,
} from "@/components/onlyoffice-web-comp";

// Create a blank document
const manager = await OnlyOfficeManager.create({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.DOCX,
  defaultFileName: "New_Document.docx",
});

// Open an existing File (fetch first, then mount)
const file = await fetch("/test.xlsx").then((r) => r.blob())
  .then((blob) => new File([blob], "test.xlsx", { type: blob.type }));
await OnlyOfficeManager.createWithFile({
  containerId: ONLYOFFICE_ID,
  fileType: FILE_TYPE.XLSX,
  defaultFileName: "test.xlsx",
}, file);
```

See [docs/02-核心API.md](./docs/02-核心API.md) for events, export, multi-instance, theme, language, and read-only toggling.

## Demos in This Repo

Documentation Markdown under `docs/` is rendered by the demo site. Live editors are embedded on the demo pages below.

| Route | Description |
|-------|-------------|
| `/docs` | All docs from this `docs/` folder |
| `/docs/demos?tab=single` | Single-instance demo ([08](./docs/08-单实例示例.md)) |
| `/docs/demos?tab=multi` | Multi-instance Tab demo ([09](./docs/09-多实例示例.md)) |

Demo components: `src/features/demo/` (`office-preview-page.tsx`, `tabs-multi-page.tsx`)

Run locally: `pnpm dev` → http://localhost:3001

## Links

- [Repository README (project overview)](../../../README.md)
- [OnlyOffice official API](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
