# OnlyOffice Web Comp

> 📖 English | [中文](README.zh.md)

🌐 **Live Demo**: https://onlyoffice-web-comp.vercel.app/

A browser-based document solution built on the OnlyOffice static SDK. View, edit, and convert Word, Excel, and PowerPoint entirely on the client—**no Document Server required**.

This repository has two parts:

| Part | Path | Description |
|------|------|-------------|
| **Component library** | [`src/components/onlyoffice-web-comp/`](src/components/onlyoffice-web-comp/) | Reusable Web editor wrapper + Markdown docs |
| **Demo site** | [`src/app/`](src/app/) + [`src/features/`](src/features/) | Next.js landing, docs, and live demos |

## Core Advantages

- **Local processing**: Documents stay in the browser
- **Format support**: Word, Excel, PowerPoint, CSV, DOCM, and more
- **No backend**: Host static SDK assets only
- **Engineering APIs**: Read-only/edit toggle, theme, language, multi-instance isolation

## Quick Try

1. Visit the [live demo](https://onlyoffice-web-comp.vercel.app/) or run locally:

```bash
git clone <repository-url>
cd onlyoffice-web-comp
pnpm install
pnpm dev
# http://localhost:3001
```

2. Open a route:

| Route | Description |
|-------|-------------|
| `/` | Product landing page |
| `/docs` | Component library documentation (rendered from Markdown) |
| `/docs/demos/single` | Single-instance editor demo |
| `/docs/demos/multi` | Multi-instance Tab demo |

3. Upload a file → edit → export

Legacy routes `/examples` and `/multi` redirect to the multi-instance demo.

## Component Library Docs

**API details live in the component library docs**, not duplicated here.

- **Entry**: [Component README (English)](src/components/onlyoffice-web-comp/readme.md)
- **Overview**: [docs/概述.md](src/components/onlyoffice-web-comp/docs/概述.md)

| Doc | Topic |
|-----|-------|
| [Quick Start](src/components/onlyoffice-web-comp/docs/快速开始.md) | Init and container mount |
| [Core API](src/components/onlyoffice-web-comp/docs/核心API.md) | `OnlyOfficeManager`, multi-instance |
| [Events](src/components/onlyoffice-web-comp/docs/事件系统.md) | EventBus |
| [Examples](src/components/onlyoffice-web-comp/docs/完整示例.md) | React integration patterns |
| [Reference](src/components/onlyoffice-web-comp/docs/API参考.md) | Constants and types |
| [Notes & Formats](src/components/onlyoffice-web-comp/docs/注意事项与支持格式.md) | Prerequisites and formats |
| [Fonts](src/components/onlyoffice-web-comp/docs/字体配置.md) | Custom font registration |
| [Comments & Revisions](src/components/onlyoffice-web-comp/docs/批注修订与-Word-API.md) | Comments and revisions |
| [Single-instance Demo](src/components/onlyoffice-web-comp/docs/单实例示例.md) | Single editor demo + source |
| [Multi-instance Demo](src/components/onlyoffice-web-comp/docs/多实例示例.md) | Full Tab demo source |

```typescript
import { OnlyOfficeManager, FILE_TYPE, ONLYOFFICE_ID } from "@/components/onlyoffice-web-comp";
```

## Project Structure

```
onlyoffice-web-comp/
├── src/
│   ├── app/                              # Next.js routes
│   │   ├── page.tsx                      # Landing
│   │   ├── docs/                         # Documentation site
│   │   │   ├── page.tsx                  # /docs (overview md)
│   │   │   ├── [slug]/page.tsx           # /docs/*
│   │   │   └── demos/                    # /docs/demos/single|multi
│   │   └── examples/                     # → redirect to multi demo
│   ├── features/
│   │   ├── docs/                         # Docs shell, markdown renderer, site-map
│   │   ├── demo/                         # Live demo components
│   │   ├── marketing/                    # Landing page
│   │   └── shell/                        # Site header / footer / layout
│   └── components/
│       └── onlyoffice-web-comp/          # SDK wrapper + docs/*.md source
├── public/                               # OnlyOffice SDK static assets
└── ...
```

Docs pages read Markdown directly from `src/components/onlyoffice-web-comp/docs/`. Demo tabs embed live editors from `src/features/demo/`.

## Tech Stack

- **OnlyOffice SDK**: Core editing
- **x2t + WebAssembly**: Format conversion
- **Next.js 15 + React 19**: Demo application

## Deployment

```bash
pnpm install
pnpm build
```

Deploy to Vercel or any static host. Live demo: https://onlyoffice-web-comp.vercel.app/

### Deploy SDK Assets to Cloudflare Pages CDN

OnlyOffice SDK assets can be hosted separately from the app. Deploy the contents of `public/packages` to Cloudflare Pages, then register the Pages URL before creating an editor.

```bash
# one-time project creation
npx wrangler pages project create onlyoffice-packages

# upload public/packages as the CDN root
npx wrangler pages deploy public/packages \
  --project-name onlyoffice-packages \
  --commit-dirty=true
```

After deployment, the asset URL should look like:

```text
https://<project>.pages.dev/onlyoffice/9.3.0/web-apps/apps/api/documents/api.js
```

Use that Pages origin as the runtime resource root:

```typescript
import { OnlyOfficeManager } from "@/components/onlyoffice-web-comp";

OnlyOfficeManager.registerStaticResource({
  cdnOrigin: "https://<project>.pages.dev",
});
```

`cdnOrigin` points to the uploaded `public/packages` root, so do not append `/packages`. Cloudflare Pages Direct Upload supports Wrangler folder uploads; dashboard drag-and-drop is less suitable for this repository because the SDK contains many files.

## Fonts

Custom fonts are registered via **`__custom_font_registry__`**, with **`ttf-to-catalog-font.mjs`** producing OnlyOffice catalog wire-format files. See **[Fonts](src/components/onlyoffice-web-comp/docs/字体配置.md)** in the component docs for the full guide.

Quick outline:

1. Run `ttf-to-catalog-font.mjs --id <id> --verify` to produce `fonts/{id}` catalog files
2. Register the id and aliases in `window["__custom_font_registry__"]` inside `AllFonts.js`
3. Ensure aliases cover every font name used in your documents

Ensure all font files comply with applicable licenses.

## Related Resources

- [OnlyOffice API docs](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)

## Contributing

Issues and Pull Requests are welcome.

## License

See [LICENSE](LICENSE).
