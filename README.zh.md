# OnlyOffice Web Comp

> 📖 [English](README.md) | 中文

🌐 **在线演示**: [https://onlyoffice-web-comp.vercel.app/](https://onlyoffice-web-comp.vercel.app/)

基于 OnlyOffice 静态 SDK 的浏览器端文档处理方案：在客户端完成 Word / Excel / PowerPoint 的查看、编辑与转换，**无需 Document Server**。

本仓库包含两部分：

| 部分 | 路径 | 说明 |
|------|------|------|
| **组件库** | [`src/components/onlyoffice-web-comp/`](src/components/onlyoffice-web-comp/) | 可复用的 Web 端编辑器封装 + Markdown 文档源 |
| **演示站点** | [`src/app/`](src/app/) + [`src/features/`](src/features/) | Next.js 主页、文档站、在线示例 |

## 核心优势

- **数据留在本地**：文档处理在浏览器内完成
- **格式兼容**：Word、Excel、PowerPoint、CSV、DOCM 等
- **零后端**：托管静态 SDK 即可使用
- **工程化 API**：只读/编辑、主题/语言、多实例容器隔离

## 快速体验

1. 访问 [在线演示](https://onlyoffice-web-comp.vercel.app/) 或本地启动：

```bash
git clone <repository-url>
cd onlyoffice-web-comp
pnpm install
pnpm dev
# http://localhost:3001
```

2. 打开路由：

| 路由 | 说明 |
|------|------|
| `/` | 产品主页 |
| `/docs` | 组件库文档（直接渲染 Markdown） |
| `/docs/demos?tab=single` | 单实例在线示例 |
| `/docs/demos?tab=multi` | 多实例 Tab 在线示例 |

3. 上传本地文件 → 编辑 → 导出

旧路由 `/examples`、`/multi` 会重定向到多实例示例 Tab。

## 组件库文档

**API 与接入说明不在本 README 重复**，请阅读组件库文档：

- **入口**：[组件库 README（中文）](src/components/onlyoffice-web-comp/readme.zh.md)
- **概述**：[docs/00-概述.md](src/components/onlyoffice-web-comp/docs/00-概述.md)

| 文档 | 内容 |
|------|------|
| [01-快速开始](src/components/onlyoffice-web-comp/docs/01-快速开始.md) | 初始化与容器挂载 |
| [02-核心API](src/components/onlyoffice-web-comp/docs/02-核心API.md) | `OnlyOfficeManager`、多实例 |
| [03-事件系统](src/components/onlyoffice-web-comp/docs/03-事件系统.md) | EventBus |
| [04-完整示例](src/components/onlyoffice-web-comp/docs/04-完整示例.md) | React 集成模式 |
| [05-API参考](src/components/onlyoffice-web-comp/docs/05-API参考.md) | 常量与类型 |
| [06-注意事项与格式](src/components/onlyoffice-web-comp/docs/06-注意事项与支持格式.md) | 前置条件与格式 |
| [10-字体配置](src/components/onlyoffice-web-comp/docs/10-字体配置.md) | 自定义字体注册 |
| [07-批注修订](src/components/onlyoffice-web-comp/docs/07-批注修订与-Word-API.md) | 批注、修订 |
| [08-单实例示例](src/components/onlyoffice-web-comp/docs/08-单实例示例.md) | 单实例 Demo 与源码说明 |
| [09-多实例示例](src/components/onlyoffice-web-comp/docs/09-多实例示例.md) | Tab 多实例完整源码 |

```typescript
import { OnlyOfficeManager, FILE_TYPE, ONLYOFFICE_ID } from "@/components/onlyoffice-web-comp";
```

## 项目结构

```
onlyoffice-web-comp/
├── src/
│   ├── app/                              # Next.js 路由
│   │   ├── page.tsx                      # 主页
│   │   ├── docs/                         # 文档站
│   │   │   ├── page.tsx                  # /docs（概述 md）
│   │   │   ├── [slug]/page.tsx           # /docs/*
│   │   │   └── demos/page.tsx            # /docs/demos?tab=single|multi
│   │   └── examples/                     # → 重定向至多实例示例
│   ├── features/
│   │   ├── docs/                         # 文档壳、Markdown 渲染、site-map
│   │   ├── demo/                         # 在线演示组件
│   │   ├── marketing/                    # 着陆页
│   │   └── shell/                        # 站点 Header / Footer / Layout
│   └── components/
│       └── onlyoffice-web-comp/          # SDK 封装 + docs/*.md 文档源
├── public/                               # OnlyOffice SDK 静态资源
└── ...
```

文档页直接读取 `src/components/onlyoffice-web-comp/docs/` 下的 Markdown；示例 Tab 内嵌 `src/features/demo/` 的可交互编辑器。

## 技术栈

- **OnlyOffice SDK**：文档编辑核心
- **x2t + WebAssembly**：格式转换
- **Next.js 15 + React 19**：演示应用

## 部署

```bash
pnpm install
pnpm build
```

可部署至 Vercel 或任意静态托管。演示地址：[https://onlyoffice-web-comp.vercel.app/](https://onlyoffice-web-comp.vercel.app/)

## 字体配置

自定义字体通过 **`__custom_font_registry__`** 注册，配合 **`ttf-to-catalog-font.mjs`** 生成 OnlyOffice catalog 线格式。完整步骤见组件库文档 **[10 - 字体配置](src/components/onlyoffice-web-comp/docs/10-字体配置.md)**。

简要流程：

1. 运行 `ttf-to-catalog-font.mjs --id <id> --verify` 生成 `fonts/{id}` catalog 文件
2. 在 `AllFonts.js` 的 `window["__custom_font_registry__"]` 中注册 id 与别名
3. 确保别名覆盖文档内实际使用的字体名

请确保所用字体文件符合相关许可协议。

## 相关资源

- [OnlyOffice API 文档](https://api.onlyoffice.com/zh-CN/docs/docs-api/usage-api/config/document/)
- [OnlyOffice Web Apps](https://github.com/ONLYOFFICE/web-apps)
- [OnlyOffice SDK](https://github.com/ONLYOFFICE/sdkjs)
- [x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)

## 参与贡献

欢迎提交 Issue 和 Pull Request。

## 开源许可

详见 [LICENSE](LICENSE)。
