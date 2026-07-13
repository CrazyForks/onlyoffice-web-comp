import path from "node:path";
import { expect, test } from "playwright/test";

type OnlyOfficeRuntimeNamespace = Record<string, any>;

declare global {
  interface Window {
    Asc?: OnlyOfficeRuntimeNamespace;
    AscCommon?: OnlyOfficeRuntimeNamespace;
    AscFonts?: OnlyOfficeRuntimeNamespace;
    Common?: OnlyOfficeRuntimeNamespace;
    DE?: OnlyOfficeRuntimeNamespace;
  }
}

const assetOrigin = `http://${process.env.PLAYWRIGHT_HOST ?? "127.0.0.1"}:${
  process.env.PLAYWRIGHT_CDN_PORT ?? 3010
}`;
const deRoot = `${assetOrigin}/onlyoffice/9.4.0-develop`;

const sampleCanvasChecksum = (element: HTMLElement | SVGElement) => {
  const canvas = element as HTMLCanvasElement;
  const context = canvas.getContext("2d");
  const data = context?.getImageData(0, 0, canvas.width, canvas.height).data;
  if (!data) return 0;
  let checksum = 0;
  for (let index = 0; index < data.length; index += 16) {
    checksum = (checksum + data[index] + data[index + 1] + data[index + 2]) >>> 0;
  }
  return checksum;
};

const clickHtmlElement = (element: HTMLElement | SVGElement) => {
  (element as HTMLElement).click();
};

test("9.4 DE serves root plugin and theme configs", async ({ page }) => {
  await page.goto(`${deRoot}/web-apps/apps/api/documents/preload.html`, {
    waitUntil: "networkidle",
  });
  const status = await page.evaluate(async () =>
    Promise.all(
      ["plugins.json", "themes.json"].map(async (fileName) => ({
        fileName,
        status: (await fetch(`../../../../${fileName}`)).status,
      })),
    ),
  );
  expect(status).toEqual([
    { fileName: "plugins.json", status: 200 },
    { fileName: "themes.json", status: 200 },
  ]);
});

test("9.4 DE registers the migrated custom font catalog", async ({ page }) => {
  await page.goto(`${deRoot}/web-apps/apps/api/documents/preload.html`, {
    waitUntil: "networkidle",
  });

  const result = await page.evaluate(async () => {
    await new Promise<void>((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "../../../../sdkjs/common/AllFonts.js";
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("Failed to load AllFonts.js"));
      document.head.append(script);
    });

    const win = window as unknown as {
      __custom_font_registry__: Record<string, string[]>;
      __fonts_files: string[];
      __fonts_infos: Array<[string, number]>;
    };
    const ids = ["1002", "1003", "1004", "1005"];
    const names = [
      "FZXiaoBiaoSong-B05S",
      "FangSong_GB2312",
      "SimHei",
      "KaiTi_GB2312",
    ];
    const fontStatus = await Promise.all(
      ids.map((id) => fetch(`../../../../fonts/${id}`).then((response) => response.status)),
    );

    return {
      registry: ids.every((id) => Array.isArray(win.__custom_font_registry__?.[id])),
      catalogFiles: ids.every((id) => win.__fonts_files.includes(id)),
      catalogNames: names.every((name) =>
        win.__fonts_infos.some(([fontName]) => fontName === name),
      ),
      fontStatus,
    };
  });

  expect(result).toEqual({
    registry: true,
    catalogFiles: true,
    catalogNames: true,
    fontStatus: [200, 200, 200, 200],
  });
});

test("9.4 DE presenter view installs the opener bridge before RequireJS", async ({
  page,
}) => {
  await page.goto(`${deRoot}/web-apps/apps/api/documents/preload.html`, {
    waitUntil: "networkidle",
  });
  await page.evaluate(() => {
    const win = window as unknown as {
      __ONLYOFFICE_REPORTER_BRIDGE__: { install: (target: Window) => void };
    };
    win.__ONLYOFFICE_REPORTER_BRIDGE__ = {
      install(target) {
        (target as Window & { __reporterBridgeInstalled?: boolean })
          .__reporterBridgeInstalled = true;
      },
    };
  });

  const popupPromise = page.waitForEvent("popup");
  await page.evaluate(() =>
    window.open(
      "/onlyoffice/9.4.0-develop/web-apps/apps/presentationeditor/main/index.reporter.html",
      "_blank",
    ),
  );
  const popup = await popupPromise;
  await popup.waitForLoadState("domcontentloaded");

  await expect
    .poll(() =>
      popup.evaluate(() => {
        const win = window as Window & {
          __reporterBridgeInstalled?: boolean;
          require?: unknown;
        };
        return {
          injected: win.__reporterBridgeInstalled === true,
          require: typeof win.require,
        };
      }),
    )
    .toEqual({ injected: true, require: "function" });

  await popup.close();
});

test("9.4 demo calls the Developer Edition connector", async ({ page }) => {
  await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/")),
    )
    .toBe(true);

  const editorFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/"));
  if (!editorFrame) {
    throw new Error("The 9.4 Word editor iframe did not load");
  }
  // iframe 创建完成不代表编辑器已开始处理 Connector 消息；等待文档画布就绪。
  await expect(editorFrame.locator("#id_viewer_overlay")).toBeVisible({
    timeout: 60_000,
  });

  const documentCanvas = editorFrame.locator("#id_viewer");
  const before = await documentCanvas.evaluate(sampleCanvasChecksum);

  await page.getByRole("button", { name: "连接器写入" }).click();
  await expect(page.getByRole("status")).toHaveText(
    "Connector: inserted a paragraph",
    { timeout: 10_000 },
  );
  await expect
    .poll(() =>
      documentCanvas.evaluate(sampleCanvasChecksum),
    )
    .not.toBe(before);
});

test("9.4 Word exports PDF through the compatible x2t pair", async ({ page }) => {
  const x2tRequests: string[] = [];
  const exportErrors: string[] = [];
  page.on("request", (request) => {
    if (request.url().includes("/x2t/")) {
      x2tRequests.push(request.url());
    }
  });
  page.on("console", (message) => {
    if (
      message.type() === "error" &&
      /x2t conversion produced no output|downloadAs failed/.test(message.text())
    ) {
      exportErrors.push(message.text());
    }
  });

  await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/")),
    )
    .toBe(true);
  const editorFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/"));
  if (!editorFrame) {
    throw new Error("The 9.4 Word editor iframe did not load");
  }
  await expect(editorFrame.locator("#id_viewer_overlay")).toBeVisible({
    timeout: 60_000,
  });

  await editorFrame.getByRole("tab", { name: "文件" }).click();
  await editorFrame.locator("a.menu-item").filter({ hasText: "下载为" }).click();
  const downloadPromise = page.waitForEvent("download", { timeout: 60_000 });
  await editorFrame.locator('.btn-doc-format[format="513"]').click();
  const download = await downloadPromise;

  expect(download.suggestedFilename()).toMatch(/\.pdf$/i);
  await expect.poll(() => download.failure()).toBeNull();
  expect(x2tRequests).toEqual(
    expect.arrayContaining([
      expect.stringContaining("/packages/onlyoffice/9.3.0/x2t/x2t.js"),
      expect.stringContaining("/packages/onlyoffice/9.3.0/x2t/x2t.wasm"),
    ]),
  );
  expect(exportErrors).toEqual([]);
});

test("9.4 multi-instance demo keeps one connector per editor", async ({ page }) => {
  await page.goto("/docs/demos/multi", { waitUntil: "domcontentloaded" });
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/")),
    )
    .toBe(true);

  await page.getByRole("button", { name: "连接器写入" }).click();
  await expect(page.getByRole("status")).toHaveText(
    "Connector: inserted a paragraph",
    { timeout: 10_000 },
  );

  await page.getByTitle("新建 Excel 标签页").click();
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);
  const spreadsheetFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/"));
  if (!spreadsheetFrame) {
    throw new Error("The multi-instance Excel editor did not load");
  }
  await expect(spreadsheetFrame.locator("#editor_sdk")).toBeVisible({
    timeout: 60_000,
  });
  await page.getByRole("button", { name: "连接器写入" }).click();
  await expect(page.getByRole("status")).toHaveText("Connector: wrote to A1", {
    timeout: 10_000,
  });
});

test("9.4 spreadsheet paste filters scripts before writing its sandboxed frame", async ({
  page,
}) => {
  await page.goto("/docs/demos/multi", { waitUntil: "domcontentloaded" });
  await page.getByTitle("新建 Excel 标签页").click();
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/spreadsheeteditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);

  const spreadsheetFrame = page
    .frames()
    .find((frame) => frame.url().includes("/spreadsheeteditor/main/"));
  if (!spreadsheetFrame) {
    throw new Error("The spreadsheet editor iframe did not load");
  }
  await expect(spreadsheetFrame.locator("#editor_sdk")).toBeVisible({
    timeout: 60_000,
  });
  await expect
    .poll(() =>
      spreadsheetFrame.evaluate(() => {
        const win = window as unknown as {
          AscCommon?: { PL?: { "p$b"?: unknown; Wb?: unknown } };
        };
        return (
          typeof win.AscCommon?.PL?.["p$b"] === "function" &&
          !!win.AscCommon.PL.Wb
        );
      }),
      { timeout: 60_000 },
    )
    .toBe(true);

  const result = await spreadsheetFrame.evaluate(() => {
    const win = window as unknown as {
      AscCommon: { PL: { "p$b": (html: string, text: string) => void } };
      __pasteScriptExecuted?: boolean;
    };
    win.__pasteScriptExecuted = false;
    win.AscCommon.PL["p$b"](
      '<img alt="clipboard-image" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==" onload="window.__pasteScriptExecuted=true"><iframe srcdoc="<script>window.__pasteScriptExecuted=true</script>"></iframe><script>window.__pasteScriptExecuted=true</script>',
      "",
    );
    const pasteFrame = document.getElementById(
      "asc_pasteFrame",
    ) as HTMLIFrameElement | null;
    return {
      sandbox: pasteFrame?.getAttribute("sandbox"),
      scriptExecuted: win.__pasteScriptExecuted,
      html: pasteFrame?.contentDocument?.body.innerHTML ?? "",
    };
  });

  expect(result.sandbox).toBe("allow-same-origin");
  expect(result.scriptExecuted).toBe(false);
  expect(result.html).not.toContain("<script");
  expect(result.html).not.toContain("onload=");
  expect(result.html).not.toContain("<iframe");
});

test("9.4 Word selects the 仿宋_GB2312 custom-font alias without invalid thumbnails", async ({
  page,
}) => {
  const runtimeErrors: string[] = [];
  page.on("console", (message) => {
    if (message.type() === "error") runtimeErrors.push(message.text());
  });
  page.on("pageerror", (error) => runtimeErrors.push(error.message));

  await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/documenteditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);

  const documentFrame = page
    .frames()
    .find((frame) => frame.url().includes("/documenteditor/main/"));
  if (!documentFrame) {
    throw new Error("The document editor iframe did not load");
  }
  await expect
    .poll(() =>
      documentFrame.evaluate(() =>
        Boolean(window.Common?.UI?.ComboBoxFonts),
      ),
    )
    .toBe(true);

  const fontInput = documentFrame.locator("#slot-field-fontname input");
  await fontInput.click();
  await fontInput.press("ArrowDown");
  await fontInput.fill("仿宋_GB2312");
  await fontInput.press("Enter");
  await page.waitForTimeout(500);

  await expect(fontInput).toHaveValue("仿宋_GB2312");
  expect(
    runtimeErrors.filter((message) =>
      message.includes("Invalid typed array length"),
    ),
  ).toEqual([]);
});

test("9.4 Cell and Slide keep native fonts while resolving 方正小标宋简体", async ({
  page,
}) => {
  let customFontBinaryLoaded = false;
  const runtimeErrors: string[] = [];
  page.on("response", (response) => {
    if (/\/fonts\/1002(?:[?#]|$)/.test(response.url()) && response.ok()) {
      customFontBinaryLoaded = true;
    }
  });
  page.on("pageerror", (error) => runtimeErrors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") runtimeErrors.push(message.text());
  });
  await page.goto("/docs/demos/multi", { waitUntil: "domcontentloaded" });

  await page.getByTitle("新建 Excel 标签页").click();
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);
  const spreadsheetFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/"));
  if (!spreadsheetFrame) throw new Error("The Excel editor did not load");
  await expect(spreadsheetFrame.locator("#editor_sdk")).toBeVisible({
    timeout: 60_000,
  });
  await expect
    .poll(() =>
      spreadsheetFrame.evaluate(() => ({
        patched: window.AscFonts?.__CUSTOM_PICK_FONT_PATCHED__ === true,
        resolver:
          window.AscFonts?.JZ?.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__ === true,
        custom: window.AscFonts?.pickFont?.("方正小标宋简体")?.Xa,
        native: window.AscFonts?.JZ
          ?.rM?.("方正小标宋简体")
          ?.SEe?.(window.AscCommon?.O4, 0)?.file?.Xa,
        arial: window.AscFonts?.pickFont?.("Arial")?.Xa,
      })),
    )
    .toEqual({
      patched: true,
      resolver: true,
      custom: "1002",
      native: "1002",
      arial: "022",
    });

  // 使用编辑器公开 API 验证实际异步加载器会请求 custom binary；避免刚创建
  // tab 时的初始化遮罩让 UI click 的时序主导测试结果。
  expect(
    await spreadsheetFrame.evaluate(() => {
      const editor = window.Asc?.editor as {
        asc_loadFontsFromServer?: (fontNames: string[]) => void;
      };
      if (typeof editor.asc_loadFontsFromServer !== "function") {
        return false;
      }
      editor.asc_loadFontsFromServer(["方正小标宋简体"]);
      return true;
    }),
  ).toBe(true);

  await page.getByTitle("新建 PPT 标签页").click();
  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/presentationeditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);
  const presentationFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/presentationeditor/main/"));
  if (!presentationFrame) throw new Error("The PPT editor did not load");
  await expect(presentationFrame.locator("#editor_sdk")).toBeVisible({
    timeout: 60_000,
  });
  await expect
    .poll(() =>
      presentationFrame.evaluate(() => ({
        patched: window.AscFonts?.__CUSTOM_PICK_FONT_PATCHED__ === true,
        resolver:
          window.AscFonts?.LU?.__CUSTOM_FONT_REGISTRY_RESOLVER_PATCHED__ === true,
        picker: window.AscFonts?.LU?.gLc?.("方正小标宋简体")?.fba,
        custom: window.AscFonts?.pickFont?.("方正小标宋简体")?.Na,
        native: window.AscFonts?.LU
          ?.FG?.("方正小标宋简体")
          ?.Gee?.(window.AscCommon?.jY, 0)?.file?.Na,
        arial: window.AscFonts?.pickFont?.("Arial")?.Na,
      })),
    )
    .toEqual({
      patched: true,
      resolver: true,
      picker: "方正小标宋简体",
      custom: "1002",
      native: "1002",
      arial: "022",
    });
  expect(
    await presentationFrame.evaluate(() => {
      const editor = window.Asc?.editor as {
        asc_loadFontsFromServer?: (fontNames: string[]) => void;
      };
      if (typeof editor.asc_loadFontsFromServer !== "function") {
        return false;
      }
      editor.asc_loadFontsFromServer(["方正小标宋简体"]);
      return true;
    }),
  ).toBe(true);
  await expect.poll(() => customFontBinaryLoaded).toBe(true);
  expect(runtimeErrors).toEqual([]);
});

test("9.4 demo writes A1 after uploading an Excel workbook", async ({ page }) => {
  let customFontLoadedFromWorkbook = false;
  page.on("response", (response) => {
    if (/\/fonts\/1003(?:[?#]|$)/.test(response.url()) && response.ok()) {
      customFontLoadedFromWorkbook = true;
    }
  });
  await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });
  const fileChooser = page.waitForEvent("filechooser");
  await page.getByRole("button", { name: "上传" }).click();
  await (await fileChooser).setFiles(path.join(process.cwd(), "public/test.xlsx"));

  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/")),
      { timeout: 60_000 },
    )
    .toBe(true);
  const editorFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/spreadsheeteditor/main/"));
  if (!editorFrame) {
    throw new Error("The uploaded XLSX did not open in the spreadsheet editor");
  }
  await expect(editorFrame.locator("#editor_sdk")).toBeVisible({
    timeout: 60_000,
  });
  const worksheetCanvas = editorFrame.locator("#ws-canvas");
  await expect(worksheetCanvas).toBeVisible();
  await expect(editorFrame.locator('input[aria-label="字体 "]')).toHaveValue(
    "仿宋_GB2312",
  );
  await expect
    .poll(() =>
      editorFrame.evaluate(() => ({
        resolvedName: window.AscFonts?.JZ?.xMd?.("仿宋_GB2312")?.Rma,
        resolvedFile: window.AscFonts?.JZ
          ?.rM?.("仿宋_GB2312")
          ?.SEe?.(window.AscCommon?.O4, 0)?.file?.Xa,
      })),
    )
    .toEqual({ resolvedName: "仿宋_GB2312", resolvedFile: "1003" });
  await expect.poll(() => customFontLoadedFromWorkbook).toBe(true);

  const before = await worksheetCanvas.evaluate(sampleCanvasChecksum);

  // public/test.xlsx 的活动单元格 C4 已使用仿宋_GB2312。切换到 Arial 后
  // 画布必须变化；再切回 custom font 必须恢复原图，覆盖文件加载后的排版路径。
  await editorFrame.evaluate(() => window.Asc?.editor?.asc_setCellFontName?.("Arial"));
  await expect
    .poll(() =>
      worksheetCanvas.evaluate(sampleCanvasChecksum),
    )
    .not.toBe(before);
  await editorFrame.evaluate(() =>
    window.Asc?.editor?.asc_setCellFontName?.("仿宋_GB2312"),
  );
  await expect
    .poll(() =>
      worksheetCanvas.evaluate(sampleCanvasChecksum),
    )
    .toBe(before);

  await page.getByRole("button", { name: "连接器写入" }).click();
  await expect(page.getByRole("status")).toHaveText("Connector: wrote to A1", {
    timeout: 10_000,
  });
  await expect
    .poll(() =>
      worksheetCanvas.evaluate(sampleCanvasChecksum),
    )
    .not.toBe(before);
});

test("9.4 Word loads built-in and 方正小标宋简体 fonts", async ({ page }) => {
  let customFontBinaryLoaded = false;
  let builtInFontBinaryLoaded = false;
  page.on("response", (response) => {
    if (
      /\/fonts\/\d+(?:[?#]|$)/.test(response.url()) &&
      !/\/fonts\/100[2-5](?:[?#]|$)/.test(response.url()) &&
      response.ok()
    ) {
      builtInFontBinaryLoaded = true;
    }
    if (/\/fonts\/1002(?:[?#]|$)/.test(response.url()) && response.ok()) {
      customFontBinaryLoaded = true;
    }
  });

  await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });

  await expect
    .poll(() =>
      page
        .frames()
        .some((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/")),
    )
    .toBe(true);

  const editorFrame = page
    .frames()
    .find((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/"));
  if (!editorFrame) {
    throw new Error("The 9.4 Word editor iframe did not load");
  }

  const fontCombo = editorFrame.locator(".combobox.fonts");
  const fontInput = fontCombo.locator('input[role="combobox"]');
  await expect(fontCombo).toBeVisible({ timeout: 60_000 });
  await expect
    .poll(() =>
      editorFrame.evaluate(
        () =>
          !!(
            window.Common?.UI?.ComboBoxFonts?.prototype as {
              __CUSTOM_FONT_REGISTRY_COMBO_PATCHED__?: boolean;
            }
          )?.__CUSTOM_FONT_REGISTRY_COMBO_PATCHED__,
      ),
    )
    .toBe(true);
  await fontCombo.locator(".dropdown-toggle").click();
  await expect(fontInput).toHaveAttribute("aria-expanded", "true");

  const customFontItemId = await editorFrame.evaluate(() =>
    window.DE?.getController?.("Toolbar")
      ?.getView?.("Toolbar")
      ?.cmbFontName?.store?.findWhere?.({ name: "方正小标宋简体" })
      ?.get?.("id"),
  );
  expect(customFontItemId).toEqual(expect.any(String));

  const customFontItem = fontCombo.locator(
    `li[id="${customFontItemId}"] .font-item`,
  );
  await expect(customFontItem).toHaveCount(1);
  await customFontItem.evaluate(clickHtmlElement);
  await expect(fontInput).toHaveValue("方正小标宋简体");

  const documentCanvas = editorFrame.locator("#id_viewer");
  const documentOverlay = editorFrame.locator("#id_viewer_overlay");
  await expect(documentCanvas).toBeVisible();
  await expect(documentOverlay).toBeVisible();
  const before = await documentCanvas.evaluate(sampleCanvasChecksum);

  await documentOverlay.click({ position: { x: 120, y: 100 } });
  await page.keyboard.type("方正小标宋简体");
  await expect
    .poll(() =>
      documentCanvas.evaluate(sampleCanvasChecksum),
    )
    .not.toBe(before);
  await expect.poll(() => builtInFontBinaryLoaded).toBe(true);
  await expect.poll(() => customFontBinaryLoaded).toBe(true);
});

test("9.4 Word rasterizes 方正小标宋简体 instead of the fallback font", async ({ browser }) => {
  const render = async (fontName?: string) => {
    const page = await browser.newPage();
    let customFontRequests = 0;
    page.on("response", (response) => {
      if (/\/fonts\/1002(?:[?#]|$)/.test(response.url()) && response.ok()) {
        customFontRequests += 1;
      }
    });
    await page.goto("/docs/demos/single", { waitUntil: "domcontentloaded" });
    await expect
      .poll(() =>
        page
          .frames()
          .some((frame) => frame.url().includes("/web-apps/apps/documenteditor/main/")),
      )
      .toBe(true);
    const frame = page
      .frames()
      .find((item) => item.url().includes("/web-apps/apps/documenteditor/main/"));
    if (!frame) throw new Error("The 9.4 Word editor iframe did not load");
    const overlay = frame.locator("#id_viewer_overlay");
    await expect(overlay).toBeVisible();
    await frame.evaluate(() => {
      const trace: string[] = [];
      const catalog = window.AscFonts.Vbb as Array<{
        xa?: string;
        IH?: (...args: unknown[]) => unknown;
      }>;
      for (const entry of catalog) {
        if (!entry?.IH) continue;
        const original = entry.IH;
        entry.IH = function (...args: unknown[]) {
          trace.push(entry.xa ?? "unknown");
          return original.apply(this, args);
        };
      }
      (window as unknown as { __fontRenderTrace?: string[] }).__fontRenderTrace = trace;
    });
    await frame.evaluate(() => {
      (window as unknown as { __fontRenderTrace?: string[] }).__fontRenderTrace?.splice(0);
    });
    await overlay.click({ position: { x: 120, y: 100 } });
    await page.keyboard.type("东莞市人民政府");
    if (fontName) {
      await page.keyboard.press("ControlOrMeta+A");
      const fontCombo = frame.locator(".combobox.fonts");
      const fontInput = fontCombo.locator('input[role="combobox"]');
      await fontCombo.locator(".dropdown-toggle").click();
      const itemId = await frame.evaluate((name) =>
        window.DE?.getController?.("Toolbar")
          ?.getView?.("Toolbar")
          ?.cmbFontName?.store?.findWhere?.({ name })
          ?.get?.("id"),
        fontName,
      );
      if (!itemId) throw new Error(`Missing ${fontName} in the font menu`);
      await fontCombo
        .locator(`li[id="${itemId}"] .font-item`)
        .evaluate(clickHtmlElement);
      await expect(fontInput).toHaveValue(fontName);
    }
    await page.waitForTimeout(1_000);
    const canvases = await frame.evaluate(() =>
      Array.from(document.querySelectorAll("canvas"))
        .map((element, index) => {
          const canvas = element as HTMLCanvasElement;
          const context = canvas.getContext("2d");
          const data = context?.getImageData(0, 0, canvas.width, canvas.height).data;
          let hash = 2_166_136_261;
          if (data) {
            for (let offset = 0; offset < data.length; offset += 16) {
              hash = Math.imul(hash ^ data[offset], 16_777_619);
            }
          }
          return { id: canvas.id || `canvas-${index}`, hash: hash >>> 0 };
        })
        .filter(({ id }) => id.startsWith("id_viewer")),
    );
    const fontFaceCache = await frame.evaluate(() => {
      const fontApplication = window.AscCommon.Ay as {
        yX?: { fqb?: { rea?: Record<string, unknown> } };
      };
      return {
        word: Object.keys(fontApplication.yX?.fqb?.rea ?? {}).some((key) =>
          key.startsWith("1002"),
        ),
        common: Object.keys(
          (window.AscCommon.DA as { fqb?: { rea?: Record<string, unknown> } })
            ?.fqb?.rea ?? {},
        ).some((key) => key.startsWith("1002")),
      };
    });
    const fontRenderTrace = await frame.evaluate(
      () =>
        Array.from(
          new Set(
            (window as unknown as { __fontRenderTrace?: string[] })
              .__fontRenderTrace ?? [],
          ),
        ),
    );
    const fontResolution = await frame.evaluate(() => {
      const picker = window.AscCommon.Ay as {
        NRa: (codePoint: number, name: string, style: number) => unknown;
        ila: (name: string, size: number, style: number) => unknown;
      };
      const codePoint = "东".codePointAt(0)!;
      return {
        arial: picker.NRa(codePoint, "Arial", 0),
        custom: picker.NRa(codePoint, "方正小标宋简体", 0),
      };
    });
    await page.close();
    return {
      canvases,
      customFontRequests,
      fontFaceCache,
      fontRenderTrace,
      fontResolution,
    };
  };

  const fallback = await render();
  const custom = await render("方正小标宋简体");
  const canvasHash = (result: typeof custom) =>
    result.canvases.find(({ id }) => id === "id_viewer")?.hash;

  expect(custom.customFontRequests).toBeGreaterThan(0);
  expect(custom.fontFaceCache).toEqual({ word: true, common: true });
  expect(custom.fontRenderTrace).toContain("方正小标宋简体");
  expect(custom.fontResolution.custom).not.toBe(fallback.fontResolution.custom);
  expect(canvasHash(custom)).not.toBe(canvasHash(fallback));
});
