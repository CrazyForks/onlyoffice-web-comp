import { expect, test } from "playwright/test";

const assetOrigin = `http://${process.env.PLAYWRIGHT_HOST ?? "127.0.0.1"}:${
  process.env.PLAYWRIGHT_CDN_PORT ?? 3010
}`;
const deRoot = `${assetOrigin}/onlyoffice/9.4.0-develop`;

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
  await customFontItem.evaluate((element) => element.click());
  await expect(fontInput).toHaveValue("方正小标宋简体");

  const documentCanvas = editorFrame.locator("#id_viewer");
  const documentOverlay = editorFrame.locator("#id_viewer_overlay");
  await expect(documentCanvas).toBeVisible();
  await expect(documentOverlay).toBeVisible();
  const before = await documentCanvas.evaluate((canvas) => {
    const context = (canvas as HTMLCanvasElement).getContext("2d");
    const data = context?.getImageData(0, 0, canvas.width, canvas.height).data;
    if (!data) return 0;
    let checksum = 0;
    for (let index = 0; index < data.length; index += 16) {
      checksum = (checksum + data[index] + data[index + 1] + data[index + 2]) >>> 0;
    }
    return checksum;
  });

  await documentOverlay.click({ position: { x: 120, y: 100 } });
  await page.keyboard.type("方正小标宋简体");
  await expect
    .poll(() =>
      documentCanvas.evaluate((canvas) => {
        const context = (canvas as HTMLCanvasElement).getContext("2d");
        const data = context?.getImageData(0, 0, canvas.width, canvas.height).data;
        if (!data) return 0;
        let checksum = 0;
        for (let index = 0; index < data.length; index += 16) {
          checksum = (checksum + data[index] + data[index + 1] + data[index + 2]) >>> 0;
        }
        return checksum;
      }),
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
        .evaluate((element) => element.click());
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
