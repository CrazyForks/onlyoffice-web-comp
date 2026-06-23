import {
  STATIC_RESOURCE,
  resolveSiteUrl,
  X2T_PDF_FONT_ALIASES,
} from "../../const";

let cachedFonts: Record<string, Uint8Array> | null = null;
let loadingPromise: Promise<Record<string, Uint8Array>> | null = null;

async function fetchDefaultPdfFont(origin: string) {
  const response = await fetch(
    resolveSiteUrl(origin, STATIC_RESOURCE.x2t.pdfFonts.default),
  );
  if (!response.ok) {
    return null;
  }
  return new Uint8Array(await response.arrayBuffer());
}

/** 加载 PDF 导出默认 TTF（单文件、多别名），结果缓存在内存中。 */
export async function loadX2tPdfFonts(
  origin: string,
): Promise<Record<string, Uint8Array>> {
  if (cachedFonts) {
    return cachedFonts;
  }
  if (loadingPromise) {
    return loadingPromise;
  }

  loadingPromise = (async () => {
    const bytes = await fetchDefaultPdfFont(origin);
    if (!bytes?.byteLength) {
      console.warn(
        "[x2t-pdf-fonts] default font missing:",
        STATIC_RESOURCE.x2t.pdfFonts.default,
      );
      cachedFonts = {};
      return cachedFonts;
    }

    const fonts: Record<string, Uint8Array> = {};
    for (const name of X2T_PDF_FONT_ALIASES) {
      fonts[name] = bytes;
    }

    cachedFonts = fonts;
    return fonts;
  })();

  try {
    return await loadingPromise;
  } finally {
    loadingPromise = null;
  }
}
