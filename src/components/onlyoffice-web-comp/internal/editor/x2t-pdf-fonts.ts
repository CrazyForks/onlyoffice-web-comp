import {
  STATIC_RESOURCE,
  resolveSiteUrl,
  X2T_PDF_FONT_MANIFEST,
} from "../../const";

let cachedFonts: Record<string, Uint8Array> | null = null;
let cachedFontsRoot = "";
let loadingPromise: Promise<Record<string, Uint8Array>> | null = null;
let loadingFontsRoot = "";

async function fetchPdfFontFile(origin: string, root: string, file: string) {
  const url = resolveSiteUrl(
    origin,
    `${root}/${file}`,
  );
  const response = await fetch(url);
  if (!response.ok) {
    console.warn("[x2t-pdf-fonts] font fetch failed:", url, response.status);
    return null;
  }
  const bytes = new Uint8Array(await response.arrayBuffer());
  return bytes.byteLength ? bytes : null;
}

/** 加载 PDF 导出字体，结果缓存在内存中。 */
export async function loadX2tPdfFonts(
  origin: string,
  root = STATIC_RESOURCE.x2t.pdfFonts.root,
): Promise<Record<string, Uint8Array>> {
  if (cachedFonts && cachedFontsRoot === root) {
    return cachedFonts;
  }
  if (loadingPromise && loadingFontsRoot === root) {
    return loadingPromise;
  }
  loadingFontsRoot = root;

  loadingPromise = (async () => {
    const fonts: Record<string, Uint8Array> = {};

    for (const entry of X2T_PDF_FONT_MANIFEST) {
      const bytes = await fetchPdfFontFile(origin, root, entry.file);
      if (!bytes) {
        continue;
      }
      for (const alias of entry.aliases) {
        fonts[alias] = bytes;
      }
    }

    if (!fonts["Carlito.ttf"]?.byteLength) {
      console.warn(
        "[x2t-pdf-fonts] Carlito regular missing under",
        root,
      );
      return {};
    }

    cachedFonts = fonts;
    cachedFontsRoot = root;
    return cachedFonts;
  })();

  try {
    return await loadingPromise;
  } finally {
    loadingPromise = null;
  }
}
