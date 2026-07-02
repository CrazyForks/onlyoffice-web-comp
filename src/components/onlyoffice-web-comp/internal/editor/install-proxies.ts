import { createFetchProxy } from "./fetch";
import { createXHRProxy } from "./xhr";
import type { EditorServer } from "./server";
import type { MockSocket, MockSocketOptions } from "./socket";

export function shouldBypassOnlyOfficeProxy(url: string, baseUrl: string) {
  const pathname = new URL(url, baseUrl).pathname;

  return (
    pathname.includes("/sdkjs/common/AllFonts.js") ||
    pathname.includes("/sdkjs/common/libfont/") ||
    pathname.includes("/fonts/")
  );
}

export type ScopedIoFactory = (
  url?: string,
  options?: MockSocketOptions,
) => MockSocket;

export type OnlyOfficeParentWindow = Window & {
  __ONLYOFFICE_SCOPED_IO__?: Record<string, ScopedIoFactory>;
};

export function getScopedIoRegistry(
  win: Window = window,
): Record<string, ScopedIoFactory> {
  const parent = win as OnlyOfficeParentWindow;
  if (!parent.__ONLYOFFICE_SCOPED_IO__) {
    parent.__ONLYOFFICE_SCOPED_IO__ = {};
  }
  return parent.__ONLYOFFICE_SCOPED_IO__;
}

export function registerScopedIo(
  containerId: string,
  factory: ScopedIoFactory,
  win: Window = window,
) {
  getScopedIoRegistry(win)[containerId] = factory;
}

export function unregisterScopedIo(containerId: string, win: Window = window) {
  const registry = (win as OnlyOfficeParentWindow).__ONLYOFFICE_SCOPED_IO__;
  if (registry) {
    delete registry[containerId];
  }
}

export type OnlyOfficeProxyWindow = Window & {
  __ONLYOFFICE_PROXIES_INSTALLED__?: boolean;
  __ONLYOFFICE_GETFILE_PATCHED__?: boolean;
  __ONLYOFFICE_PROXY_SERVER__?: EditorServer;
  XMLHttpRequest: typeof XMLHttpRequest;
  Worker: typeof Worker;
  AscCommon?: {
    getFile?: (url: string) => void;
  };
};

export type InstallOnlyOfficeProxyOptions = {
  installIo?: boolean;
};

function extractDownloadFileName(url: string) {
  if (!url) {
    return "download";
  }

  try {
    const parsed = new URL(url, window.location.href);
    const fromQuery = parsed.searchParams.get("filename");
    if (fromQuery) {
      return decodeURIComponent(fromQuery);
    }

    const pathname = parsed.pathname;
    const name = decodeURIComponent(pathname.split("/").pop() || "");
    if (name.startsWith("output.")) {
      return name;
    }
    if (name) {
      return name;
    }
  } catch {
    const fallback = url.split("/").pop() || url.split("?")[0];
    if (fallback) {
      return decodeURIComponent(fallback);
    }
  }

  return "download";
}

function parseContentDispositionFileName(header: string | null) {
  if (!header) {
    return "";
  }

  const utf8Match = /filename\*=UTF-8''([^;]+)/i.exec(header);
  if (utf8Match?.[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch {
      return utf8Match[1];
    }
  }

  const asciiMatch = /filename="([^"]+)"/i.exec(header);
  if (asciiMatch?.[1]) {
    return asciiMatch[1];
  }

  return "";
}

function scheduleNamedDownloadPatch(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
  retries = 100,
) {
  if (win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }

  if (win.AscCommon?.getFile) {
    installNamedDownloadPatch(win, server);
    return;
  }

  if (retries > 0) {
    win.setTimeout(
      () => scheduleNamedDownloadPatch(win, server, retries - 1),
      50,
    );
  }
}

function extractOutputNameFromCacheUrl(url: string) {
  const match = /\/cache\/files\/data\/[^/]+\/(output\.[^/?#]+)/i.exec(url);
  return match?.[1] ?? "";
}

function triggerBlobDownload(win: Window, blob: Blob, fileName: string) {
  const objectUrl = URL.createObjectURL(blob);
  const link = win.document.createElement("a");
  link.href = objectUrl;
  link.download = fileName;
  link.style.display = "none";
  win.document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}

/** mock cache URL 无法通过 iframe 下载携带文件名，改为 fetch + <a download>。 */
function installNamedDownloadPatch(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;
  const ascCommon = win.AscCommon;
  if (!ascCommon?.getFile || win.__ONLYOFFICE_GETFILE_PATCHED__) {
    return;
  }

  const nativeGetFile = ascCommon.getFile.bind(ascCommon);
  const fetchFile = win.fetch.bind(win);

  ascCommon.getFile = (url: string) => {
    if (typeof url !== "string" || !url) {
      nativeGetFile(url);
      return;
    }

    const needsNamedDownload =
      url.includes("/cache/files/") || url.startsWith("blob:");

    if (!needsNamedDownload) {
      nativeGetFile(url);
      return;
    }

    const fallbackName = extractDownloadFileName(url);
    void fetchFile(url)
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Download failed: ${response.status}`);
        }
        const fileName =
          parseContentDispositionFileName(
            response.headers.get("Content-Disposition"),
          ) || fallbackName;
        return response.blob().then((blob) => ({ blob, fileName }));
      })
      .then(({ blob, fileName }) => {
        triggerBlobDownload(win, blob, fileName);
      })
      .catch((err) => {
        console.warn("[OnlyOffice] named download fetch failed:", err);
        const outputName = extractOutputNameFromCacheUrl(url);
        const currentServer = win.__ONLYOFFICE_PROXY_SERVER__ ?? server;
        const blobUrl = outputName
          ? currentServer.getStoredOutputUrl(outputName)
          : null;
        const fileName =
          (outputName && currentServer.getStoredOutputFileName(outputName)) ||
          fallbackName;
        if (blobUrl) {
          void fetchFile(blobUrl)
            .then((response) => response.blob())
            .then((blob) => triggerBlobDownload(win, blob, fileName))
            .catch((fallbackErr) => {
              console.warn("[OnlyOffice] blob download fallback:", fallbackErr);
              nativeGetFile(url);
            });
          return;
        }
        nativeGetFile(url);
      });
  };

  win.__ONLYOFFICE_GETFILE_PATCHED__ = true;
}

export function installOnlyOfficeProxies(
  win: OnlyOfficeProxyWindow,
  server: EditorServer,
  createIo: ScopedIoFactory,
  options: InstallOnlyOfficeProxyOptions = {},
) {
  win.__ONLYOFFICE_PROXY_SERVER__ = server;

  if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
    scheduleNamedDownloadPatch(win, server);
    return;
  }

  const xhr = createXHRProxy(win.XMLHttpRequest, {
    baseUrl: win.location.href,
    shouldBypass: (url) => shouldBypassOnlyOfficeProxy(url, win.location.href),
  });
  const fetchProxy = createFetchProxy(win);
  const WorkerCtor = win.Worker;

  xhr.use((request) => win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null);
  fetchProxy.use(
    (request) => win.__ONLYOFFICE_PROXY_SERVER__?.handleRequest(request) ?? null,
  );

  const patches: Partial<OnlyOfficeProxyWindow> & { io?: ScopedIoFactory } = {
    XMLHttpRequest: xhr,
    fetch: fetchProxy,
    Worker: function Worker(url: string, options?: WorkerOptions) {
      const u = new URL(url, win.location.origin);
      return new WorkerCtor(
        u.href.replace(u.origin, win.location.origin),
        options,
      );
    } as unknown as typeof Worker,
  };
  if (options.installIo !== false) {
    patches.io = createIo;
  }

  Object.assign(win, patches);
  win.__ONLYOFFICE_PROXIES_INSTALLED__ = true;
  scheduleNamedDownloadPatch(win, server);
}

export const REPORTER_HTML = "index.reporter.html";

export type ReporterBridge = {
  install: (target: Window) => void;
};

export type ReporterHookWindow = Window & {
  open: typeof window.open;
  __ONLYOFFICE_REPORTER_HOOK__?: boolean;
  __ONLYOFFICE_REPORTER_BRIDGE__?: ReporterBridge;
};

export function installReporterWindowHook(
  win: ReporterHookWindow,
  installProxies: (target: Window) => void,
) {
  if (win.__ONLYOFFICE_REPORTER_HOOK__) {
    return;
  }

  win.__ONLYOFFICE_REPORTER_BRIDGE__ = { install: installProxies };
  win.__ONLYOFFICE_REPORTER_HOOK__ = true;

  const nativeOpen = win.open.bind(win);
  win.open = function openReporter(
    url?: string | URL,
    target?: string,
    features?: string,
  ) {
    const popup = nativeOpen(url, target, features);
    const href = typeof url === "string" ? url : url?.toString() ?? "";

    if (popup && href.includes(REPORTER_HTML)) {
      watchReporterWindow(popup, installProxies);
    }

    return popup;
  };
}

function watchReporterWindow(
  popup: Window,
  installProxies: (target: Window) => void,
) {
  const tryInstall = () => {
    if (popup.closed) {
      return true;
    }

    try {
      if (popup.location.href.includes(REPORTER_HTML)) {
        installProxies(popup);
        return true;
      }
    } catch {
      // Navigation in progress; keep polling.
    }

    return false;
  };

  if (tryInstall()) {
    return;
  }

  const interval = window.setInterval(() => {
    if (tryInstall()) {
      window.clearInterval(interval);
    }
  }, 1);

  popup.addEventListener(
    "load",
    () => {
      tryInstall();
      window.clearInterval(interval);
    },
    { once: true },
  );
}
