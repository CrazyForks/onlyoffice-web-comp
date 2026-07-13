"use client";

/**
 * 多实例 Tab 演示：onlyOfficeManagerFactory 按 containerId 隔离，切换 Tab 时隐藏不销毁。
 * 文档与完整源码说明见 `onlyoffice-web-comp/docs/多实例示例.md`。
 */
import { memo, useEffect, useRef, useState } from "react";
import { nanoid } from "nanoid";
import {
  DemoButton,
  DemoField,
  DemoMenu,
  DemoMenuRow,
  DemoSelect,
  demoHeaderClass,
  demoHeaderInnerClass,
  demoSubtitleClass,
  demoTitleClass,
  demoToolbarClass,
} from "./demo-toolbar";
import { DocxCommentsCrud } from "./docx-comments-crud";
import { DocxRevisionsCrud } from "./docx-revisions-crud";
import {
  createConnectorDemo,
  type ConnectorDemo,
} from "./connector-demo";
import { getFileExtension } from "./office-formats";
import {
  applyDemoResourceMode,
  getDemoResourceState,
  ResourceSwitcher,
  subscribeDemoResourceChange,
} from "./resource-switcher";
import {
  DEFAULT_OFFICE_THEME,
  FILE_TYPE,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  OFFICE_XML_EVENT_CONFIG,
  OFFICE_THEME_OPTIONS,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
  type FileType,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";

type DocKind = "word" | "excel" | "ppt";

type DocPreset = {
  label: string;
  badge: string;
  fileType: FileType;
  defaultFileName: string;
  accept: string;
};

type TabItem = {
  id: string;
  label: string;
  containerId: string;
  fileName: string;
  readOnly: boolean;
  docKind: DocKind;
};

type ConnectorMessage = {
  text: string;
  tone: "success" | "error";
};

const DOC_PRESETS: Record<DocKind, DocPreset> = {
  word: {
    label: "Word",
    badge: "W",
    fileType: FILE_TYPE.DOCX,
    defaultFileName: "New_Document.docx",
    accept: ".docx,.doc,.docm,.odt,.rtf,.txt",
  },
  excel: {
    label: "Excel",
    badge: "E",
    fileType: FILE_TYPE.XLSX,
    defaultFileName: "New_Spreadsheet.xlsx",
    accept: ".xlsx,.xls,.ods,.csv",
  },
  ppt: {
    label: "PPT",
    badge: "P",
    fileType: FILE_TYPE.PPTX,
    defaultFileName: "New_Presentation.pptx",
    accept: ".pptx,.ppt,.odp",
  },
};

const OnlyOfficeTabHost = memo(function OnlyOfficeTabHost({
  containerId,
}: {
  containerId: string;
}) {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} absolute inset-0`}
      data-onlyoffice-container-id={containerId}
    >
      <div id={containerId} className="absolute inset-0" />
    </div>
  );
});

function getPreset(docKind: DocKind) {
  return DOC_PRESETS[docKind];
}

function isNewDocument(tab: TabItem) {
  return tab.fileName === getPreset(tab.docKind).defaultFileName;
}

function createTab(index: number, docKind: DocKind): TabItem {
  const id = nanoid(6);
  const preset = getPreset(docKind);
  return {
    id,
    label: `${preset.label} ${index}`,
    containerId: `tab-editor-${id}`,
    fileName: preset.defaultFileName,
    readOnly: false,
    docKind,
  };
}

function createInitialTabState() {
  const initialTab = createTab(1, "word");
  return { tabs: [initialTab], activeId: initialTab.id };
}

export function TabsMultiPage({ embedded = false }: { embedded?: boolean }) {
  const [tabs, setTabs] = useState<TabItem[]>([]);
  const [activeId, setActiveId] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [theme, setTheme] = useState<OfficeTheme>(DEFAULT_OFFICE_THEME);
  const [officeXmlEventEnabled, setOfficeXmlEventEnabled] = useState<boolean>(
    OFFICE_XML_EVENT_CONFIG.default.isEnable,
  );
  const [officeXmlLimitMb, setOfficeXmlLimitMb] = useState(
    Math.round(OFFICE_XML_EVENT_CONFIG.default.limitBytes / 1024 / 1024),
  );
  const [cdnOrigin, setCdnOrigin] = useState(
    () => getDemoResourceState().cdnOrigin,
  );
  const [resourceRevision, setResourceRevision] = useState(0);
  const initializedRef = useRef(new Set<string>());
  const connectorsRef = useRef(new Map<string, ConnectorDemo>());
  const fileInputRef = useRef<HTMLInputElement>(null);
  const connectorMessageTimerRef = useRef<number | null>(null);
  const [connectorMessage, setConnectorMessage] =
    useState<ConnectorMessage | null>(null);

  const removeTabConnector = (tabId: string) => {
    const connector = connectorsRef.current.get(tabId);
    if (connector?.isConnected) {
      connector.disconnect();
    }
    connectorsRef.current.delete(tabId);
  };

  const replaceTabConnector = (
    tabId: string,
    manager: Awaited<ReturnType<typeof onlyOfficeManagerFactory.open>>,
  ) => {
    removeTabConnector(tabId);
    const connector = createConnectorDemo(() => manager);
    connectorsRef.current.set(tabId, connector);
    return connector;
  };

  const showConnectorMessage = (text: string, tone: ConnectorMessage["tone"]) => {
    if (connectorMessageTimerRef.current !== null) {
      window.clearTimeout(connectorMessageTimerRef.current);
    }
    setConnectorMessage({ text, tone });
    connectorMessageTimerRef.current = window.setTimeout(() => {
      setConnectorMessage(null);
      connectorMessageTimerRef.current = null;
    }, 4_000);
  };

  useEffect(
    () =>
      subscribeDemoResourceChange((state) => {
        setCdnOrigin(state.cdnOrigin);
        setLoading(true);
        connectorsRef.current.forEach((connector) => connector.disconnect());
        connectorsRef.current.clear();
        onlyOfficeManagerFactory.destroyAll();
        initializedRef.current.clear();
        setResourceRevision(state.revision);
      }),
    [],
  );

  useEffect(() => {
    const { tabs: initialTabs, activeId: initialActiveId } = createInitialTabState();
    setTabs(initialTabs);
    setActiveId(initialActiveId);
  }, []);

  const activeTab = tabs.find((tab) => tab.id === activeId);
  const activePreset = activeTab ? getPreset(activeTab.docKind) : null;

  const updateTab = (tabId: string, patch: Partial<TabItem>) => {
    setTabs((prev) =>
      prev.map((tab) => (tab.id === tabId ? { ...tab, ...patch } : tab)),
    );
  };

  const getOfficeXmlEventConfig = () => ({
    isEnable: officeXmlEventEnabled,
    limitBytes: officeXmlLimitMb * 1024 * 1024,
  });

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const openTabEditor = async (tab: TabItem) => {
    const preset = getPreset(tab.docKind);

    const manager = await onlyOfficeManagerFactory.open(
      {
        containerId: tab.containerId,
        fileType: preset.fileType,
        defaultFileName: preset.defaultFileName,
        readOnly: tab.readOnly,
        theme,
        officeXmlEvent: getOfficeXmlEventConfig(),
      },
      {
        fileName: tab.fileName,
        isNew: isNewDocument(tab),
        readOnly: tab.readOnly,
      },
    );

    replaceTabConnector(tab.id, manager);
    initializedRef.current.add(tab.id);
  };

  useEffect(() => {
    if (!activeId) return;

    const tab = tabs.find((item) => item.id === activeId);
    if (!tab || initializedRef.current.has(tab.id)) return;

    let cancelled = false;

    openTabEditor(tab)
      .then(() => {
        if (cancelled) {
          onlyOfficeManagerFactory.destroy(tab.containerId);
          initializedRef.current.delete(tab.id);
          return;
        }
        setLoading(false);
      })
      .catch((err) => {
        if (cancelled) return;
        setError("无法加载编辑器");
        setLoading(false);
        console.error("Failed to open tab editor:", err);
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeId, tabs, resourceRevision]);

  useEffect(() => {
    const handleLoadingChange = (data: { loading: boolean }) => {
      setLoading(data.loading);
    };
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handleLoadingChange);

    return () => {
      onlyofficeEventbus.off(
        ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE,
        handleLoadingChange,
      );
      connectorsRef.current.forEach((connector) => connector.disconnect());
      connectorsRef.current.clear();
      onlyOfficeManagerFactory.destroyAll();
      initializedRef.current.clear();
      if (connectorMessageTimerRef.current !== null) {
        window.clearTimeout(connectorMessageTimerRef.current);
      }
    };
  }, []);

  const addTab = (docKind: DocKind) => {
    const nextTab = createTab(tabs.length + 1, docKind);
    setTabs((prev) => [...prev, nextTab]);
    setActiveId(nextTab.id);
  };

  const closeTab = (tabId: string) => {
    if (tabs.length <= 1) return;

    const tab = tabs.find((item) => item.id === tabId);
    if (tab) {
      removeTabConnector(tab.id);
      onlyOfficeManagerFactory.destroy(tab.containerId);
      initializedRef.current.delete(tab.id);
    }

    setTabs((prev) => {
      const next = prev.filter((item) => item.id !== tabId);
      if (activeId === tabId) {
        setActiveId(next[0]?.id ?? "");
      }
      return next;
    });
  };

  const ensureActiveManager = async () => {
    if (!activeTab) throw new Error("No active tab");

    if (!initializedRef.current.has(activeTab.id)) {
      await openTabEditor(activeTab);
    }

    const manager = onlyOfficeManagerFactory.get(activeTab.containerId);
    if (!manager) throw new Error("Editor is not initialized");
    return manager;
  };

  const uploadFile = (file: File) =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      const manager = await onlyOfficeManagerFactory.open(
        {
          containerId: activeTab.containerId,
          fileType: preset.fileType,
          defaultFileName: preset.defaultFileName,
          readOnly: activeTab.readOnly,
          theme,
          officeXmlEvent: getOfficeXmlEventConfig(),
        },
        {
          fileName: file.name,
          file,
          readOnly: activeTab.readOnly,
        },
      );

      replaceTabConnector(activeTab.id, manager);
      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: file.name });
    }, "上传失败");

  const newDocument = () =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);

      const manager = await onlyOfficeManagerFactory.open(
        {
          containerId: activeTab.containerId,
          fileType: preset.fileType,
          defaultFileName: preset.defaultFileName,
          readOnly: activeTab.readOnly,
          theme,
          officeXmlEvent: getOfficeXmlEventConfig(),
        },
        {
          fileName: preset.defaultFileName,
          isNew: true,
          readOnly: activeTab.readOnly,
        },
      );

      replaceTabConnector(activeTab.id, manager);
      initializedRef.current.add(activeTab.id);
      updateTab(activeTab.id, { fileName: preset.defaultFileName });
    }, "新建失败");

  const exportDocument = () =>
    runAction(async () => {
      const manager = await ensureActiveManager();
      await manager.downloadExport();
    }, "导出失败");

  const printActiveLogs = () =>
    runAction(async () => {
      const manager = await ensureActiveManager();
      manager.printLogs();
    }, "打印日志失败");

  const writeWithConnector = () =>
    runAction(async () => {
      if (!activeTab) throw new Error("No active tab");

      const manager = await ensureActiveManager();
      const logger = manager.getLogger();
      let connector = connectorsRef.current.get(activeTab.id);
      if (!connector) {
        connector = replaceTabConnector(activeTab.id, manager);
      }

      try {
        logger.operation("Connector command started", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: getPreset(activeTab.docKind).fileType,
        });
        const result = await connector.write(
          activeTab.fileName,
          getPreset(activeTab.docKind).fileType,
        );
        logger.operation("Connector command completed", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: result.fileType,
        });
        showConnectorMessage(result.message, "success");
      } catch (error) {
        logger.error("operation", "Connector command failed", {
          tabId: activeTab.id,
          fileName: activeTab.fileName,
          fileType: getPreset(activeTab.docKind).fileType,
          error,
        });
        showConnectorMessage(
          error instanceof Error ? error.message : "Connector command failed",
          "error",
        );
        throw error;
      }
    }, "连接器调用失败");

  const toggleReadOnly = () =>
    runAction(async () => {
      if (!activeTab) return;

      const preset = getPreset(activeTab.docKind);
      const manager = await ensureActiveManager();
      const nextReadOnly = !activeTab.readOnly;

      if (manager.isReady()) {
        await manager.setReadOnly(nextReadOnly);
      } else {
        await onlyOfficeManagerFactory.open(
          {
            containerId: activeTab.containerId,
            fileType: preset.fileType,
            defaultFileName: preset.defaultFileName,
            readOnly: nextReadOnly,
            theme,
            officeXmlEvent: getOfficeXmlEventConfig(),
          },
          {
            fileName: activeTab.fileName,
            isNew: isNewDocument(activeTab),
            readOnly: nextReadOnly,
          },
        );
      }

      updateTab(activeTab.id, { readOnly: nextReadOnly });
    }, "切换模式失败");

  const applyTheme = (nextTheme: OfficeTheme) =>
    runAction(async () => {
      setTheme(nextTheme);

      await Promise.all(
        tabs.map(async (tab) => {
          if (!initializedRef.current.has(tab.id)) return;

          const manager = onlyOfficeManagerFactory.get(tab.containerId);
          if (manager?.isReady()) {
            await manager.setTheme(nextTheme);
          }
        }),
      );
    }, "切换主题失败");

  const loadResource = () =>
    runAction(async () => {
      applyDemoResourceMode("cdn", cdnOrigin);
    }, "切换资源失败");

  const isActiveDocx = activeTab
    ? getFileExtension(activeTab.fileName, activePreset?.fileType) === "docx"
    : false;

  return (
    <div
      className={`flex flex-col bg-neutral-100 ${
        embedded ? "h-full min-h-0" : "h-screen"
      }`}
    >
      <header className={demoHeaderClass}>
        <div className={demoHeaderInnerClass}>
          <div className="mr-auto min-w-0">
            <h1 className={demoTitleClass}>{embedded ? "示例" : "多实例"}</h1>
            <p className={demoSubtitleClass}>
              {activeTab
                ? `${activePreset?.label} · ${activeTab.fileName}`
                : "切换标签页，实例状态会保留"}
            </p>
          </div>

          <div className={demoToolbarClass}>
            <DemoButton onClick={() => fileInputRef.current?.click()}>
              上传
            </DemoButton>
            <DemoButton onClick={newDocument}>
              新建{activePreset?.label ?? "文档"}
            </DemoButton>
            <DemoButton onClick={exportDocument}>导出</DemoButton>
            <DemoButton onClick={printActiveLogs}>打印日志</DemoButton>
            <DemoButton onClick={writeWithConnector}>连接器写入</DemoButton>
            <DemoButton active={!!activeTab?.readOnly} onClick={toggleReadOnly}>
              {activeTab?.readOnly ? "只读" : "编辑"}
            </DemoButton>
            <DemoMenu label="更多" disabled={loading}>
              <DemoMenuRow>
                <ResourceSwitcher
                  cdnOrigin={cdnOrigin}
                  disabled={loading}
                  onCdnOriginChange={setCdnOrigin}
                  onLoad={loadResource}
                />
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoField label="主题">
                  <DemoSelect
                    value={theme}
                    onChange={(event) =>
                      applyTheme(event.target.value as OfficeTheme)
                    }
                  >
                    {OFFICE_THEME_OPTIONS.map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </DemoSelect>
                </DemoField>
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoField label="XML 检测">
                  <input
                    type="checkbox"
                    checked={officeXmlEventEnabled}
                    onChange={(event) =>
                      setOfficeXmlEventEnabled(event.target.checked)
                    }
                    className="h-4 w-4"
                  />
                </DemoField>
                <DemoField label="阈值 MB">
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={officeXmlLimitMb}
                    onChange={(event) =>
                      setOfficeXmlLimitMb(
                        Math.max(1, Number(event.target.value) || 1),
                      )
                    }
                    className="h-6 w-20 border-0 bg-transparent py-0 pl-0.5 text-[13px] text-neutral-800 outline-none"
                  />
                </DemoField>
              </DemoMenuRow>
              {isActiveDocx && (
                <>
                  <DocxCommentsCrud
                    disabled={loading || !!activeTab?.readOnly}
                    getManager={ensureActiveManager}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                  <DocxRevisionsCrud
                    disabled={loading || !!activeTab?.readOnly}
                    getManager={ensureActiveManager}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                </>
              )}
            </DemoMenu>
          </div>
        </div>

        <div className="border-t border-neutral-200/80 bg-[#f5f4f3] px-2 py-1">
          <div className="flex items-end gap-0.5 overflow-x-auto">
            {tabs.map((tab) => {
              const preset = getPreset(tab.docKind);
              const isActive = activeId === tab.id;
              return (
                <div
                  key={tab.id}
                  className={`group relative flex max-w-[200px] min-w-[88px] shrink-0 items-stretch border border-b-0 ${
                    isActive
                      ? "z-10 -mb-px border-neutral-300 bg-white"
                      : "border-transparent hover:bg-white/50"
                  }`}
                >
                  <button
                    type="button"
                    onClick={() => setActiveId(tab.id)}
                    className={`flex min-w-0 flex-1 items-center gap-1.5 px-2.5 py-1.5 text-[12px] ${
                      isActive ? "text-neutral-900" : "text-neutral-500"
                    }`}
                    title={tab.fileName}
                  >
                    <span className="shrink-0 text-[11px] text-neutral-400">
                      {preset.badge}
                    </span>
                    <span className="truncate">{tab.label}</span>
                  </button>
                  {tabs.length > 1 && (
                    <button
                      type="button"
                      onClick={() => closeTab(tab.id)}
                      className={`mr-1 self-center rounded px-1 text-[10px] transition-colors ${
                        isActive
                          ? "text-gray-400 hover:text-gray-600"
                          : "text-gray-300 opacity-0 hover:text-gray-500 group-hover:opacity-100"
                      }`}
                      aria-label={`关闭 ${tab.label}`}
                    >
                      ×
                    </button>
                  )}
                </div>
              );
            })}

            <div className="mb-px flex shrink-0 items-center gap-1 pl-1.5">
              {(Object.keys(DOC_PRESETS) as DocKind[]).map((kind) => {
                const preset = getPreset(kind);
                return (
                  <button
                    key={kind}
                    type="button"
                    onClick={() => addTab(kind)}
                    className="inline-flex h-7 items-center border border-dashed border-neutral-300 bg-transparent px-2 text-[12px] text-neutral-600 hover:border-neutral-400 hover:bg-white"
                    title={`新建 ${preset.label} 标签页`}
                  >
                    + {preset.label}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </header>

      {connectorMessage && (
        <output
          aria-live="polite"
          className={`fixed top-4 left-1/2 z-50 -translate-x-1/2 rounded border px-4 py-2 text-sm shadow-lg ${
            connectorMessage.tone === "success"
              ? "border-emerald-200 bg-emerald-50 text-emerald-800"
              : "border-red-200 bg-red-50 text-red-800"
          }`}
          role="status"
        >
          {connectorMessage.text}
        </output>
      )}

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-3 text-sm text-red-700">
          {error}
        </div>
      )}

      <div className="relative min-h-0 flex-1 bg-white">
        {tabs.length === 0 && (
          <div className="absolute inset-0 flex items-center justify-center text-sm text-gray-400">
            加载中...
          </div>
        )}
        {tabs.map((tab) => (
          <div
            key={tab.id}
            className={`absolute inset-0 ${
              activeId === tab.id ? "visible z-10" : "invisible z-0"
            }`}
          >
            <OnlyOfficeTabHost containerId={tab.containerId} />
          </div>
        ))}

        {loading && (
          <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
            <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
              加载中...
            </div>
          </div>
        )}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={activePreset?.accept ?? ".docx,.doc,.odt,.rtf,.txt"}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            uploadFile(file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
