"use client";

/**
 * 单实例演示页：OnlyOfficeManager 门面 + 工具栏（上传/导出/主题/语言/只读）。
 * 文档说明见 `onlyoffice-web-comp/docs/08-单实例示例.md`。
 */
import { memo, useEffect, useRef, useState } from "react";
import {
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_ID,
  ONLYOFFICE_LANG_KEY,
  OFFICE_THEME_OPTIONS,
  DEFAULT_OFFICE_THEME,
  OnlyOfficeManager,
  editorManagerFactory,
  type FileType,
  type OfficeTheme,
} from "@/components/onlyoffice-web-comp";

import {
  DemoButton,
  DemoField,
  DemoMenu,
  DemoMenuRow,
  DemoSelect,
  demoHeaderClass,
  demoHeaderInnerClass,
  demoTitleClass,
  demoToolbarClass,
} from "./demo-toolbar";
import { DocxCommentsCrud } from "./docx-comments-crud";
import { DocxRevisionsCrud } from "./docx-revisions-crud";
import { getFileExtension, OFFICE_UPLOAD_ACCEPT } from "./office-formats";
import {
  applyDemoResourceMode,
  getDemoResourceState,
  ResourceSwitcher,
  subscribeDemoResourceChange,
} from "./resource-switcher";

type OfficePreviewPageProps = {
  title: string;
  defaultFileName: string;
  fileType: FileType;
  accept?: string;
  newButtonLabel: string;
  /** public 目录下的默认文件路径，如 /test.xlsx */
  initialFileUrl?: string;
  /** 嵌入文档页或父容器时使用 h-full */
  embedded?: boolean;
};

function LoadingOverlay() {
  return (
    <div className="absolute inset-0 z-20 flex items-center justify-center bg-white/70 backdrop-blur-sm">
      <div className="rounded-lg border border-gray-200 bg-white px-4 py-3 text-sm text-gray-700 shadow">
        加载中...
      </div>
    </div>
  );
}

const OnlyOfficeHost = memo(function OnlyOfficeHost() {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} absolute inset-0`}
    >
      <div id={ONLYOFFICE_ID} className="absolute inset-0" />
    </div>
  );
});

async function fetchPublicFile(url: string, fileName: string) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}`);
  }
  const blob = await response.blob();
  return new File([blob], fileName, { type: blob.type });
}

export function OfficePreviewPage({
  title,
  defaultFileName,
  fileType,
  accept = OFFICE_UPLOAD_ACCEPT,
  newButtonLabel,
  initialFileUrl,
  embedded = false,
}: OfficePreviewPageProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const managerRef = useRef<OnlyOfficeManager | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [readOnly, setReadOnly] = useState(false);
  const [currentLang, setCurrentLangState] = useState(
    ONLYOFFICE_LANG_KEY.ZH as string,
  );
  const [currentTheme, setCurrentThemeState] = useState<OfficeTheme>(
    DEFAULT_OFFICE_THEME,
  );
  const [editorReady, setEditorReady] = useState(false);
  const [activeFileName, setActiveFileName] = useState(defaultFileName);
  const [cdnOrigin, setCdnOrigin] = useState(
    () => getDemoResourceState().cdnOrigin,
  );
  const [resourceRevision, setResourceRevision] = useState(0);

  useEffect(
    () =>
      subscribeDemoResourceChange((state) => {
        setCdnOrigin(state.cdnOrigin);
        setLoading(true);
        managerRef.current?.destroy();
        managerRef.current = null;
        editorManagerFactory.destroy(ONLYOFFICE_ID);
        setActiveFileName(defaultFileName);
        setResourceRevision(state.revision);
      }),
    [defaultFileName],
  );

  useEffect(() => {
    let unsubscribeLoading: (() => void) | undefined;
    let disposed = false;
    let ownedManager: OnlyOfficeManager | null = null;
    const containerId = ONLYOFFICE_ID;

    const init = async () => {
      editorManagerFactory.destroy(containerId);
      const loadSession = editorManagerFactory.beginLoadSession(containerId);
      setEditorReady(false);

      let manager: OnlyOfficeManager;

      if (initialFileUrl) {
        const file = await fetchPublicFile(initialFileUrl, defaultFileName);
        if (disposed) return;

        manager = await OnlyOfficeManager.createWithFile(
          {
            containerId,
            fileType,
            defaultFileName,
            readOnly,
            theme: currentTheme,
            loadSession,
          },
          file,
        );
      } else {
        manager = await OnlyOfficeManager.create({
          containerId,
          fileType,
          defaultFileName,
          readOnly,
          theme: currentTheme,
          loadSession,
          user: {
            id: "uid",
            name: "demo-user",
          },
        });
      }

      if (
        disposed ||
        !editorManagerFactory.isLoadSessionActive(containerId, loadSession)
      ) {
        return;
      }

      ownedManager = manager;
      managerRef.current = manager;
      setCurrentLangState(manager.getLanguage());
      setCurrentThemeState(manager.getTheme());
      setEditorReady(true);
      setLoading(false);
      unsubscribeLoading = manager.onLoadingChange(({ loading: next }) => {
        setLoading(next);
      });
    };

    init().catch((err) => {
      if (disposed) return;
      setError("无法加载编辑器组件");
      setLoading(false);
      console.error("Failed to initialize OnlyOffice:", err);
    });

    return () => {
      disposed = true;
      unsubscribeLoading?.();
      ownedManager?.destroy();
      editorManagerFactory.destroy(containerId);
      managerRef.current = null;
      setEditorReady(false);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resourceRevision]);

  const runAction = async (action: () => Promise<void>, message: string) => {
    try {
      setError(null);
      await action();
    } catch (err) {
      setError(message);
      console.error(message, err);
    }
  };

  const handleOpenDocument = (
    fileName: string,
    file?: File,
    nextReadOnly = readOnly,
  ) =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.openDocument({ fileName, file, readOnly: nextReadOnly });
      setActiveFileName(fileName);
      setReadOnly(nextReadOnly);
    }, "操作失败");

  const handleLanguageSwitch = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      const nextLang = await manager.toggleLanguage();
      setCurrentLangState(nextLang);
    }, "切换语言失败");

  const handleThemeChange = (theme: OfficeTheme) =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.setTheme(theme);
      setCurrentThemeState(manager.getTheme());
    }, "切换主题失败");

  const handleExport = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }

      await manager.downloadExport();
    }, "导出失败");

  const handleToggleReadOnly = () =>
    runAction(async () => {
      const manager = managerRef.current;
      if (!manager) {
        throw new Error("Editor is not initialized");
      }
      await manager.toggleReadOnly();
      setReadOnly(manager.getReadOnly());
    }, "切换模式失败");

  const handleResourceLoad = () =>
    runAction(async () => {
      applyDemoResourceMode("cdn", cdnOrigin);
    }, "切换资源失败");

  const isDocxFile = getFileExtension(activeFileName, fileType) === "docx";

  return (
    <div
      className={`flex flex-col bg-white ${
        embedded ? "h-full min-h-0" : "h-screen"
      }`}
    >
      <header className={demoHeaderClass}>
        <div className={demoHeaderInnerClass}>
          <div className="mr-auto flex min-w-0 items-baseline gap-2.5">
            <h1 className={demoTitleClass}>{title}</h1>
          </div>

          <div className={demoToolbarClass}>
            <DemoButton onClick={() => fileInputRef.current?.click()}>
              上传
            </DemoButton>
            <DemoButton onClick={() => handleOpenDocument(defaultFileName)}>
              {newButtonLabel}
            </DemoButton>
            {editorReady && (
              <>
                <DemoButton onClick={handleExport}>导出</DemoButton>
                <DemoButton active={readOnly} onClick={handleToggleReadOnly}>
                  {readOnly ? "只读" : "编辑"}
                </DemoButton>
              </>
            )}
            <DemoMenu label="更多" disabled={loading}>
              <DemoMenuRow>
                <ResourceSwitcher
                  cdnOrigin={cdnOrigin}
                  disabled={loading}
                  onCdnOriginChange={setCdnOrigin}
                  onLoad={handleResourceLoad}
                />
              </DemoMenuRow>
              <DemoMenuRow>
                <DemoButton onClick={handleLanguageSwitch}>
                  {currentLang === ONLYOFFICE_LANG_KEY.ZH ? "中文" : "English"}
                </DemoButton>
                <DemoField label="主题">
                  <DemoSelect
                    value={currentTheme}
                    onChange={(event) =>
                      handleThemeChange(event.target.value as OfficeTheme)
                    }
                    disabled={!editorReady}
                  >
                    {OFFICE_THEME_OPTIONS.map((option) => (
                      <option key={option.id} value={option.id}>
                        {option.label}
                      </option>
                    ))}
                  </DemoSelect>
                </DemoField>
              </DemoMenuRow>
              {isDocxFile && (
                <>
                  <DocxCommentsCrud
                    disabled={!editorReady || loading || readOnly}
                    getManager={() => managerRef.current}
                    onError={(message, err) => {
                      setError(message);
                      console.error(message, err);
                    }}
                  />
                  <DocxRevisionsCrud
                    disabled={!editorReady || loading || readOnly}
                    getManager={() => managerRef.current}
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
      </header>

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-4 text-red-700">
          <p className="font-medium">错误：{error}</p>
        </div>
      )}

      <div className="relative min-h-0 flex-1">
        <OnlyOfficeHost />
        {loading && <LoadingOverlay />}
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            handleOpenDocument(file.name, file);
            if (fileInputRef.current) fileInputRef.current.value = "";
          }
        }}
      />
    </div>
  );
}
