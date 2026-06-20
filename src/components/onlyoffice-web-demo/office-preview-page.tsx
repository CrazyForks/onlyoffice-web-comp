"use client";

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
  DemoSelect,
  demoHeaderClass,
  demoHeaderInnerClass,
  demoTitleClass,
  demoToolbarClass,
} from "./demo-toolbar";

type OfficePreviewPageProps = {
  title: string;
  defaultFileName: string;
  fileType: FileType;
  accept: string;
  newButtonLabel: string;
  /** public 目录下的默认文件路径，如 /test.xlsx */
  initialFileUrl?: string;
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
  accept,
  newButtonLabel,
  initialFileUrl,
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

  useEffect(() => {
    let unsubscribeLoading: (() => void) | undefined;
    let disposed = false;
    let ownedManager: OnlyOfficeManager | null = null;
    const containerId = ONLYOFFICE_ID;

    const init = async () => {
      editorManagerFactory.destroy(containerId);
      const loadSession = editorManagerFactory.beginLoadSession(containerId);

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
      unsubscribeLoading = manager.onLoadingChange(({ loading: next }) => {
        setLoading(next);
      });
    };

    init().catch((err) => {
      if (disposed) return;
      setError("无法加载编辑器组件");
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
  }, []);

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
      await managerRef.current?.downloadExport();
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

  return (
    <div className="flex h-screen flex-col bg-white">
      <header className={demoHeaderClass}>
        <div className={demoHeaderInnerClass}>
          <div className="mr-auto flex min-w-0 items-baseline gap-2.5">
            <h1 className={demoTitleClass}>{title}</h1>
          </div>

          <div className={demoToolbarClass}>
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
          </div>
        </div>
      </header>

      {error && (
        <div className="mx-4 mt-4 rounded border-l-4 border-red-500 bg-red-50 p-4 text-red-700">
          <p className="font-medium">错误：{error}</p>
        </div>
      )}

      <div className="relative flex-1">
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
