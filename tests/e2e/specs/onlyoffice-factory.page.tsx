"use client";

import { useEffect, useMemo, useState } from "react";
import {
  DEFAULT_OFFICE_THEME,
  FILE_TYPE,
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  OFFICE_THEME,
  OnlyOfficeManager,
  editorManagerFactory,
  isOnlyOfficeCdnMode,
  onlyOfficeManagerFactory,
  onlyofficeEventbus,
} from "@/components/onlyoffice-web-comp";
import { converter } from "@/components/onlyoffice-web-comp/internal/editor/x2t";
import { getX2tConvertFormats } from "@/components/onlyoffice-web-comp/internal/editor/utils";
import type {
  FileType,
  OfficeTheme,
} from "@/components/onlyoffice-web-comp";
import type {
  ResourceMode,
  ScenarioResult,
  StepResult,
} from "./onlyoffice-factory.contract";

export const CONTAINER_IDS = {
  factory: "e2e-factory-editor",
  create: "e2e-create-editor",
  file: "e2e-file-editor",
  textFallback: "e2e-text-fallback-editor",
  fromEditor: "e2e-from-editor",
  fixture: "e2e-fixture-editor",
} as const;

const DOCUMENT_READY_TIMEOUT_MS = 30_000;

function waitForDocumentReady() {
  let timeoutId: number | undefined;
  let handler: (() => void) | undefined;

  const promise = new Promise<void>((resolve, reject) => {
    handler = () => {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
      if (handler) {
        onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      }
      resolve();
    };

    timeoutId = window.setTimeout(() => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      reject(new Error("Timed out waiting for OnlyOffice documentReady"));
    }, DOCUMENT_READY_TIMEOUT_MS);

    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
  });

  return {
    promise,
    cancel() {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
      if (handler) {
        onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, handler);
      }
    },
  };
}

async function withDocumentReady<T>(action: () => Promise<T>) {
  const ready = waitForDocumentReady();
  try {
    const value = await action();
    await ready.promise;
    return value;
  } catch (error) {
    ready.cancel();
    throw error;
  }
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}

async function fetchPublicFile(url: string, fileName: string) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}`);
  }
  const blob = await response.blob();
  return new File([blob], fileName, { type: blob.type });
}

function frameOrigin(containerId: string) {
  const frames = Array.from(
    document.querySelectorAll<HTMLIFrameElement>('iframe[name="frameEditor"]'),
  );
  const frame =
    frames.find((item) => {
      try {
        return new URL(item.src, window.location.href).searchParams.get(
          "frameEditorId",
        ) === containerId;
      } catch {
        return false;
      }
    }) ?? frames[0];

  assert(frame?.src, `Missing frameEditor iframe for ${containerId}`);
  return new URL(frame.src, window.location.href).origin;
}

async function assertExport(
  manager: OnlyOfficeManager,
  expectedFileType: FileType,
) {
  const data = await manager.exportDocument();
  assert(
    data.fileType === expectedFileType.toLowerCase(),
    "Unexpected export type",
  );
  assert(data.binData.byteLength > 0, "Expected exported bin data");
}

async function assertX2tImport(fileName: string, fileType: FileType) {
  const file = await fetchPublicFile(`/e2e/fixtures/${fileName}`, fileName);
  const data = await file.arrayBuffer();
  const { formatFrom, formatTo } = getX2tConvertFormats(fileType);
  const result = await converter.convert({
    data,
    fileFrom: `doc.${fileType.toLowerCase()}`,
    fileTo: "Editor.bin",
    formatFrom,
    formatTo,
  });

  assert(
    result.output && result.output.byteLength > 0,
    `Expected x2t to import ${fileName}`,
  );
}

export function resetAll() {
  onlyOfficeManagerFactory.destroyAll();
  editorManagerFactory.destroyAll();
  converter.terminate();
}

export async function runScenario(mode: ResourceMode, cdnOrigin: string) {
  const steps: StepResult[] = [];

  const runStep = async (name: string, action: () => Promise<string | void>) => {
    try {
      const detail = await action();
      steps.push({ name, status: "passed", detail: detail || undefined });
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      steps.push({ name, status: "failed", detail });
      throw new Error(`${name}: ${detail}`);
    }
  };

  resetAll();

  await runStep("resource mode", async () => {
    if (mode === "cdn") {
      OnlyOfficeManager.registerStaticResource({ cdnOrigin });
    } else {
      OnlyOfficeManager.resetStaticResource();
    }

    assert(isOnlyOfficeCdnMode() === (mode === "cdn"), "CDN mode mismatch");
    return mode === "cdn" ? cdnOrigin : "local packages";
  });

  await runStep("manager factory open/get", async () => {
    const manager = await withDocumentReady(() =>
      onlyOfficeManagerFactory.open(
        {
          containerId: CONTAINER_IDS.factory,
          fileType: FILE_TYPE.DOCX,
          defaultFileName: "Factory.docx",
          readOnly: false,
          theme: DEFAULT_OFFICE_THEME,
          user: { id: "factory-user", name: "Factory User" },
        },
        {
          fileName: "Factory.docx",
          isNew: true,
        },
      ),
    );

    assert(manager.isReady(), "Factory manager was not ready");
    assert(
      onlyOfficeManagerFactory.get(CONTAINER_IDS.factory) === manager,
      "Factory get did not return the opened manager",
    );

    const origin = frameOrigin(CONTAINER_IDS.factory);
    if (mode === "cdn") {
      assert(origin === new URL(cdnOrigin).origin, "Expected CDN iframe origin");
    } else {
      assert(origin === window.location.origin, "Expected local iframe origin");
    }
  });

  await runStep("manager facade from factory", async () => {
    const manager = onlyOfficeManagerFactory.get(CONTAINER_IDS.factory);
    assert(manager, "Factory manager is missing");

    manager.setUser({ id: "updated-user", name: "Updated User" });
    assert(manager.getUser().id === "updated-user", "setUser/getUser failed");

    await withDocumentReady(() => manager.setLanguage("en"));
    assert(manager.getLanguage() === "en", "setLanguage failed");

    await withDocumentReady(() => manager.setTheme(OFFICE_THEME.DARK));
    assert(
      manager.getTheme() === (OFFICE_THEME.DARK as OfficeTheme),
      "setTheme failed",
    );

    await manager.setReadOnly(true);
    assert(manager.getReadOnly(), "setReadOnly(true) failed");
    await manager.setReadOnly(false);
    assert(!manager.getReadOnly(), "setReadOnly(false) failed");

    await withDocumentReady(() => manager.openNew("Factory-Reopened.docx"));
    await assertExport(manager, FILE_TYPE.DOCX);
  });

  await runStep("manager factory destroy", async () => {
    onlyOfficeManagerFactory.destroy(CONTAINER_IDS.factory);
    assert(
      !onlyOfficeManagerFactory.get(CONTAINER_IDS.factory),
      "Factory destroy did not remove manager",
    );
  });

  await runStep("fixture manifest", async () => {
    const response = await fetch("/e2e/fixtures/manifest.json");
    assert(response.ok, "Fixture manifest is not available");
    const manifest = (await response.json()) as Array<{
      name: string;
      kind: "positive" | "negative";
      source: string;
    }>;
    const names = new Set(manifest.map((fixture) => fixture.name));
    assert(
      names.has("edge-invalid-bookmark.docx"),
      "Missing invalid bookmark fixture",
    );
    assert(names.has("xml-limit.docx"), "Missing XML limit fixture");
    assert(
      names.has("mismatch-xlsx-as-docx.docx"),
      "Missing mismatch fixture",
    );
    return `${manifest.length} generated fixtures`;
  });

  await runStep("x2t edge imports", async () => {
    await assertX2tImport("edge-invalid-bookmark.docx", FILE_TYPE.DOCX);
    return "DingTalk invalid bookmark DOCX converted to Editor.bin";
  });

  await runStep("generated negative fixtures", async () => {
    const file = await fetchPublicFile(
      "/e2e/fixtures/xml-limit.docx",
      "xml-limit.docx",
    );
    const manager = await OnlyOfficeManager.createWithFile(
      {
        containerId: CONTAINER_IDS.fixture,
        fileType: FILE_TYPE.DOCX,
        defaultFileName: file.name,
        readOnly: false,
        theme: DEFAULT_OFFICE_THEME,
        officeXmlEvent: {
          isEnable: true,
          limitBytes: 1024,
        },
      },
      file,
    );

    assert(
      manager.getEditor().isOfficeXmlSizeLimitExceeded(),
      "XML size guard did not block oversized Office XML",
    );
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.fixture);

    const mismatch = await fetchPublicFile(
      "/e2e/fixtures/mismatch-xlsx-as-docx.docx",
      "mismatch-xlsx-as-docx.docx",
    );
    assert(mismatch.size > 0, "Mismatch negative fixture is empty");

    return "xml guard blocked, mismatch fixture available";
  });

  await runStep("manager create", async () => {
    const manager = await withDocumentReady(() =>
      OnlyOfficeManager.create({
        containerId: CONTAINER_IDS.create,
        fileType: FILE_TYPE.PPTX,
        defaultFileName: "FactoryDeck.pptx",
        readOnly: false,
        theme: DEFAULT_OFFICE_THEME,
      }),
    );

    assert(manager.isReady(), "OnlyOfficeManager.create was not ready");
    assert(manager.getEditor().exists(), "Created editor does not exist");
    await manager.toggleReadOnly();
    assert(manager.getReadOnly(), "toggleReadOnly failed");
    await assertExport(manager, FILE_TYPE.PPTX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.create);
  });

  await runStep("manager createWithFile", async () => {
    const file = await fetchPublicFile("/test.xlsx", "test.xlsx");
    const manager = await withDocumentReady(() =>
      OnlyOfficeManager.createWithFile(
        {
          containerId: CONTAINER_IDS.file,
          fileType: FILE_TYPE.XLSX,
          defaultFileName: "test.xlsx",
          readOnly: false,
          theme: DEFAULT_OFFICE_THEME,
        },
        file,
      ),
    );

    assert(manager.isReady(), "OnlyOfficeManager.createWithFile was not ready");
    await withDocumentReady(() => manager.openFile(file));
    await assertExport(manager, FILE_TYPE.XLSX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.file);
  });

  await runStep("text fallback files", async () => {
    const textDocx = await fetchPublicFile(
      "/e2e/fixtures/plain-text-as-docx.docx",
      "plain-text-as-docx.docx",
    );
    const docxManager = await withDocumentReady(() =>
      OnlyOfficeManager.createWithFile(
        {
          containerId: CONTAINER_IDS.textFallback,
          fileType: FILE_TYPE.DOCX,
          defaultFileName: textDocx.name,
          readOnly: false,
          theme: DEFAULT_OFFICE_THEME,
        },
        textDocx,
      ),
    );

    assert(docxManager.isReady(), "Text DOCX fallback was not ready");
    await assertExport(docxManager, FILE_TYPE.DOCX);
    docxManager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.textFallback);

    const textXlsx = await fetchPublicFile(
      "/e2e/fixtures/plain-text-as-xlsx.xlsx",
      "plain-text-as-xlsx.xlsx",
    );
    const xlsxManager = await withDocumentReady(() =>
      OnlyOfficeManager.createWithFile(
        {
          containerId: CONTAINER_IDS.textFallback,
          fileType: FILE_TYPE.XLSX,
          defaultFileName: textXlsx.name,
          readOnly: false,
          theme: DEFAULT_OFFICE_THEME,
        },
        textXlsx,
      ),
    );

    assert(xlsxManager.isReady(), "Text XLSX fallback was not ready");
    await assertExport(xlsxManager, FILE_TYPE.XLSX);
    xlsxManager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.textFallback);
  });

  await runStep("manager fromEditor", async () => {
    const editor = editorManagerFactory.get(CONTAINER_IDS.fromEditor);
    const manager = OnlyOfficeManager.fromEditor(editor, {
      containerId: CONTAINER_IDS.fromEditor,
      fileType: FILE_TYPE.DOCX,
      defaultFileName: "FromEditor.docx",
      readOnly: false,
      theme: DEFAULT_OFFICE_THEME,
    });

    await withDocumentReady(() =>
      manager.openDocument({
        fileName: "FromEditor.docx",
        isNew: true,
      }),
    );

    assert(manager.isReady(), "OnlyOfficeManager.fromEditor was not ready");
    assert(manager.getEditor() === editor, "fromEditor did not keep editor");
    await assertExport(manager, FILE_TYPE.DOCX);
    manager.destroy();
    editorManagerFactory.destroy(CONTAINER_IDS.fromEditor);
  });

  await runStep("manager factory destroyAll", async () => {
    onlyOfficeManagerFactory.destroyAll();
    assert(
      !onlyOfficeManagerFactory.get(CONTAINER_IDS.factory),
      "Factory destroyAll left a manager behind",
    );
    editorManagerFactory.destroyAll();
  });

  return steps;
}

function OnlyOfficeTestEditor({ containerId }: { containerId: string }) {
  return (
    <div
      className={`${ONLYOFFICE_CONTAINER_CONFIG.PARENT_CLASS_NAME} relative h-[420px] min-h-[420px] border border-neutral-200 bg-white`}
      data-onlyoffice-container-id={containerId}
    >
      <div id={containerId} className="absolute inset-0" />
    </div>
  );
}

export function OnlyOfficeFactoryE2EPage() {
  const params = useMemo<{ mode: ResourceMode; cdnOrigin: string }>(() => {
    if (typeof window === "undefined") {
      return {
        mode: "local",
        cdnOrigin: "",
      };
    }

    const search = new URLSearchParams(window.location.search);
    return {
      mode: search.get("mode") === "cdn" ? "cdn" : "local",
      cdnOrigin: search.get("cdnOrigin") || "http://127.0.0.1:3010",
    };
  }, []);

  const [result, setResult] = useState<ScenarioResult>({
    mode: params.mode,
    status: "idle",
    steps: [],
  });

  useEffect(() => {
    let disposed = false;

    window.requestAnimationFrame(() => {
      if (disposed) return;

      setResult({ mode: params.mode, status: "running", steps: [] });

      // 用例入口
      runScenario(params.mode, params.cdnOrigin)
        .then((steps) => {
          if (!disposed) {
            setResult({ mode: params.mode, status: "passed", steps });
          }
        })
        .catch((error) => {
          if (!disposed) {
            setResult((current) => ({
              ...current,
              status: "failed",
              error: error instanceof Error ? error.message : String(error),
            }));
          }
        });
    });

    return () => {
      disposed = true;
      resetAll();
      OnlyOfficeManager.resetStaticResource();
    };
  }, [params.cdnOrigin, params.mode]);

  return (
    <main className="min-h-screen bg-neutral-50 p-4 text-neutral-900">
      <section className="mb-4 border border-neutral-200 bg-white p-3">
        <h1 className="text-base font-semibold">OnlyOffice factory e2e</h1>
        <p className="text-sm text-neutral-600">
          <span data-testid="scenario-status">{result.status}</span>
          {" · "}
          <span>{result.mode}</span>
        </p>
        {result.error && (
          <p className="mt-2 text-sm text-red-600" data-testid="scenario-error">
            {result.error}
          </p>
        )}
        <pre
          className="mt-3 max-h-64 overflow-auto bg-neutral-950 p-3 text-xs text-neutral-50"
          data-testid="scenario-result"
        >
          {JSON.stringify(result, null, 2)}
        </pre>
      </section>

      <div className="grid gap-4">
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.factory} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.create} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.file} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.textFallback} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.fromEditor} />
        <OnlyOfficeTestEditor containerId={CONTAINER_IDS.fixture} />
      </div>
    </main>
  );
}
