import type { EditorServer } from "./server";
import type { MockSocket } from "./socket";
import type { ScopedIoFactory } from "./install-proxies";
import { CROSS_ORIGIN_BRIDGE_MESSAGE } from "./cross-origin-protocol";

const BRIDGE_SOURCE = "onlyoffice-bridge";

type BridgeMessage = {
  source?: string;
  type?: string;
  frameEditorId?: string;
  requestId?: string;
  command?: string;
  payload?: Record<string, unknown>;
  result?: unknown;
  error?: string;
  event?: string;
  args?: unknown[];
  method?: string;
  url?: string;
  headers?: Record<string, string>;
  body?: string | null;
  bodyEncoding?: "base64";
  readOnly?: boolean;
  status?: number;
  responseHeaders?: Record<string, string>;
  responseBody?: string | null;
  responseType?: XMLHttpRequestResponseType;
};

type BridgeSession = {
  frameEditorId: string;
  server: EditorServer;
  createIo: ScopedIoFactory;
  iframe: HTMLIFrameElement;
  targetOrigin: string;
  socket: MockSocket | null;
  bridgeReady: boolean;
  handshakeSent: boolean;
  pendingReadOnly: boolean | null;
};

const sessions = new Map<string, BridgeSession>();
const pendingRequests = new Map<
  string,
  {
    resolve: (value: unknown) => void;
    reject: (reason?: unknown) => void;
    timer: number;
  }
>();
const pendingReadyWaiters = new Map<
  string,
  Set<{
    resolve: (session: BridgeSession) => void;
    reject: (reason?: unknown) => void;
    timer: number;
  }>
>();
const editorEventSubscribers = new Map<
  string,
  Map<string, Set<(args: unknown[]) => void>>
>();
let listenerInstalled = false;

function isBridgeMessage(data: unknown): data is BridgeMessage {
  return (
    typeof data === "object" &&
    data !== null &&
    (data as BridgeMessage).source === BRIDGE_SOURCE
  );
}

function getTargetOrigin(iframe: HTMLIFrameElement): string {
  try {
    return new URL(iframe.src, window.location.href).origin;
  } catch {
    return "*";
  }
}

function postToIframe(
  session: BridgeSession,
  message: BridgeMessage,
  targetWindow?: Window | null,
) {
  const target = targetWindow ?? session.iframe.contentWindow;
  if (!target) {
    return;
  }
  target.postMessage(
    { ...message, source: BRIDGE_SOURCE },
    session.targetOrigin,
  );
}

function getIframeByWindow(source: MessageEventSource | null) {
  if (!source || !("postMessage" in source)) {
    return null;
  }

  const frames = document.querySelectorAll<HTMLIFrameElement>(
    'iframe[name="frameEditor"]',
  );
  for (const frame of frames) {
    if (frame.contentWindow === source) {
      return frame;
    }
  }

  return null;
}

function updateSessionIframe(
  session: BridgeSession,
  iframe: HTMLIFrameElement,
) {
  if (session.iframe === iframe) {
    return;
  }

  detachSocket(session);
  session.iframe = iframe;
  session.targetOrigin = getTargetOrigin(iframe);
  session.bridgeReady = false;
  session.handshakeSent = false;
}

function attachSocket(session: BridgeSession) {
  if (session.socket) {
    return;
  }

  const socket = session.createIo();
  session.socket = socket;
  session.server.registerSocketTransport(socket);

  const nativeServerEmit = socket.server.emit.bind(socket.server);
  socket.server.emit = (event: string, ...args: unknown[]) => {
    postToIframe(session, {
      type: "socket:event",
      frameEditorId: session.frameEditorId,
      event,
      args,
    });
    return nativeServerEmit(event, ...args);
  };
}

function detachSocket(session: BridgeSession) {
  if (!session.socket) {
    return;
  }
  session.server.handleDisconnect({ socket: session.socket });
  session.socket = null;
  session.handshakeSent = false;
}

async function handleHttpRequest(
  session: BridgeSession,
  message: BridgeMessage,
) {
  const requestId = message.requestId;
  if (!requestId || !message.url || !message.method) {
    return;
  }

  try {
    const init: RequestInit = {
      method: message.method,
      headers: message.headers,
    };
    if (message.method !== "GET" && message.method !== "HEAD") {
      init.body = decodeRequestBody(message);
    }

    let response = await session.server.handleRequest(
      new Request(message.url, init),
    );
    if (!response) {
      response = await fetch(message.url, init);
    }

    const responseHeaders: Record<string, string> = {};
    response.headers.forEach((value, key) => {
      responseHeaders[key] = value;
    });

    let responseBody: string | null = null;
    if (message.responseType === "arraybuffer") {
      const buffer = await response.arrayBuffer();
      responseBody = arrayBufferToBase64(buffer);
    } else {
      responseBody = await response.text();
    }

    postToIframe(session, {
      type: "http:response",
      frameEditorId: session.frameEditorId,
      requestId,
      status: response.status,
      responseHeaders,
      responseBody,
      responseType: message.responseType,
    });
  } catch (error) {
    postToIframe(session, {
      type: "http:response",
      frameEditorId: session.frameEditorId,
      requestId,
      status: 0,
      responseBody: null,
    });
    console.error("[OnlyOfficeBridge] HTTP proxy failed:", error);
  }
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]!);
  }
  return btoa(binary);
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function decodeRequestBody(
  message: BridgeMessage,
): BodyInit | null | undefined {
  if (message.body == null) {
    return null;
  }
  if (message.bodyEncoding === "base64") {
    return base64ToArrayBuffer(message.body);
  }
  return message.body;
}

function describeBridgeState(frameEditorId: string) {
  const session = sessions.get(frameEditorId);
  if (!session) {
    return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: session is not registered`;
  }

  return `OnlyOffice cross-origin bridge is not ready for ${frameEditorId}: iframe=${session.iframe.src || "(empty)"}`;
}

function resolveReadyWaiters(frameEditorId: string, session: BridgeSession) {
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }

  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    window.clearTimeout(waiter.timer);
    waiter.resolve(session);
  });
}

function rejectReadyWaiters(frameEditorId: string, error: Error) {
  const waiters = pendingReadyWaiters.get(frameEditorId);
  if (!waiters) {
    return;
  }

  pendingReadyWaiters.delete(frameEditorId);
  waiters.forEach((waiter) => {
    window.clearTimeout(waiter.timer);
    waiter.reject(error);
  });
}

function waitForBridgeReady(frameEditorId: string, timeout: number) {
  const session = sessions.get(frameEditorId);
  if (session?.bridgeReady) {
    return Promise.resolve(session);
  }

  return new Promise<BridgeSession>((resolve, reject) => {
    const timer = window.setTimeout(() => {
      const waiters = pendingReadyWaiters.get(frameEditorId);
      waiters?.delete(waiter);
      if (waiters?.size === 0) {
        pendingReadyWaiters.delete(frameEditorId);
      }
      reject(new Error(describeBridgeState(frameEditorId)));
    }, timeout);
    const waiter = { resolve, reject, timer };

    let waiters = pendingReadyWaiters.get(frameEditorId);
    if (!waiters) {
      waiters = new Set();
      pendingReadyWaiters.set(frameEditorId, waiters);
    }
    waiters.add(waiter);
  });
}

function handleBridgeMessage(event: MessageEvent) {
  if (!isBridgeMessage(event.data)) {
    return;
  }

  const message = event.data;
  if (
    message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_RESPONSE &&
    message.requestId
  ) {
    const pending = pendingRequests.get(message.requestId);
    if (pending) {
      window.clearTimeout(pending.timer);
      pendingRequests.delete(message.requestId);
      if (message.error) {
        pending.reject(new Error(message.error));
      } else {
        pending.resolve(message.result);
      }
    }
    return;
  }

  if (
    message.type === CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_EVENT &&
    message.event
  ) {
    const frameEditorId = message.frameEditorId;
    const frameSubscribers = frameEditorId
      ? editorEventSubscribers.get(frameEditorId)
      : undefined;
    const subscribers = frameSubscribers?.get(message.event);
    subscribers?.forEach((handler) => {
      handler(message.args ?? []);
    });
    return;
  }

  const frameEditorId = message.frameEditorId;
  if (!frameEditorId) {
    return;
  }

  const session = sessions.get(frameEditorId);
  if (!session) {
    return;
  }

  const sourceIframe = getIframeByWindow(event.source);
  if (sourceIframe) {
    updateSessionIframe(session, sourceIframe);
  }

  switch (message.type) {
    case "hello": {
      session.bridgeReady = true;
      resolveReadyWaiters(frameEditorId, session);
      attachSocket(session);
      const targetWindow =
        event.source && "postMessage" in event.source
          ? (event.source as Window)
          : undefined;
      postToIframe(session, { type: "hello:ack", frameEditorId }, targetWindow);
      if (!session.handshakeSent) {
        session.handshakeSent = true;
        setTimeout(() => {
          if (session.socket) {
            session.server.sendCoAuthoringHandshake(session.socket);
          }
        }, 0);
      }
      if (session.pendingReadOnly !== null) {
        postToIframe(session, {
          type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
          frameEditorId,
          readOnly: session.pendingReadOnly,
        });
        session.pendingReadOnly = null;
      }
      break;
    }
    case "socket:emit": {
      if (!session.socket || !message.event) {
        break;
      }
      (
        session.socket as { emit: (event: string, ...args: unknown[]) => void }
      ).emit(message.event, ...(message.args ?? []));
      break;
    }
    case "http": {
      void handleHttpRequest(session, message);
      break;
    }
    default:
      break;
  }
}

export function registerCrossOriginBridge(
  frameEditorId: string,
  iframe: HTMLIFrameElement,
  server: EditorServer,
  createIo: ScopedIoFactory,
) {
  const existing = sessions.get(frameEditorId);
  if (existing) {
    existing.server = server;
    existing.createIo = createIo;
    updateSessionIframe(existing, iframe);
    return;
  }

  sessions.set(frameEditorId, {
    frameEditorId,
    server,
    createIo,
    iframe,
    targetOrigin: getTargetOrigin(iframe),
    socket: null,
    bridgeReady: false,
    handshakeSent: false,
    pendingReadOnly: null,
  });

  if (!listenerInstalled) {
    window.addEventListener("message", handleBridgeMessage);
    listenerInstalled = true;
  }
}

export function unregisterCrossOriginBridge(frameEditorId: string) {
  const session = sessions.get(frameEditorId);
  if (session) {
    detachSocket(session);
    sessions.delete(frameEditorId);
    rejectReadyWaiters(
      frameEditorId,
      new Error(
        `OnlyOffice cross-origin bridge was unregistered: ${frameEditorId}`,
      ),
    );
  }
  editorEventSubscribers.delete(frameEditorId);
}

export function setCrossOriginReadOnly(
  frameEditorId: string,
  readOnly: boolean,
) {
  const session = sessions.get(frameEditorId);
  if (!session) {
    return false;
  }

  session.pendingReadOnly = readOnly;
  if (!session.bridgeReady) {
    return true;
  }

  postToIframe(session, {
    type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_SET_READONLY,
    frameEditorId,
    readOnly,
  });
  session.pendingReadOnly = null;
  return true;
}

export function callCrossOriginEditor(
  frameEditorId: string,
  command: string,
  payload: Record<string, unknown> = {},
  timeout = 5000,
) {
  const requestId = `editor-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2)}`;

  return waitForBridgeReady(frameEditorId, timeout).then(
    (session) =>
      new Promise<unknown>((resolve, reject) => {
        const timer = window.setTimeout(() => {
          pendingRequests.delete(requestId);
          reject(
            new Error(`OnlyOffice cross-origin command timed out: ${command}`),
          );
        }, timeout);

        pendingRequests.set(requestId, { resolve, reject, timer });
        postToIframe(session, {
          type: CROSS_ORIGIN_BRIDGE_MESSAGE.EDITOR_COMMAND,
          frameEditorId,
          requestId,
          command,
          payload,
        });
      }),
  );
}

export function subscribeCrossOriginEditorEvent(
  frameEditorId: string,
  event: string,
  handler: (args: unknown[]) => void,
) {
  let frameSubscribers = editorEventSubscribers.get(frameEditorId);
  if (!frameSubscribers) {
    frameSubscribers = new Map();
    editorEventSubscribers.set(frameEditorId, frameSubscribers);
  }

  let subscribers = frameSubscribers.get(event);
  if (!subscribers) {
    subscribers = new Set();
    frameSubscribers.set(event, subscribers);
  }

  subscribers.add(handler);

  return () => {
    subscribers?.delete(handler);
    if (subscribers?.size === 0) {
      frameSubscribers?.delete(event);
    }
    if (frameSubscribers?.size === 0) {
      editorEventSubscribers.delete(frameEditorId);
    }
  };
}

export function canAccessIframeWindow(
  iframe: HTMLIFrameElement | null | undefined,
) {
  if (!iframe) {
    return false;
  }
  try {
    void iframe.contentWindow?.location.href;
    return true;
  } catch {
    return false;
  }
}

export function watchCrossOriginIframe(
  frameEditorId: string,
  getIframe: () => HTMLIFrameElement | null | undefined,
  server: EditorServer,
  createIo: ScopedIoFactory,
) {
  let registered = false;

  const tryRegister = () => {
    const iframe = getIframe();
    if (!iframe?.src || canAccessIframeWindow(iframe)) {
      return false;
    }
    if (registered) {
      registerCrossOriginBridge(frameEditorId, iframe, server, createIo);
      return true;
    }
    registered = true;
    registerCrossOriginBridge(frameEditorId, iframe, server, createIo);
    return true;
  };

  tryRegister();

  const timer = window.setInterval(() => {
    if (tryRegister()) {
      window.clearInterval(timer);
    }
  }, 10);

  const observer = new MutationObserver(() => {
    tryRegister();
  });
  observer.observe(document.body, { childList: true, subtree: true });

  return () => {
    window.clearInterval(timer);
    observer.disconnect();
    if (registered) {
      unregisterCrossOriginBridge(frameEditorId);
    }
  };
}
