import type { EditorServer } from "./server";
import type { MockSocket } from "./socket";
import type { ScopedIoFactory } from "./install-proxies";

const BRIDGE_SOURCE = "onlyoffice-bridge";

type BridgeMessage = {
  source?: string;
  type?: string;
  frameEditorId?: string;
  requestId?: string;
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
  target.postMessage({ ...message, source: BRIDGE_SOURCE }, session.targetOrigin);
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

function updateSessionIframe(session: BridgeSession, iframe: HTMLIFrameElement) {
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

async function handleHttpRequest(session: BridgeSession, message: BridgeMessage) {
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

    let response = await session.server.handleRequest(new Request(message.url, init));
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

function decodeRequestBody(message: BridgeMessage): BodyInit | null | undefined {
  if (message.body == null) {
    return null;
  }
  if (message.bodyEncoding === "base64") {
    return base64ToArrayBuffer(message.body);
  }
  return message.body;
}

function handleBridgeMessage(event: MessageEvent) {
  if (!isBridgeMessage(event.data)) {
    return;
  }

  const message = event.data;
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
      attachSocket(session);
      const targetWindow =
        event.source && "postMessage" in event.source
          ? (event.source as Window)
          : undefined;
      postToIframe(
        session,
        { type: "hello:ack", frameEditorId },
        targetWindow,
      );
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
          type: "editor:set-readonly",
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
      (session.socket as { emit: (event: string, ...args: unknown[]) => void }).emit(
        message.event,
        ...(message.args ?? []),
      );
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
  }
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
    type: "editor:set-readonly",
    frameEditorId,
    readOnly,
  });
  session.pendingReadOnly = null;
  return true;
}

export function canAccessIframeWindow(iframe: HTMLIFrameElement | null | undefined) {
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
