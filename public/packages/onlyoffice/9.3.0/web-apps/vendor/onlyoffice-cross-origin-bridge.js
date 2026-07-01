(function () {
  "use strict";

  var BRIDGE_SOURCE = "onlyoffice-bridge";
  var ASC_RESTRICTION_NONE = 0;
  var ASC_RESTRICTION_VIEW = 128;
  var requestSeq = 0;
  var pendingHttp = Object.create(null);

  function getFrameEditorId() {
    try {
      return new URLSearchParams(window.location.search).get("frameEditorId") || "";
    } catch (error) {
      return "";
    }
  }

  function isSameOriginParent() {
    try {
      if (!window.parent || window.parent === window) {
        return true;
      }
      return window.parent.location.origin === window.location.origin;
    } catch (error) {
      return false;
    }
  }

  var frameEditorId = getFrameEditorId();
  if (!frameEditorId) {
    return;
  }
  if (isSameOriginParent()) {
    return;
  }

  function post(message) {
    if (!window.parent || window.parent === window) {
      return;
    }
    window.parent.postMessage(
      Object.assign({}, message, {
        source: BRIDGE_SOURCE,
        frameEditorId: frameEditorId,
      }),
      "*",
    );
  }

  var handshakeReady = false;
  var queuedSocketEmits = [];
  var queuedSocketEvents = [];

  function flushSocketEmits() {
    if (!handshakeReady) {
      return;
    }
    while (queuedSocketEmits.length) {
      post(queuedSocketEmits.shift());
    }
  }

  function sendHello() {
    post({ type: "hello" });
  }

  var helloTimer = window.setInterval(sendHello, 50);
  sendHello();

  function createSocket(options) {
    var handlers = Object.create(null);
    var onceHandlers = Object.create(null);
    var managerHandlers = Object.create(null);
    var queuedMessageEvents = [];
    var manager = {
      opts: Object.assign({}, options || {}),
      setOpenToken: function () {
        return manager;
      },
      setSessionToken: function () {
        return manager;
      },
      on: function (event, handler) {
        if (typeof handler === "function") {
          (managerHandlers[event] || (managerHandlers[event] = [])).push(handler);
        }
        return manager;
      },
      off: function (event, handler) {
        if (!event) {
          managerHandlers = Object.create(null);
          return manager;
        }
        if (!handler) {
          delete managerHandlers[event];
          return manager;
        }
        managerHandlers[event] = (managerHandlers[event] || []).filter(function (item) {
          return item !== handler;
        });
        return manager;
      },
      reconnectionAttempts: function () {
        return manager;
      },
      reconnectionDelay: function () {
        return manager;
      },
      reconnectionDelayMax: function () {
        return manager;
      },
      timeout: function () {
        return manager;
      },
      transports: function () {
        return manager;
      },
      upgrade: function () {
        return manager;
      },
      upgradeTransport: function () {
        return manager;
      },
      upgradeTimeout: function () {
        return manager;
      },
    };
    var socket = {
      active: true,
      connected: true,
      disconnected: false,
      recovered: false,
      id: frameEditorId,
      io: manager,
      nsp: "/",
      on: function (event, handler) {
        if (typeof handler !== "function") {
          return socket;
        }
        (handlers[event] || (handlers[event] = [])).push(handler);
        if (event === "message" && queuedMessageEvents.length) {
          var queued = queuedMessageEvents.splice(0);
          queued.forEach(function (args) {
            dispatchSocketEvent("message", args);
          });
        }
        return socket;
      },
      once: function (event, handler) {
        if (typeof handler !== "function") {
          return socket;
        }
        (onceHandlers[event] || (onceHandlers[event] = [])).push(handler);
        return socket;
      },
      off: function (event, handler) {
        if (!event) {
          handlers = Object.create(null);
          onceHandlers = Object.create(null);
          return socket;
        }
        if (!handler) {
          delete handlers[event];
          delete onceHandlers[event];
          return socket;
        }
        handlers[event] = (handlers[event] || []).filter(function (item) {
          return item !== handler;
        });
        onceHandlers[event] = (onceHandlers[event] || []).filter(function (item) {
          return item !== handler;
        });
        return socket;
      },
      removeListener: function (event, handler) {
        return socket.off(event, handler);
      },
      removeAllListeners: function (event) {
        return socket.off(event);
      },
      send: function () {
        var args = ["message"].concat(Array.prototype.slice.call(arguments));
        return socket.emit.apply(socket, args);
      },
      emit: function (event) {
        var args = Array.prototype.slice.call(arguments, 1);
        var message = { type: "socket:emit", event: event, args: args };
        if (handshakeReady) {
          post(message);
        } else {
          queuedSocketEmits.push(message);
        }
        return socket;
      },
      open: function () {
        return socket.connect();
      },
      connect: function () {
        socket.connected = true;
        socket.disconnected = false;
        sendHello();
        dispatchSocketEvent("connect", []);
        return socket;
      },
      disconnect: function () {
        socket.connected = false;
        socket.disconnected = true;
        post({ type: "socket:emit", event: "disconnect", args: [] });
        dispatchSocketEvent("disconnect", []);
        return socket;
      },
      close: function () {
        return socket.disconnect();
      },
      compress: function () {
        return socket;
      },
    };

    function dispatchSocketEvent(event, args) {
      if (event === "message" && !(handlers.message || onceHandlers.message)) {
        queuedMessageEvents.push(args);
        return;
      }

      (handlers[event] || []).slice().forEach(function (handler) {
        handler.apply(socket, args);
      });

      var once = onceHandlers[event];
      if (once && once.length) {
        delete onceHandlers[event];
        once.slice().forEach(function (handler) {
          handler.apply(socket, args);
        });
      }
    }

    socket.__onlyofficeDispatch = dispatchSocketEvent;
    window.setTimeout(function () {
      dispatchSocketEvent("connect", []);
    }, 0);

    return socket;
  }

  function registerScopedSocket() {
    window.__ONLYOFFICE_SCOPED_IO__ = window.__ONLYOFFICE_SCOPED_IO__ || {};
    window.__ONLYOFFICE_SCOPED_IO__[frameEditorId] = function (url, options) {
      return createSocket(options);
    };
  }

  registerScopedSocket();

  function arrayBufferToBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var binary = "";
    for (var i = 0; i < bytes.length; i += 1) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  function base64ToArrayBuffer(base64) {
    var binary = atob(base64 || "");
    var bytes = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
  }

  function shouldProxyUrl(url) {
    if (!url || typeof url !== "string" || url.indexOf("blob:") === 0 || url.indexOf("data:") === 0) {
      return false;
    }
    try {
      var parsed = new URL(url, window.location.href);
      return (
        parsed.pathname.indexOf("/cache/files/") !== -1 ||
        parsed.pathname.indexOf("/downloadas/") !== -1 ||
        parsed.pathname.indexOf("/upload/") !== -1 ||
        parsed.pathname === "/plugins.json"
      );
    } catch (error) {
      return false;
    }
  }

  function isForbiddenRequestHeader(name) {
    var lowerName = String(name || "").toLowerCase();
    return (
      lowerName === "accept-charset" ||
      lowerName === "accept-encoding" ||
      lowerName === "access-control-request-headers" ||
      lowerName === "access-control-request-method" ||
      lowerName === "connection" ||
      lowerName === "content-length" ||
      lowerName === "cookie" ||
      lowerName === "cookie2" ||
      lowerName === "date" ||
      lowerName === "dnt" ||
      lowerName === "expect" ||
      lowerName === "host" ||
      lowerName === "keep-alive" ||
      lowerName === "origin" ||
      lowerName === "referer" ||
      lowerName === "te" ||
      lowerName === "trailer" ||
      lowerName === "transfer-encoding" ||
      lowerName === "upgrade" ||
      lowerName === "via" ||
      lowerName.indexOf("proxy-") === 0 ||
      lowerName.indexOf("sec-") === 0
    );
  }

  function setHeader(target, name, value) {
    if (!name || isForbiddenRequestHeader(name)) {
      return;
    }
    target[name] = value;
  }

  function headersToObject(headers) {
    var result = {};
    if (!headers) {
      return result;
    }
    if (headers instanceof Headers) {
      headers.forEach(function (value, key) {
        setHeader(result, key, value);
      });
      return result;
    }
    if (Array.isArray(headers)) {
      headers.forEach(function (entry) {
        setHeader(result, entry[0], entry[1]);
      });
      return result;
    }
    Object.keys(headers).forEach(function (key) {
      setHeader(result, key, headers[key]);
    });
    return result;
  }

  function normalizeBody(body) {
    if (body == null || typeof body === "string") {
      return Promise.resolve({ body: body == null ? null : body });
    }
    if (body instanceof ArrayBuffer) {
      return Promise.resolve({
        body: arrayBufferToBase64(body),
        bodyEncoding: "base64",
      });
    }
    if (ArrayBuffer.isView(body)) {
      return Promise.resolve({
        body: arrayBufferToBase64(body.buffer),
        bodyEncoding: "base64",
      });
    }
    if (body instanceof Blob) {
      return body.arrayBuffer().then(function (buffer) {
        return { body: arrayBufferToBase64(buffer), bodyEncoding: "base64" };
      });
    }
    return Promise.resolve({ body: String(body) });
  }

  function requestViaParent(method, url, headers, body, responseType) {
    return normalizeBody(body).then(function (normalized) {
      return new Promise(function (resolve) {
        var requestId = "bridge-" + Date.now() + "-" + requestSeq++;
        pendingHttp[requestId] = resolve;
        post({
          type: "http",
          requestId: requestId,
          method: method,
          url: new URL(url, window.location.href).href,
          headers: headersToObject(headers),
          body: normalized.body,
          bodyEncoding: normalized.bodyEncoding,
          responseType: responseType,
        });
      });
    });
  }

  function installFetchProxy() {
    var nativeFetch = window.fetch && window.fetch.bind(window);
    if (!nativeFetch) {
      return;
    }

    window.fetch = function bridgeFetch(input, init) {
      var request = input instanceof Request ? input : null;
      var url = request ? request.url : String(input);
      if (!shouldProxyUrl(url)) {
        return nativeFetch(input, init);
      }

      var method = (init && init.method) || (request && request.method) || "GET";
      var headers = (init && init.headers) || (request && request.headers);
      var body = init && "body" in init ? init.body : null;
      return requestViaParent(method, url, headers, body, "arraybuffer").then(function (message) {
        var responseBody = base64ToArrayBuffer(message.responseBody);
        return new Response(responseBody, {
          status: message.status || 200,
          headers: message.responseHeaders || {},
        });
      });
    };
  }

  function installXhrProxy() {
    var NativeXHR = window.XMLHttpRequest;
    if (!NativeXHR) {
      return;
    }

    window.XMLHttpRequest = function BridgeXMLHttpRequest() {
      var nativeXhr = new NativeXHR();
      var method = "GET";
      var url = "";
      var async = true;
      var requestHeaders = {};
      var responseHeaders = {};
      var listeners = {};
      var xhr = this;
      var responseType = "";

      this.readyState = 0;
      this.status = 0;
      this.statusText = "";
      this.response = null;
      this.responseText = "";
      this.onreadystatechange = null;
      this.onload = null;
      this.onerror = null;
      this.upload = nativeXhr.upload;

      Object.defineProperty(this, "responseType", {
        get: function () {
          return responseType;
        },
        set: function (value) {
          responseType = value || "";
          try {
            nativeXhr.responseType = responseType;
          } catch (error) {
            /* Native XHR can reject responseType changes after send. */
          }
        },
      });

      ["timeout", "withCredentials"].forEach(function (property) {
        Object.defineProperty(xhr, property, {
          get: function () {
            return nativeXhr[property];
          },
          set: function (value) {
            nativeXhr[property] = value;
          },
        });
      });

      function dispatch(eventName) {
        if (eventName === "readystatechange" && typeof xhr.onreadystatechange === "function") {
          xhr.onreadystatechange.call(xhr);
        }
        if (eventName === "load" && typeof xhr.onload === "function") {
          xhr.onload.call(xhr);
        }
        if (eventName === "error" && typeof xhr.onerror === "function") {
          xhr.onerror.call(xhr);
        }
        (listeners[eventName] || []).slice().forEach(function (handler) {
          handler.call(xhr, { type: eventName, target: xhr, currentTarget: xhr });
        });
      }

      function setReadyState(value) {
        xhr.readyState = value;
        dispatch("readystatechange");
      }

      this.open = function (nextMethod, nextUrl, nextAsync) {
        method = nextMethod || "GET";
        url = String(nextUrl || "");
        async = nextAsync !== false;
        if (!shouldProxyUrl(url)) {
          nativeXhr.open.apply(nativeXhr, arguments);
          return;
        }
        setReadyState(1);
      };

      this.send = function (body) {
        if (!shouldProxyUrl(url)) {
          copyNativeEvents(nativeXhr, xhr, dispatch);
          nativeXhr.send(body);
          return;
        }
        if (!async) {
          throw new Error("OnlyOffice bridge does not support synchronous XHR");
        }
        requestViaParent(method, url, requestHeaders, body, xhr.responseType || "").then(function (message) {
          responseHeaders = message.responseHeaders || {};
          xhr.status = message.status || 0;
          xhr.statusText = xhr.status >= 200 && xhr.status < 400 ? "OK" : "Error";
          setReadyState(2);
          setReadyState(3);
          if (xhr.responseType === "arraybuffer" || xhr.responseType === "blob") {
            var buffer = base64ToArrayBuffer(message.responseBody);
            xhr.response = xhr.responseType === "blob" ? new Blob([buffer]) : buffer;
          } else {
            xhr.responseText = message.responseBody || "";
            xhr.response = xhr.responseText;
          }
          setReadyState(4);
          dispatch("load");
          dispatch("loadend");
        });
      };

      this.abort = function () {
        if (!shouldProxyUrl(url)) {
          nativeXhr.abort();
        }
      };

      this.setRequestHeader = function (name, value) {
        if (isForbiddenRequestHeader(name)) {
          return;
        }
        if (!shouldProxyUrl(url)) {
          nativeXhr.setRequestHeader(name, value);
          return;
        }
        setHeader(requestHeaders, name, value);
      };

      this.overrideMimeType = function (mimeType) {
        if (nativeXhr.overrideMimeType) {
          nativeXhr.overrideMimeType(mimeType);
        }
      };

      this.getResponseHeader = function (name) {
        if (!shouldProxyUrl(url)) {
          return nativeXhr.getResponseHeader(name);
        }
        return responseHeaders[String(name).toLowerCase()] || responseHeaders[name] || null;
      };

      this.getAllResponseHeaders = function () {
        if (!shouldProxyUrl(url)) {
          return nativeXhr.getAllResponseHeaders();
        }
        return Object.keys(responseHeaders)
          .map(function (key) {
            return key + ": " + responseHeaders[key];
          })
          .join("\r\n");
      };

      this.addEventListener = function (eventName, handler) {
        (listeners[eventName] || (listeners[eventName] = [])).push(handler);
      };

      this.removeEventListener = function (eventName, handler) {
        listeners[eventName] = (listeners[eventName] || []).filter(function (item) {
          return item !== handler;
        });
      };
    };
  }

  function copyNativeEvents(nativeXhr, bridgeXhr, dispatch) {
    nativeXhr.onreadystatechange = function () {
      bridgeXhr.readyState = nativeXhr.readyState;
      bridgeXhr.status = nativeXhr.status;
      bridgeXhr.statusText = nativeXhr.statusText;
      bridgeXhr.response = nativeXhr.response;
      try {
        bridgeXhr.responseText = nativeXhr.responseText;
      } catch (error) {
        bridgeXhr.responseText = "";
      }
      dispatch("readystatechange");
    };
    nativeXhr.onload = function () {
      dispatch("load");
    };
    nativeXhr.onerror = function () {
      dispatch("error");
    };
  }

  function extractDownloadFileName(url) {
    try {
      var parsed = new URL(url, window.location.href);
      return parsed.searchParams.get("filename") || decodeURIComponent(parsed.pathname.split("/").pop() || "download");
    } catch (error) {
      return "download";
    }
  }

  function parseContentDispositionFileName(header) {
    if (!header) {
      return "";
    }
    var utf8 = /filename\*=UTF-8''([^;]+)/i.exec(header);
    if (utf8 && utf8[1]) {
      try {
        return decodeURIComponent(utf8[1]);
      } catch (error) {
        return utf8[1];
      }
    }
    var ascii = /filename="([^"]+)"/i.exec(header);
    return ascii && ascii[1] ? ascii[1] : "";
  }

  function triggerBlobDownload(blob, fileName) {
    var objectUrl = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = objectUrl;
    link.download = fileName || "download";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
  }

  function installNamedDownloadPatch(retries) {
    if (window.__ONLYOFFICE_GETFILE_PATCHED__) {
      return;
    }
    if (!window.AscCommon || !window.AscCommon.getFile) {
      if (retries > 0) {
        window.setTimeout(function () {
          installNamedDownloadPatch(retries - 1);
        }, 50);
      }
      return;
    }

    var nativeGetFile = window.AscCommon.getFile.bind(window.AscCommon);
    window.AscCommon.getFile = function (url) {
      if (typeof url !== "string" || url.indexOf("/cache/files/") === -1) {
        nativeGetFile(url);
        return;
      }
      var fallbackName = extractDownloadFileName(url);
      window.fetch(url)
        .then(function (response) {
          return response.blob().then(function (blob) {
            return {
              blob: blob,
              fileName: parseContentDispositionFileName(response.headers.get("Content-Disposition")) || fallbackName,
            };
          });
        })
        .then(function (result) {
          triggerBlobDownload(result.blob, result.fileName);
        })
        .catch(function () {
          nativeGetFile(url);
        });
    };
    window.__ONLYOFFICE_GETFILE_PATCHED__ = true;
  }

  function syncControllerReadOnly(readOnly) {
    try {
      var common = window.Common;
      var controller = common && common.Controllers && common.Controllers.Main;
      if (controller && controller.mode) {
        controller.mode.isEdit = !readOnly;
        controller.mode.canEdit = !readOnly;
      }
      if (controller && controller.editorConfig && controller.editorConfig.mode) {
        controller.editorConfig.mode.isEdit = !readOnly;
        controller.editorConfig.mode.canEdit = !readOnly;
      }
      if (common && common.NotificationCenter && common.NotificationCenter.trigger) {
        common.NotificationCenter.trigger("editing:disable", readOnly);
      }
    } catch (error) {
      /* best effort */
    }
  }

  function setReadOnly(readOnly) {
    try {
      var editor = (window.Asc && window.Asc.editor) || window.editor;
      if (!editor) {
        return false;
      }
      if (typeof editor.asc_setRestriction === "function") {
        if (readOnly) {
          editor.asc_setRestriction(ASC_RESTRICTION_VIEW);
        } else {
          if (typeof editor.asc_removeRestriction === "function") {
            editor.asc_removeRestriction(ASC_RESTRICTION_VIEW);
          }
          editor.asc_setRestriction(ASC_RESTRICTION_NONE);
        }
      }
      if (typeof editor.asc_setCanSendChanges === "function") {
        editor.asc_setCanSendChanges(!readOnly);
      }
      syncControllerReadOnly(readOnly);
      return true;
    } catch (error) {
      return false;
    }
  }

  function scheduleReadOnly(readOnly, retries) {
    if (setReadOnly(readOnly) || retries <= 0) {
      return;
    }
    window.setTimeout(function () {
      scheduleReadOnly(readOnly, retries - 1);
    }, 50);
  }

  window.addEventListener("message", function (event) {
    var message = event.data;
    if (!message || message.source !== BRIDGE_SOURCE || message.frameEditorId !== frameEditorId) {
      return;
    }

    if (message.type === "hello:ack") {
      handshakeReady = true;
      window.clearInterval(helloTimer);
      flushSocketEmits();
      return;
    }

    if (message.type === "socket:event") {
      var registry = window.__ONLYOFFICE_SCOPED_IO__;
      var sockets = registry && registry.__activeSockets;
      if (sockets && sockets.length) {
        sockets.slice().forEach(function (socket) {
          if (socket.__onlyofficeDispatch) {
            socket.__onlyofficeDispatch(message.event, message.args || []);
          }
        });
      } else {
        queuedSocketEvents.push({
          event: message.event,
          args: message.args || [],
        });
      }
      return;
    }

    if (message.type === "http:response" && message.requestId && pendingHttp[message.requestId]) {
      pendingHttp[message.requestId](message);
      delete pendingHttp[message.requestId];
      return;
    }

    if (message.type === "editor:set-readonly") {
      scheduleReadOnly(!!message.readOnly, 20);
    }
  });

  var originalFactory = window.__ONLYOFFICE_SCOPED_IO__[frameEditorId];
  window.__ONLYOFFICE_SCOPED_IO__.__activeSockets = window.__ONLYOFFICE_SCOPED_IO__.__activeSockets || [];
  window.__ONLYOFFICE_SCOPED_IO__[frameEditorId] = function () {
    var socket = originalFactory.apply(this, arguments);
    window.__ONLYOFFICE_SCOPED_IO__.__activeSockets.push(socket);
    if (queuedSocketEvents.length) {
      queuedSocketEvents.splice(0).forEach(function (queued) {
        socket.__onlyofficeDispatch(queued.event, queued.args || []);
      });
    }
    return socket;
  };

  installFetchProxy();
  installXhrProxy();
  installNamedDownloadPatch(100);
})();
