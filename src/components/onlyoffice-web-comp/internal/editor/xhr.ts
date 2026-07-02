export interface XHRMiddleware {
  (request: Request): Response | null | Promise<Response | null>;
}

export interface XHRProxyOptions {
  baseUrl?: string;
  shouldBypass?: (url: string, method: string) => boolean;
}

function isForbiddenRequestHeader(name: string) {
  const lowerName = name.toLowerCase();
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
    lowerName.startsWith("proxy-") ||
    lowerName.startsWith("sec-")
  );
}

/**
 * Creates an XMLHttpRequest proxy class that supports middleware
 * @param BaseXHR The original XMLHttpRequest class
 * @returns The enhanced XMLHttpRequest class
 */
export function createXHRProxy(
  BaseXHR = globalThis.XMLHttpRequest,
  options: XHRProxyOptions = {},
) {
  return class ProxyXMLHttpRequest extends BaseXHR {
    private static _middlewares: XHRMiddleware[] = [];

    private _isMocked: boolean = false;
    private _requestMethod: string = "GET";
    private _requestUrl: string = "";
    private _requestHeaders: Headers = new Headers();
    private _requestBody: any = null;
    private _responseHeaders: Headers = new Headers();

    /**
     * Register global middleware
     */
    static use(middleware: XHRMiddleware) {
      this._middlewares.push(middleware);
    }

    /**
     * Clear all middleware
     */
    static clearMiddlewares() {
      this._middlewares = [];
    }

    open(
      method: string,
      url: string | URL,
      async: boolean = true,
      username?: string | null,
      password?: string | null,
    ): void {
      const normalizedUrl = (() => {
        try {
          return options.baseUrl
            ? new URL(url.toString(), options.baseUrl).href
            : url.toString();
        } catch {
          return url.toString();
        }
      })();

      this._requestMethod = method;
      this._requestUrl = normalizedUrl;
      this._requestHeaders = new Headers();
      this._responseHeaders = new Headers();
      this._isMocked = false;

      // Call native open
      super.open(
        method,
        normalizedUrl,
        async,
        username ?? undefined,
        password ?? undefined,
      );
    }

    setRequestHeader(name: string, value: string): void {
      if (isForbiddenRequestHeader(name)) {
        return;
      }

      this._requestHeaders.append(name, value);

      // If it is not a mock request, also set it on the native XHR
      if (!this._isMocked) {
        super.setRequestHeader(name, value);
      }
    }

    getResponseHeader(name: string): string | null {
      if (this._isMocked) {
        return this._responseHeaders.get(name);
      }
      return super.getResponseHeader(name);
    }

    getAllResponseHeaders(): string {
      if (!this._isMocked) {
        return super.getAllResponseHeaders();
      }
      return Array.from(this._responseHeaders.entries())
        .map(([key, value]) => `${key}: ${value}`)
        .join("\r\n");
    }

    send(body?: Document | XMLHttpRequestBodyInit | null): void {
      this._requestBody = body;

      if (options.shouldBypass?.(this._requestUrl, this._requestMethod)) {
        super.send(body);
        return;
      }

      // Try to run middleware
      this._tryMiddlewares()
        .then((handled) => {
          if (!handled) {
            // No middleware handled it, use native send
            super.send(body);
          }
        })
        .catch((err) => {
          console.error("ProxyXMLHttpRequest middleware error:", err);
          // Fallback to native implementation on error
          super.send(body);
        });
    }

    private async _tryMiddlewares(): Promise<boolean> {
      // Create Request object
      let request: Request;
      try {
        const reqInit: RequestInit = {
          method: this._requestMethod,
          headers: this._requestHeaders,
          body: this._requestBody as BodyInit,
          mode: "cors",
        };

        if (this.withCredentials) {
          reqInit.credentials = "include";
        }

        request = new Request(this._requestUrl, reqInit);
        console.log("ProxyXHR created request:", {
          url: this._requestUrl,
          method: request.method,
          hasBody: !!request.body,
          originalBody: this._requestBody,
        });
      } catch (e) {
        // Unable to create Request, do not use middleware
        return false;
      }

      // Run middleware
      for (const mw of ProxyXMLHttpRequest._middlewares) {
        const response = await mw(request.clone());
        if (response) {
          this._isMocked = true;
          await this._handleMockResponse(response);
          return true;
        }
      }

      return false;
    }

    private async _handleMockResponse(response: Response) {
      this._responseHeaders = new Headers(response.headers);

      const emit = (event: Event) => {
        this.dispatchEvent(event);
      };

      // 1. Trigger loadstart
      emit(new ProgressEvent("loadstart"));

      // 2. HEADERS_RECEIVED (readyState = 2)
      Object.defineProperty(this, "readyState", {
        value: 2,
        writable: false,
        configurable: true,
      });
      emit(new Event("readystatechange"));

      // 3. LOADING (readyState = 3)
      Object.defineProperty(this, "readyState", {
        value: 3,
        writable: false,
        configurable: true,
      });
      emit(new Event("readystatechange"));

      try {
        // Read response body
        let responseData: any;

        if (this.responseType === "json") {
          responseData = await response.json();
        } else if (this.responseType === "arraybuffer") {
          responseData = await response.arrayBuffer();
        } else if (this.responseType === "blob") {
          responseData = await response.blob();
        } else if (this.responseType === "document") {
          const text = await response.text();
          responseData = new DOMParser().parseFromString(text, "text/xml");
        } else {
          responseData = await response.text();
        }

        // Set response properties
        Object.defineProperty(this, "status", {
          value: response.status,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "statusText", {
          value: response.statusText,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "response", {
          value: responseData,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseText", {
          value:
            typeof responseData === "string"
              ? responseData
              : JSON.stringify(responseData),
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseURL", {
          value: response.url,
          writable: false,
          configurable: true,
        });

        // 4. Trigger progress event
        emit(
          new ProgressEvent("progress", {
            lengthComputable: true,
            loaded: 100,
            total: 100,
          }),
        );

        // 5. DONE (readyState = 4)
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        emit(new Event("readystatechange"));

        // 6. Trigger load event
        emit(new ProgressEvent("load"));

        // 7. Trigger loadend event
        emit(new ProgressEvent("loadend"));
      } catch (e) {
        console.error("ProxyXHR: error handling response", e);

        // Set readyState to DONE
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        emit(new Event("readystatechange"));

        // Trigger error event
        emit(new ProgressEvent("error"));
        emit(new ProgressEvent("loadend"));
      }
    }
  };
}
