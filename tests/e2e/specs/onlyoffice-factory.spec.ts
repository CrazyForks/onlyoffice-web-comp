import { expect, test } from "playwright/test";
import {
  ONLYOFFICE_FACTORY_EXPECTED_STEPS,
  type ScenarioResult,
} from "./onlyoffice-factory.contract";

const cdnOrigin =
  process.env.ONLYOFFICE_E2E_CDN_ORIGIN ??
  `http://${process.env.PLAYWRIGHT_HOST ?? "127.0.0.1"}:${
    process.env.PLAYWRIGHT_CDN_PORT ?? 3010
  }`;

// factory API 在本地资源和独立 CDN 资源下都必须表现一致。
for (const mode of ["local", "cdn"] as const) {
  test(`OnlyOffice factory APIs work with ${mode} resources`, async ({
    page,
  }, testInfo) => {
    const consoleLines: string[] = [];

    // 按行附加浏览器日志，方便 CI 中定位 x2t/iframe 错误；
    // 用例通过时又不会把 reporter 输出刷得太吵。
    page.on("console", (message) => {
      const text = `[${message.type()}] ${message.text()}`;
      consoleLines.push(text);
      testInfo.attach(`console-${consoleLines.length}`, {
        body: text,
        contentType: "text/plain",
      });
    });

    page.on("pageerror", (error) => {
      consoleLines.push(`[pageerror] ${error.message}`);
    });

    // Next 路由只保留薄 wrapper；OnlyOffice SDK、worker、iframe 都必须在真实
    // 浏览器上下文中运行，所以生命周期和 API 断言放在页面侧 scenario 中。
    const url = new URL("/e2e/onlyoffice-factory", "http://e2e.local");
    url.searchParams.set("mode", mode);
    if (mode === "cdn") {
      url.searchParams.set("cdnOrigin", cdnOrigin);
    }

    await page.goto(`${url.pathname}${url.search}`, {
      waitUntil: "domcontentloaded",
    });

    await expect(page.getByTestId("scenario-status")).toBeVisible();

    // 等页面内 scenario 结束，而不是在 Playwright 里重复编排每个异步步骤；
    // 下方 JSON 结果是 browser scenario 和 spec 之间的稳定契约。
    await page.waitForFunction(
      () => {
        const status = document.querySelector(
          '[data-testid="scenario-status"]',
        )?.textContent;
        return status === "passed" || status === "failed";
      },
      undefined,
      { timeout: 60_000 },
    );

    const result = JSON.parse(
      await page.getByTestId("scenario-result").innerText(),
    ) as ScenarioResult;

    expect(result.mode).toBe(mode);
    expect(result.status, JSON.stringify(result, null, 2)).toBe("passed");
    // 显式校验步骤列表，避免已知边界场景被静默跳过；
    // 例如钉钉非法 bookmark 的 x2t 导入回归。
    expect(result.steps.map((step) => step.name)).toEqual(
      ONLYOFFICE_FACTORY_EXPECTED_STEPS,
    );
    expect(result.steps.every((step) => step.status === "passed")).toBe(true);
  });
}
