import {
  expect,
  test,
  type Browser,
  type Page,
  type TestInfo,
} from "playwright/test";
import {
  ONLYOFFICE_FACTORY_EXPECTED_STEPS,
  type ResourceMode,
  type ScenarioResult,
  type StepResult,
} from "./onlyoffice-factory.contract";

const cdnOrigin =
  process.env.ONLYOFFICE_E2E_CDN_ORIGIN ??
  `http://${process.env.PLAYWRIGHT_HOST ?? "127.0.0.1"}:${
    process.env.PLAYWRIGHT_CDN_PORT ?? 3010
  }`;

const scenarioTimeoutMs = 110_000;
const testTimeoutMs = scenarioTimeoutMs + 30_000;

type ScenarioRun = {
  result: ScenarioResult;
  consoleLines: string[];
  waitError?: string;
};

function stepFailureMessage(mode: ResourceMode, name: string, step?: StepResult) {
  if (!step) {
    return `[${mode}] missing scenario step: ${name}`;
  }

  return `[${mode}] ${name} ${step.status}${
    step.detail ? `\n${step.detail}` : ""
  }`;
}

function assertStepPassed(
  mode: ResourceMode,
  name: string,
  step?: StepResult,
): asserts step is StepResult {
  if (!step || step.status !== "passed") {
    throw new Error(stepFailureMessage(mode, name, step));
  }
}

function firstBlockingStep(result: ScenarioResult) {
  return result.steps.find((step) => step.status !== "passed");
}

async function readScenarioResult(page: Page) {
  const text = await page
    .getByTestId("scenario-result")
    .innerText({ timeout: 5_000 })
    .catch(() => null);
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text) as ScenarioResult;
  } catch {
    return null;
  }
}

async function attachScenarioRun(testInfo: TestInfo, run: ScenarioRun) {
  await testInfo.attach("scenario-result", {
    body: JSON.stringify(run.result, null, 2),
    contentType: "application/json",
  });

  if (run.consoleLines.length > 0) {
    await testInfo.attach("browser-console", {
      body: run.consoleLines.join("\n"),
      contentType: "text/plain",
    });
  }
}

async function runScenarioPage(
  browser: Browser,
  mode: ResourceMode,
): Promise<ScenarioRun> {
  const page = await browser.newPage();
  const consoleLines: string[] = [];
  let waitError: string | undefined;

  page.on("console", (message) => {
    consoleLines.push(`[${message.type()}] ${message.text()}`);
  });
  page.on("pageerror", (error) => {
    consoleLines.push(`[pageerror] ${error.message}`);
  });

  try {
    const url = new URL("/e2e/onlyoffice-factory", "http://e2e.local");
    url.searchParams.set("mode", mode);
    if (mode === "cdn") {
      url.searchParams.set("cdnOrigin", cdnOrigin);
    }

    await page.goto(`${url.pathname}${url.search}`, {
      waitUntil: "domcontentloaded",
    });
    await expect(page.getByTestId("scenario-status")).toBeVisible();

    await page
      .waitForFunction(
        () => {
          const status = document.querySelector(
            '[data-testid="scenario-status"]',
          )?.textContent;
          return status === "passed" || status === "failed";
        },
        undefined,
        { timeout: scenarioTimeoutMs },
      )
      .catch((error) => {
        waitError = error instanceof Error ? error.message : String(error);
      });

    const result = await readScenarioResult(page);
    return {
      result: result ?? {
        mode,
        status: "failed",
        steps: [],
        error: waitError ?? "Scenario result was not rendered",
      },
      consoleLines,
      waitError,
    };
  } finally {
    await page.close().catch(() => {});
  }
}

for (const mode of ["local", "cdn"] as const) {
  test.describe(`OnlyOffice factory APIs / ${mode}`, () => {
    test.describe.configure({ mode: "serial" });

    let run: ScenarioRun;

    test.beforeAll(async ({ browser }, testInfo) => {
      testInfo.setTimeout(testTimeoutMs);
      run = await runScenarioPage(browser, mode);
    });

    test(`${mode} / scenario boot`, async ({}, testInfo) => {
      await attachScenarioRun(testInfo, run);
      expect(run.result.mode).toBe(mode);
      expect(run.result.status).not.toBe("idle");
      expect(run.waitError, run.waitError).toBeUndefined();
    });

    for (const name of ONLYOFFICE_FACTORY_EXPECTED_STEPS) {
      test(`${mode} / ${name}`, async ({}, testInfo) => {
        const step = run.result.steps.find((item) => item.name === name);
        const blockingStep = firstBlockingStep(run.result);
        if (!step && blockingStep) {
          test.skip(
            true,
            `[${mode}] blocked after ${blockingStep.name}: ${
              blockingStep.detail ?? blockingStep.status
            }`,
          );
        }

        if (!step || step.status !== "passed") {
          await attachScenarioRun(testInfo, run);
        }

        assertStepPassed(mode, name, step);
      });
    }

    test(`${mode} / scenario contract`, async ({}, testInfo) => {
      const blockingStep = firstBlockingStep(run.result);
      if (blockingStep) {
        test.skip(
          true,
          `[${mode}] blocked after ${blockingStep.name}: ${
            blockingStep.detail ?? blockingStep.status
          }`,
        );
      }

      await attachScenarioRun(testInfo, run);
      expect(run.result.steps.map((step) => step.name)).toEqual(
        ONLYOFFICE_FACTORY_EXPECTED_STEPS,
      );
      expect(run.result.status, JSON.stringify(run.result, null, 2)).toBe(
        "passed",
      );
    });
  });
}
