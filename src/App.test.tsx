import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { MantineProvider, createTheme } from "@mantine/core";
import App from "./App";
import type { CheckKey, OtDocument, TestResult } from "./types";

const theme = createTheme({
  fontFamily: "Inter, sans-serif",
  primaryColor: "blue",
  defaultRadius: "md",
});

const draftKey = "create-ot-draft-v3";

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;

  class TestResizeObserver {
    observe(): void {}
    unobserve(): void {}
    disconnect(): void {}
  }

  Object.defineProperty(globalThis, "ResizeObserver", {
    writable: true,
    value: TestResizeObserver,
  });

  Object.defineProperty(window, "matchMedia", {
    writable: true,
    value: vi.fn().mockImplementation((query: string) => ({
      matches: false,
      media: query,
      onchange: null,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(),
    })),
  });

  window.localStorage.clear();
  container = document.createElement("div");
  document.body.appendChild(container);
  root = createRoot(container);
});

afterEach(() => {
  act(() => {
    root.unmount();
  });
  container.remove();
  window.localStorage.clear();
  vi.restoreAllMocks();
});

describe("App test block filters", () => {
  it("shows only matching tests inside blocks when the problem filter is selected", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createFilterDraft()));

    await act(async () => {
      root.render(
        <MantineProvider theme={theme} defaultColorScheme="light">
          <App />
        </MantineProvider>,
      );
    });

    await clickButton("Testes");

    expect(getFilterCount("Todos")).toBe("2");
    expect(getFilterCount("Sem imagens")).toBe("2");
    expect(getFilterCount("Com problema")).toBe("1");
    expect(hasInputValue("teste sem erro")).toBe(true);
    expect(hasInputValue("teste com problema")).toBe(true);

    await clickFilterButton("Com problema");

    expect(hasInputValue("teste sem erro")).toBe(false);
    expect(hasInputValue("teste com problema")).toBe(true);
  });
});

async function clickButton(label: string): Promise<void> {
  const button = Array.from(container.querySelectorAll("button")).find((element) =>
    element.textContent?.includes(label),
  );

  expect(button).toBeTruthy();

  await act(async () => {
    button?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

async function clickFilterButton(label: string): Promise<void> {
  const button = getFilterButton(label);

  expect(button).toBeTruthy();
  expect(button?.disabled).toBe(false);

  await act(async () => {
    button?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
  });
}

function getFilterButton(label: string): HTMLButtonElement | undefined {
  return Array.from(container.querySelectorAll<HTMLButtonElement>(".testBlockFilter")).find(
    (element) => element.textContent?.includes(label),
  );
}

function getFilterCount(label: string): string | null | undefined {
  return getFilterButton(label)?.querySelector("strong")?.textContent;
}

function hasInputValue(value: string): boolean {
  return Array.from(container.querySelectorAll<HTMLInputElement>("input")).some(
    (input) => input.value === value,
  );
}

function createFilterDraft(): OtDocument {
  return {
    metadata: {
      screen: "Cadastro de exemplo",
      responsible: "GABRIEL",
      date: "2026-06-08",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar filtro de blocos.",
    accessSteps: [{ id: "step-1", text: "Acessar a tela" }],
    permissionGroups: [
      {
        id: "macro-a",
        code: "MA",
        label: "Macro A",
        selected: true,
        microPermissions: [
          {
            id: "micro-at",
            code: "AT",
            label: "Atualizacao",
            selected: true,
          },
        ],
      },
    ],
    permissionBlocks: {
      "macro-a:micro-at": {
        tests: [
          {
            id: "test-ok",
            title: "teste sem erro",
            result: createResult({ sameBehavior: true }),
          },
          {
            id: "test-problem",
            title: "teste com problema",
            result: createResult({ possibleIssue: true }),
          },
        ],
      },
    },
  };
}

function createResult(checks: Partial<Record<CheckKey, boolean>>): TestResult {
  return {
    checks: {
      sameBehavior: false,
      possibleIssue: false,
      bothIssue: false,
      newIssue: false,
      errorReport: false,
      ...checks,
    },
    observations: "",
    legacyImages: [],
    newImages: [],
  };
}
