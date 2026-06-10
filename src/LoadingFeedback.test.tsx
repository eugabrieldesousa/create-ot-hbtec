import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import type { ReactNode } from "react";
import { createRoot, type Root } from "react-dom/client";
import { MantineProvider, createTheme } from "@mantine/core";
import { LoadingFeedback } from "./LoadingFeedback";

const theme = createTheme({
  fontFamily: "Inter, sans-serif",
  primaryColor: "blue",
  defaultRadius: "md",
});

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;

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

  container = document.createElement("div");
  document.body.appendChild(container);
  root = createRoot(container);
});

afterEach(() => {
  act(() => {
    root.unmount();
  });
  container.remove();
  vi.restoreAllMocks();
});

describe("LoadingFeedback", () => {
  it("renders label, detail and accessible live status", async () => {
    await renderLoading(
      <LoadingFeedback
        variant="global"
        label="Exportando DOCX..."
        detail="Gerando arquivo."
      />,
    );

    const status = container.querySelector<HTMLElement>('[role="status"]');

    expect(status).toBeTruthy();
    expect(status?.getAttribute("aria-live")).toBe("polite");
    expect(status?.getAttribute("aria-atomic")).toBe("true");
    expect(status?.textContent ?? "").toContain("Exportando DOCX...");
    expect(status?.textContent ?? "").toContain("Gerando arquivo.");
  });

  it("applies visual variants", async () => {
    await renderLoading(
      <>
        <LoadingFeedback variant="global" label="Global" />
        <LoadingFeedback variant="inline" label="Inline" />
        <LoadingFeedback variant="compact" label="Compacto" />
      </>,
    );

    expect(container.querySelector(".loadingFeedback--global")?.textContent).toContain("Global");
    expect(container.querySelector(".loadingFeedback--inline")?.textContent).toContain("Inline");
    expect(container.querySelector(".loadingFeedback--compact")?.textContent).toContain(
      "Compacto",
    );
  });
});

async function renderLoading(element: ReactNode): Promise<void> {
  await act(async () => {
    root.render(
      <MantineProvider theme={theme} defaultColorScheme="light">
        {element}
      </MantineProvider>,
    );
  });
}
