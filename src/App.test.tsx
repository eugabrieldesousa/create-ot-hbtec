import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { MantineProvider, createTheme } from "@mantine/core";
import App from "./App";
import type { CheckKey, EvidenceImage, OtDocument, TeaDocument, TestResult } from "./types";

const theme = createTheme({
  fontFamily: "Inter, sans-serif",
  primaryColor: "blue",
  defaultRadius: "md",
});

const draftKey = "create-ot-draft-v3";
const teaDraftKey = "create-tea-draft-v1";

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
    expect(container.textContent ?? "").toContain("teste sem erro");
    expect(container.textContent ?? "").toContain("teste com problema");

    await clickFilterButton("Com problema");

    expect(container.querySelector(".workspaceContent")?.textContent ?? "").not.toContain(
      "teste sem erro",
    );
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").toContain(
      "teste com problema",
    );
  });
});

describe("App review validations", () => {
  it("keeps the global OT actions focused on draft, import and export", async () => {
    await renderApp();

    const text = container.textContent ?? "";

    expect(text).toContain("Importar DOCX");
    expect(text).not.toContain("Padrão em todos");
    expect(text).not.toContain("Próxima pendência");
  });

  it("complains about missing status, missing evidence, and missing problem observations", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await act(async () => {
      root.render(
        <MantineProvider theme={theme} defaultColorScheme="light">
          <App />
        </MantineProvider>,
      );
    });

    await clickButton("Revis");

    const text = container.textContent ?? "";

    expect(container.querySelector(".reviewTabBadge")?.textContent).toBe("7");
    expect(container.querySelector(".reviewIssueGroup--danger")).toBeTruthy();
    expect(text).toContain("Status do teste vazio");
    expect(text).toContain("Imagem do legado ausente");
    expect(text).toContain("Imagem do novo ausente");
    expect(text).toContain("Observacao obrigatoria");
    expect(text).toContain("Permissao sem teste");
  });

  it("does not complain when more than one status is selected", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await act(async () => {
      root.render(
        <MantineProvider theme={theme} defaultColorScheme="light">
          <App />
        </MantineProvider>,
      );
    });

    await clickButton("Revis");

    const text = container.textContent ?? "";

    expect(text).not.toContain("Mais de um status marcado");
    expect(text).not.toContain("Status do teste vazio");
    expect(container.querySelector(".reviewTabBadge")).toBeNull();
  });
});

describe("App quick status colors", () => {
  it("marks both, new, and errors with warning and danger tones", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await act(async () => {
      root.render(
        <MantineProvider theme={theme} defaultColorScheme="light">
          <App />
        </MantineProvider>,
      );
    });

    await clickButton("Testes");
    await clickElement(getToggleByControlPrefix("test-details"));

    expect(getQuickCheckButton("Ambos")?.classList.contains("quickCheck--warning")).toBe(true);
    expect(getQuickCheckButton("Novo")?.classList.contains("quickCheck--warning")).toBe(true);
    expect(getQuickCheckButton("Erros")?.classList.contains("quickCheck--danger")).toBe(true);
  });
});

describe("App OT card UX", () => {
  it("shows inline summaries on micro and test cards", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();
    await clickButton("Testes");

    const microCard = container.querySelector<HTMLElement>(".permissionBlock");
    const firstTestCard = container.querySelector<HTMLElement>(".testCard");

    expect(microCard?.textContent ?? "").toContain("2 testes");
    expect(microCard?.textContent ?? "").toContain("2 pendentes");
    expect(microCard?.textContent ?? "").toContain("1 com problema");
    expect(microCard?.textContent ?? "").toContain("6 pendencias");
    expect(microCard?.querySelector(".summaryChip--red")).toBeTruthy();

    expect(firstTestCard?.textContent ?? "").toContain("0/5 checks");
    expect(firstTestCard?.textContent ?? "").toContain("Legado 0");
    expect(firstTestCard?.textContent ?? "").toContain("Novo 0");
    expect(firstTestCard?.textContent ?? "").toContain("3 pendencias");
    expect(firstTestCard?.textContent ?? "").toContain("Sem status");
  });

  it("keeps collapsed tests as summaries and opens detailed editing only when expanded", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await clickButton("Testes");

    expect(container.textContent ?? "").toContain("teste com varios status");
    expect(hasInputValue("teste com varios status")).toBe(false);
    expect(container.textContent ?? "").not.toContain("Status rapido");
    expect(container.textContent ?? "").not.toContain("Status detalhado");

    await clickElement(getToggleByControlPrefix("test-details"));

    expect(hasInputValue("teste com varios status")).toBe(true);
    expect(container.textContent ?? "").toContain("Status rapido");
    expect(container.textContent ?? "").toContain("Status detalhado");
    expect(container.querySelector(".evidenceGrid")).toBeTruthy();
  });

  it("moves secondary OT actions into menus and keeps them working", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await clickButton("Testes");

    expect(container.textContent ?? "").not.toContain("Adicionar pacote");
    expect(container.textContent ?? "").not.toContain("Duplicar teste");

    await clickAriaButtonAt("Mais ações da micro");
    await clickMenuItem("Adicionar teste");
    expect(container.querySelectorAll(".testCard")).toHaveLength(2);

    await clickAriaButtonAt("Mais ações do teste");
    await clickMenuItem("Duplicar teste");
    expect(container.querySelectorAll(".testCard")).toHaveLength(3);

    await clickAriaButtonAt("Mais ações do teste");
    await clickMenuItem("Remover teste");
    await waitForBodyText("Remover teste?");
    await cancelDialog();
    expect(container.querySelectorAll(".testCard")).toHaveLength(3);

    await clickAriaButtonAt("Mais ações do teste");
    await clickMenuItem("Remover teste");
    await confirmDialog("Remover teste");
    expect(container.querySelectorAll(".testCard")).toHaveLength(2);
  });

  it("uses clearer global and document action labels", async () => {
    await renderApp();

    expect(container.textContent ?? "").toContain("Adicionar passo");

    await clickButton("Limpar documento");
    await waitForBodyText("Limpar rascunho atual?");
    await waitForBodyText("Limpar documento");

    await cancelDialog();
  });

  it("exposes core navigation and bulk permission form accessibly", async () => {
    await renderApp();

    expect(container.querySelector<HTMLAnchorElement>(".skipLink")?.getAttribute("href")).toBe(
      "#main-content",
    );
    expect(container.querySelector("main#main-content")).toBeTruthy();
    expect(
      container.querySelector('[role="tablist"][aria-label="Navegação do documento OT"]'),
    ).toBeTruthy();

    await clickButton("Permiss");

    const bulkPermissions = Array.from(
      container.querySelectorAll<HTMLTextAreaElement>("textarea"),
    ).find((textarea) => textarea.labels?.[0]?.textContent?.includes("Permissões em lote"));

    expect(bulkPermissions).toBeTruthy();
  });

  it("announces quick status buttons as pressed state controls", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await clickButton("Testes");
    await clickElement(getToggleByControlPrefix("test-details"));

    const sameBehaviorButton = getQuickCheckButton("OK");

    expect(sameBehaviorButton?.getAttribute("aria-pressed")).toBe("false");

    await clickElement(sameBehaviorButton);

    expect(getQuickCheckButton("OK")?.getAttribute("aria-pressed")).toBe("true");
  });

  it("keeps OT FAQ out of the tab flow and opens it from the top help action", async () => {
    await renderApp();

    expect(Array.from(container.querySelectorAll('[role="tab"]')).map((tab) => tab.textContent)).not.toContain(
      "FAQ",
    );

    await clickButton("Ajuda");

    expect(document.body.textContent ?? "").toContain("FAQ do Gerador de OT");
  });
});

describe("App document outline", () => {
  it("shows an app-wide OT outline with groups and pending indicators", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();

    const outline = container.querySelector<HTMLElement>(".documentOutline");

    expect(outline).toBeTruthy();
    expect(outline?.textContent ?? "").toContain("Índice");
    expect(outline?.textContent ?? "").toContain("Documento");
    expect(outline?.textContent ?? "").toContain("Passo a passo");
    expect(outline?.textContent ?? "").toContain("Permissões");
    expect(outline?.textContent ?? "").toContain("Testes");
    expect(outline?.textContent ?? "").toContain("Revisão");
    expect(outline?.textContent ?? "").toContain("MA (Macro A)");
    expect(outline?.textContent ?? "").toContain("teste sem status");
    expect(outline?.textContent ?? "").toContain("Pendente");
  });

  it("opens the OT test path from the outline", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();

    expect(container.querySelector("#test-card-macro-a-micro-at-test-missing-status")).toBeNull();

    await clickOutlineItem("teste sem status");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 100));
    });

    expect(container.textContent ?? "").toContain("Status rapido");
    expect(document.activeElement?.id).toBe("test-card-macro-a-micro-at--test-missing-status");
  });

  it("shows and navigates the TEA outline with real titles and block targets", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createMixedTeaPendingDraft()));

    await renderApp();
    await clickControl("TEA");

    const outline = container.querySelector<HTMLElement>(".documentOutline");

    expect(outline?.textContent ?? "").toContain("Documento TEA");
    expect(outline?.textContent ?? "").toContain("Imagem geral");
    expect(outline?.textContent ?? "").toContain("Atividade com pendência");
    expect(outline?.textContent ?? "").toContain("Bloco 1 - Texto");
    expect(outline?.textContent ?? "").toContain("Pendente");

    await clickButton("Atividades");
    await clickButton("Recolher todos");
    expect(
      getToggleByControlId("tea-content-block-pending-empty-text-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("false");

    await clickButton("Documento");
    await clickOutlineItem("Bloco 1 - Texto");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 100));
    });

    expect(
      getToggleByControlId("tea-content-block-pending-empty-text-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("true");
    expect(document.activeElement?.id).toBe("tea-content-block-pending-empty-text");
  });
});

describe("App TEA content blocks", () => {
  it("shows the DOCX import action in TEA mode", async () => {
    await renderApp();
    await clickControl("TEA");

    expect(container.textContent ?? "").toContain("Importar DOCX");
  });

  it("migrates legacy activity content into reorderable blocks", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createLegacyTeaDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    expect(getTeaBlockTexts()).toEqual([
      expect.stringContaining("Texto"),
      expect.stringContaining("Lista"),
      expect.stringContaining("Imagens"),
    ]);

    await clickAriaButtonAt("Mais ações do bloco", 1);
    await clickMenuItem("Mover bloco para cima");

    expect(getTeaBlockTexts()).toEqual([
      expect.stringContaining("Lista"),
      expect.stringContaining("Texto"),
      expect.stringContaining("Imagens"),
    ]);

    await clickButton("Adicionar bloco");
    await clickMenuItem("Texto");
    expect(getTeaBlockTexts()).toHaveLength(4);

    await clickAriaButtonAt("Mais ações do bloco", 0);
    await clickMenuItem("Duplicar bloco");
    expect(getTeaBlockTexts()).toHaveLength(5);

    await clickAriaButtonAt("Mais ações do bloco", 0);
    await clickMenuItem("Remover bloco");
    await confirmDialog("Remover bloco");
    expect(getTeaBlockTexts()).toHaveLength(4);
  });

  it("confirms removal only when a TEA block has filled content", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createLegacyTeaDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    expect(getTeaBlockTexts()).toHaveLength(3);

    await clickAriaButtonAt("Mais ações do bloco", 1);
    await clickMenuItem("Remover bloco");
    await waitForBodyText("Remover bloco?");
    await cancelDialog();

    expect(getTeaBlockTexts()).toHaveLength(3);

    await clickAriaButtonAt("Mais ações do bloco", 1);
    await clickMenuItem("Remover bloco");
    await confirmDialog("Remover bloco");

    expect(getTeaBlockTexts()).toHaveLength(2);

    await clickButton("Adicionar bloco");
    await clickMenuItem("Texto");
    expect(getTeaBlockTexts()).toHaveLength(3);

    await clickAriaButtonAt("Mais ações do bloco", 2);
    await clickMenuItem("Remover bloco");

    expect(container.textContent ?? "").not.toContain("Remover bloco?");
    expect(getTeaBlockTexts()).toHaveLength(2);
  });

  it("counts TEA images from activity and subactivity content blocks", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Revis");

    expect(getReviewMetricValue("Imagens")).toBe("2");
    expect(container.textContent ?? "").toContain("Nenhuma pendência encontrada.");
  });
  it("removes presets and exposes collapsible grouped TEA sections", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    const text = container.textContent ?? "";
    expect(text).not.toContain("Imagem antes da lista");
    expect(text).not.toContain("Imagem + lista");
    expect(text).not.toContain("Somente imagens");
    expect(container.querySelector(".teaSubActivityList .teaSubActivityCard")).toBeTruthy();

    const activityToggle = getToggleByControlPrefix("tea-activity");
    const composerToggle = getToggleByControlPrefix("tea-composer");
    const blockToggle = getToggleByControlPrefix("tea-content-block");

    expect(activityToggle?.getAttribute("aria-expanded")).toBe("true");
    expect(composerToggle?.getAttribute("aria-expanded")).toBe("true");
    expect(blockToggle?.getAttribute("aria-expanded")).toBe("true");

    await clickButton("Recolher todos");
    expect(getToggleByControlPrefix("tea-activity")?.getAttribute("aria-expanded")).toBe(
      "false",
    );
    expect(getToggleByControlPrefix("tea-composer")?.getAttribute("aria-expanded")).toBe(
      "false",
    );
    expect(getToggleByControlPrefix("tea-content-block")?.getAttribute("aria-expanded")).toBe(
      "false",
    );

    await clickButton("Expandir todos");
    expect(getToggleByControlPrefix("tea-activity")?.getAttribute("aria-expanded")).toBe(
      "true",
    );
    expect(getToggleByControlPrefix("tea-composer")?.getAttribute("aria-expanded")).toBe(
      "true",
    );
    expect(getToggleByControlPrefix("tea-content-block")?.getAttribute("aria-expanded")).toBe(
      "true",
    );

    await clickElement(activityToggle);
    expect(getToggleByControlPrefix("tea-activity")?.getAttribute("aria-expanded")).toBe(
      "false",
    );

    await clickElement(getToggleByControlPrefix("tea-activity"));
    await clickElement(getToggleByControlPrefix("tea-composer"));
    expect(getToggleByControlPrefix("tea-composer")?.getAttribute("aria-expanded")).toBe(
      "false",
    );
  });

  it("shows TEA activity summaries", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    const activityCard = container.querySelector<HTMLElement>(".teaActivityCard");
    const subActivityCard = container.querySelector<HTMLElement>(".teaSubActivityCard");
    const text = container.textContent ?? "";

    expect(activityCard?.textContent ?? "").toContain("2.1 Atividade com imagens");
    expect(activityCard?.textContent ?? "").toContain("3 blocos");
    expect(activityCard?.textContent ?? "").toContain("1 subtópico");
    expect(activityCard?.textContent ?? "").toContain("2 imagens");
    expect(activityCard?.textContent ?? "").not.toContain("0 pendências");
    expect(subActivityCard?.textContent ?? "").toContain("2.1.1 Subtopico com imagem");
    expect(text).toContain("Blocos de Atividade com imagens");
    expect(text).toContain("Blocos de Subtopico com imagem");
    expect(text).toContain("Bloco 1 - Texto");
    expect(text).not.toContain("Conteúdo");
    expect(text).toContain("Legenda da imagem");
    expect(text).toContain("Adicionar imagem");
  });

  it("shows contextual empty state for TEA activity blocks", async () => {
    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    const activityTitle = container.querySelector<HTMLInputElement>(
      "#tea-activity-title-tea-activity-default",
    );
    const text = container.textContent ?? "";

    expect(text).toContain("2.1 Atividade sem título");
    expect(text).toContain("Blocos da atividade 2.1");
    expect(text).toContain("Nenhum bloco adicionado nesta atividade.");
    expect(text).toContain("Adicionar primeiro bloco");
    expect(activityTitle?.getAttribute("aria-invalid")).toBe("true");
  });

  it("shows inline review state in incomplete TEA activities", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createIncompleteTeaDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    const incompleteActivity = container.querySelector<HTMLElement>(".teaActivityCard");

    expect(incompleteActivity?.textContent ?? "").toContain("2.1 Atividade sem título");
    expect(incompleteActivity?.textContent ?? "").toContain("2.1.1 Subtópico sem título");
    expect(incompleteActivity?.textContent ?? "").toContain("7 pendências");
    expect(incompleteActivity?.querySelector(".teaSummaryChip--red")).toBeTruthy();
  });

  it("shows TEA inline errors and accessible draft status", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createIncompleteTeaDraft()));

    await renderApp();
    await clickControl("TEA");

    const status = container.querySelector<HTMLElement>('[role="status"]');
    const serviceOrderInput = container.querySelector<HTMLInputElement>(
      "#tea-metadata-service-order",
    );

    expect(status?.getAttribute("aria-live")).toBe("polite");
    expect(status?.textContent ?? "").toContain("pendentes");
    expect(serviceOrderInput?.getAttribute("aria-invalid")).toBe("true");

    await clickButton("Atividades");

    const activityTitle = container.querySelector<HTMLInputElement>(
      "#tea-activity-title-incomplete-activity",
    );
    const textBlockInput = container.querySelector<HTMLTextAreaElement>(
      "#tea-content-block-input-empty-text-block",
    );

    expect(activityTitle?.getAttribute("aria-invalid")).toBe("true");
    expect(textBlockInput?.getAttribute("aria-invalid")).toBe("true");
  });

  it("focuses the TEA field selected from the review panel", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createIncompleteTeaDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Revis");

    expect(container.textContent ?? "").toContain("Documento");
    expect(container.textContent ?? "").toContain("Atividades");
    expect(container.textContent ?? "").toContain("Imagens");
    expect(container.textContent ?? "").toContain("Corrigir agora");
    await clickReviewIssue("Bloco de texto vazio");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 80));
    });

    expect(document.activeElement?.id).toBe("tea-content-block-input-empty-text-block");
  });

  it("reports stricter TEA review issues", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createIncompleteTeaDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Revis");

    const text = container.textContent ?? "";

    expect(container.querySelector(".reviewIssueGroup--danger")).toBeTruthy();
    expect(text).toContain("Fase/Etapa vazia");
    expect(text).toContain("Chamado vazio");
    expect(text).toContain("Texto inicial vazio");
    expect(text).toContain("Negrito incompleto");
    expect(text).toContain("Bloco de texto vazio");
    expect(text).toContain("Lista vazia");
    expect(text).not.toContain("Imagem sem legenda");
    expect(text).toContain("Imagem sem dados");
    expect(text).toContain("Crítica");
    expect(text).toContain("Aviso");
  });

  it("expands only TEA activity panels that have pending review issues", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createMixedTeaPendingDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");
    await clickButton("Recolher todos");

    expect(getToggleByControlId("tea-activity-pending-activity-panel")?.getAttribute("aria-expanded")).toBe(
      "false",
    );
    expect(getToggleByControlId("tea-activity-clean-activity-panel")?.getAttribute("aria-expanded")).toBe(
      "false",
    );

    await clickButton("Expandir pendências");

    expect(getToggleByControlId("tea-activity-pending-activity-panel")?.getAttribute("aria-expanded")).toBe(
      "true",
    );
    expect(getToggleByControlId("tea-composer-pending-activity-panel")?.getAttribute("aria-expanded")).toBe(
      "true",
    );
    expect(getToggleByControlId("tea-content-block-pending-empty-text-panel")?.getAttribute("aria-expanded")).toBe(
      "true",
    );
    expect(getToggleByControlId("tea-activity-clean-activity-panel")?.getAttribute("aria-expanded")).toBe(
      "false",
    );
  });

  it("disables TEA pending expansion when there are no pending activities", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    const expandPendingButton = getButton("Expandir pendências");

    expect(expandPendingButton).toBeTruthy();
    expect(expandPendingButton?.disabled).toBe(true);
  });
});

async function renderApp(): Promise<void> {
  await act(async () => {
    root.render(
      <MantineProvider theme={theme} defaultColorScheme="light">
        <App />
      </MantineProvider>,
    );
  });
}

async function clickButton(label: string): Promise<void> {
  const button = getButton(label);

  expect(button).toBeTruthy();

  await act(async () => {
    button?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    button?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function getButton(label: string): HTMLButtonElement | undefined {
  return Array.from(container.querySelectorAll<HTMLButtonElement>("button")).find((element) =>
    element.textContent?.includes(label),
  );
}

async function confirmDialog(confirmLabel: string): Promise<void> {
  await clickDialogButton(confirmLabel);
}

async function cancelDialog(): Promise<void> {
  await clickDialogButton("Cancelar");
}

async function waitForBodyText(text: string): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    if ((document.body.textContent ?? "").includes(text)) {
      return;
    }

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
  }

  expect(document.body.textContent ?? "").toContain(text);
}

async function clickDocumentButton(label: string): Promise<void> {
  const root = document.body.querySelector<HTMLElement>('[role="dialog"]') ?? document.body;
  const button = Array.from(root.querySelectorAll<HTMLButtonElement>("button")).find(
    (element) => element.textContent?.includes(label),
  );

  expect(button).toBeTruthy();

  await act(async () => {
    button?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

async function clickDialogButton(label: string): Promise<void> {
  let dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');

  for (let attempt = 0; !dialog && attempt < 10; attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');
  }

  expect(dialog).toBeTruthy();

  const button = Array.from(dialog?.querySelectorAll<HTMLButtonElement>("button") ?? []).find(
    (element) => element.textContent?.includes(label),
  );

  expect(button).toBeTruthy();

  await act(async () => {
    button?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

async function clickControl(label: string): Promise<void> {
  const control = Array.from(container.querySelectorAll<HTMLElement>("button,label")).find(
    (element) => element.textContent?.includes(label),
  );

  expect(control).toBeTruthy();

  await act(async () => {
    control?.click();
  });
}

async function clickAriaButtonAt(label: string, index = 0): Promise<void> {
  const buttons = Array.from(
    container.querySelectorAll<HTMLButtonElement>(`button[aria-label="${label}"]`),
  ).filter((element) => !element.disabled);
  const button = buttons[index];

  expect(button).toBeTruthy();

  await act(async () => {
    button?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

async function clickMenuItem(label: string): Promise<void> {
  let item = getMenuItem(label);

  for (let attempt = 0; !item && attempt < 5; attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    item = getMenuItem(label);
  }

  expect(item).toBeTruthy();

  await act(async () => {
    item?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function getMenuItem(label: string): HTMLElement | undefined {
  return Array.from(
    document.body.querySelectorAll<HTMLElement>('[role="menuitem"]'),
  ).find((element) => element.textContent?.includes(label));
}

async function clickElement(element: HTMLElement | null | undefined): Promise<void> {
  expect(element).toBeTruthy();

  await act(async () => {
    element?.click();
  });
}

async function clickReviewIssue(label: string): Promise<void> {
  const issueButton = Array.from(
    container.querySelectorAll<HTMLButtonElement>(".reviewIssueButton"),
  ).find((button) => button.textContent?.includes(label));

  expect(issueButton).toBeTruthy();

  await act(async () => {
    issueButton?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

async function clickOutlineItem(label: string): Promise<void> {
  const item = Array.from(
    container.querySelectorAll<HTMLButtonElement>(".documentOutlineItem"),
  ).find((button) => button.textContent?.includes(label));

  expect(item).toBeTruthy();

  await act(async () => {
    item?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function getToggleByControlPrefix(prefix: string): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    `button[aria-controls^="${prefix}"][aria-expanded]`,
  );
}

function getToggleByControlId(controlId: string): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    `button[aria-controls="${controlId}"][aria-expanded]`,
  );
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

function getQuickCheckButton(label: string): HTMLButtonElement | undefined {
  return Array.from(container.querySelectorAll<HTMLButtonElement>(".quickCheck")).find(
    (button) => button.textContent === label,
  );
}

function getTeaBlockTexts(): string[] {
  return Array.from(container.querySelectorAll<HTMLElement>(".teaContentBlock")).map(
    (element) => element.textContent ?? "",
  );
}

function getReviewMetricValue(label: string): string | undefined {
  const metric = Array.from(container.querySelectorAll<HTMLElement>(".reviewMetric")).find(
    (element) => element.textContent?.includes(label),
  );

  return metric?.querySelectorAll("p")[1]?.textContent ?? undefined;
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

function createReviewDraft(): OtDocument {
  return {
    metadata: {
      screen: "Cadastro de exemplo",
      responsible: "GABRIEL",
      date: "2026-06-08",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar revisao de testes.",
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
          {
            id: "micro-co",
            code: "CO",
            label: "Consulta",
            selected: true,
          },
        ],
      },
    ],
    permissionBlocks: {
      "macro-a:micro-at": {
        tests: [
          {
            id: "test-missing-status",
            title: "teste sem status",
            result: createResult({}),
          },
          {
            id: "test-problem",
            title: "teste com erro",
            result: createResult({ possibleIssue: true }),
          },
        ],
      },
    },
  };
}

function createCompleteReviewDraft(): OtDocument {
  return {
    metadata: {
      screen: "Cadastro de exemplo",
      responsible: "GABRIEL",
      date: "2026-06-08",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar revisao com varios status.",
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
            id: "test-multiple-status",
            title: "teste com varios status",
            result: {
              ...createResult({
                bothIssue: true,
                newIssue: true,
                errorReport: true,
              }),
              observations: "Falha validada para revisao.",
              legacyImages: [createImage("legacy-image")],
              newImages: [createImage("new-image")],
            },
          },
        ],
      },
    },
  };
}

function createLegacyTeaDraft(): unknown {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-06-08",
      author: "Gabriel Sousa",
    },
    overview: "Validar novo layout.",
    activityIntro: "Atividades realizadas:",
    activityImages: [],
    activities: [
      {
        id: "legacy-activity",
        title: "Atividade legada",
        description: "Texto migrado",
        items: [{ id: "legacy-item", text: "Item migrado" }],
        images: [createImage("legacy-tea-image")],
        subActivities: [],
      },
    ],
  };
}

function createTeaDraftWithBlockImages(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-06-08",
      author: "Gabriel Sousa",
    },
    overview: "Validar novo layout.",
    activityIntro: "Atividades realizadas:",
    activityImages: [],
    activities: [
      {
        id: "tea-activity",
        title: "Atividade com imagens",
        blocks: [
          {
            id: "activity-text",
            type: "text",
            text: "Descricao da atividade.",
          },
          {
            id: "activity-images",
            type: "images",
            images: [createImage("activity-image")],
          },
        ],
        subActivities: [
          {
            id: "tea-subactivity",
            title: "Subtopico com imagem",
            blocks: [
              {
                id: "subactivity-images",
                type: "images",
                images: [createImage("subactivity-image")],
              },
            ],
          },
        ],
      },
    ],
  };
}

function createMixedTeaPendingDraft(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-06-08",
      author: "Gabriel Sousa",
    },
    overview: "Validar novo layout.",
    activityIntro: "Atividades realizadas:",
    activityImages: [],
    activities: [
      {
        id: "pending-activity",
        title: "Atividade com pendência",
        blocks: [
          {
            id: "pending-empty-text",
            type: "text",
            text: "",
          },
        ],
        subActivities: [],
      },
      {
        id: "clean-activity",
        title: "Atividade completa",
        blocks: [
          {
            id: "clean-text",
            type: "text",
            text: "Conteúdo preenchido.",
          },
        ],
        subActivities: [],
      },
    ],
  };
}

function createIncompleteTeaDraft(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "",
      phase: "",
      ticket: "",
      subject: "",
      date: "",
      author: "",
    },
    overview: "Visao geral com **negrito aberto",
    activityIntro: "",
    activityImages: [
      {
        id: "general-missing-data",
        label: "",
        name: "general.png",
        width: 10,
        height: 10,
      },
    ],
    activities: [
      {
        id: "incomplete-activity",
        title: "",
        blocks: [
          {
            id: "empty-text-block",
            type: "text",
            text: "",
          },
          {
            id: "empty-list-block",
            type: "list",
            items: [],
          },
          {
            id: "bad-list-block",
            type: "list",
            items: [{ id: "bad-list-item", text: "Item com **negrito aberto" }],
          },
          {
            id: "image-block",
            type: "images",
            images: [
              {
                id: "missing-image-data",
                label: "",
                name: "missing.png",
                width: 10,
                height: 10,
              },
            ],
          },
        ],
        subActivities: [
          {
            id: "incomplete-subactivity",
            title: "",
            blocks: [],
          },
        ],
      },
    ],
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

function createImage(id: string): EvidenceImage {
  return {
    id,
    label: "Evidencia",
    name: `${id}.png`,
    width: 10,
    height: 10,
    dataUrl: "data:image/png;base64,AAAA",
  };
}
