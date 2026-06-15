import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { MantineProvider, createTheme } from "@mantine/core";
import App from "./App";
import type { OtDocxImportResult } from "./docxImport";
import type { CheckKey, EvidenceImage, OtDocument, TeaDocument, TestResult } from "./types";

const appMocks = vi.hoisted(() => ({
  deleteEvidenceImageData: vi.fn(),
  deleteEvidenceImageDataBatch: vi.fn(),
  exportOtDocument: vi.fn(),
  exportTeaDocument: vi.fn(),
  hydrateDocumentImages: vi.fn(),
  hydrateTeaDocumentImages: vi.fn(),
  optimizeImageFile: vi.fn(),
  parseDocxFile: vi.fn(),
  persistEmbeddedEvidenceImages: vi.fn(),
  persistEmbeddedTeaImages: vi.fn(),
  saveEvidenceImageDataBatch: vi.fn(),
}));

vi.mock("./docxExport", () => ({
  exportOtDocument: appMocks.exportOtDocument,
  exportTeaDocument: appMocks.exportTeaDocument,
}));

vi.mock("./docxImport", () => ({
  parseDocxFile: appMocks.parseDocxFile,
}));

vi.mock("./imageOptimizer", () => ({
  optimizeImageFile: appMocks.optimizeImageFile,
}));

vi.mock("./imageStorage", () => ({
  deleteEvidenceImageData: appMocks.deleteEvidenceImageData,
  deleteEvidenceImageDataBatch: appMocks.deleteEvidenceImageDataBatch,
  hydrateDocumentImages: appMocks.hydrateDocumentImages,
  hydrateTeaDocumentImages: appMocks.hydrateTeaDocumentImages,
  persistEmbeddedEvidenceImages: appMocks.persistEmbeddedEvidenceImages,
  persistEmbeddedTeaImages: appMocks.persistEmbeddedTeaImages,
  saveEvidenceImageDataBatch: appMocks.saveEvidenceImageDataBatch,
}));

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

  appMocks.deleteEvidenceImageData.mockReset();
  appMocks.deleteEvidenceImageData.mockResolvedValue(undefined);
  appMocks.deleteEvidenceImageDataBatch.mockReset();
  appMocks.deleteEvidenceImageDataBatch.mockResolvedValue(undefined);
  appMocks.exportOtDocument.mockReset();
  appMocks.exportOtDocument.mockResolvedValue(undefined);
  appMocks.exportTeaDocument.mockReset();
  appMocks.exportTeaDocument.mockResolvedValue(undefined);
  appMocks.hydrateDocumentImages.mockReset();
  appMocks.hydrateDocumentImages.mockImplementation(async (documentData: OtDocument) => documentData);
  appMocks.hydrateTeaDocumentImages.mockReset();
  appMocks.hydrateTeaDocumentImages.mockImplementation(async (documentData: TeaDocument) => documentData);
  appMocks.optimizeImageFile.mockReset();
  appMocks.optimizeImageFile.mockResolvedValue({
    dataUrl: "data:image/png;base64,T1BUSU1JWkVE",
    width: 10,
    height: 10,
    originalBytes: 12,
    savedBytes: 10,
    optimized: true,
  });
  appMocks.parseDocxFile.mockReset();
  appMocks.persistEmbeddedEvidenceImages.mockReset();
  appMocks.persistEmbeddedEvidenceImages.mockImplementation(
    async (documentData: OtDocument) => documentData,
  );
  appMocks.persistEmbeddedTeaImages.mockReset();
  appMocks.persistEmbeddedTeaImages.mockImplementation(
    async (documentData: TeaDocument) => documentData,
  );
  appMocks.saveEvidenceImageDataBatch.mockReset();
  appMocks.saveEvidenceImageDataBatch.mockResolvedValue(undefined);

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

  Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
    configurable: true,
    value: vi.fn(),
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
    expect(
      Array.from(container.querySelectorAll(".reviewTabBadge")).some((badge) =>
        badge.getAttribute("aria-label")?.includes("revis"),
      ),
    ).toBe(false);
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
    expect(container.textContent ?? "").not.toContain("Status detalhado");
    expect(container.querySelector(".evidenceGrid")).toBeTruthy();
  });

  it("groups Novo tests in Para corrigir and reflects corrected state on test cards", async () => {
    const documentData = createCompleteReviewDraft();
    const test = documentData.permissionBlocks["macro-a:micro-at"].tests[0];
    test.correction = {
      corrected: false,
      hotfixTag: "hotfix 1.2.2",
      correctedBy: "Gabriel",
      cloudStage: "dev",
      beforeImages: [createImage("before-fix")],
      afterImages: [createImage("after-fix")],
    };
    window.localStorage.setItem(draftKey, JSON.stringify(documentData));

    await renderApp();
    await clickButton("Para corrigir");

    const correctionCard = container.querySelector<HTMLElement>(".correctionCard");

    expect(correctionCard?.textContent ?? "").toContain("teste com varios status");
    expect(correctionCard?.textContent ?? "").toContain("Antes 1");
    expect(correctionCard?.textContent ?? "").toContain("Depois 1");
    expect(correctionCard?.textContent ?? "").not.toContain("Falha validada para revisao.");
    expect(correctionCard?.textContent ?? "").not.toContain("Status rapido");
    expect(correctionCard?.textContent ?? "").not.toContain("Status detalhado");

    await clickElement(getToggleByControlPrefix("correction-details"));
    expect(correctionCard?.textContent ?? "").toContain("Falha validada para revisao.");
    expect(correctionCard?.textContent ?? "").toContain("Corrigido por");

    await clickButton("Marcar como corrigido");
    await clickButton("Testes");

    const testCard = container.querySelector<HTMLElement>(".testCard");
    expect(testCard?.textContent ?? "").toContain("Corrigido");
  });

  it("preserves multiline observations and groups Para corrigir by macro and micro", async () => {
    const documentData = createCorrectionDraft();
    documentData.permissionBlocks["macro-a:micro-at"].tests[0].result.observations =
      "Primeira linha\nSegunda linha";
    window.localStorage.setItem(draftKey, JSON.stringify(documentData));

    await renderApp();
    await clickButton("Para corrigir");

    expect(container.querySelectorAll(".correctionPermissionGroup")).toHaveLength(1);
    expect(container.querySelectorAll(".correctionMicroGroup")).toHaveLength(2);

    await clickElement(getToggleByControlPrefix("correction-details"));

    const observation = container.querySelector<HTMLElement>(".readonlyMultilineText");
    expect(observation?.textContent ?? "").toContain("Primeira linha\nSegunda linha");
  });

  it("filters Para corrigir and bulk-updates only visible items", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCorrectionDraft()));

    await renderApp();
    await clickButton("Para corrigir");

    expect(getFilterCount("Todos")).toBe("2");
    expect(getFilterCount("Pendentes")).toBe("1");
    expect(getFilterCount("Corrigidos")).toBe("1");
    expect(getFilterCount("Sem hotfix")).toBe("2");

    await clickFilterButton("Pendentes");
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").toContain(
      "pendente visivel",
    );
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").not.toContain(
      "corrigido oculto",
    );

    await setFieldValue(getInputByLabel("Tag da hotfix"), "hotfix bulk");
    await setFieldValue(getInputByLabel("Corrigido por"), "Ana");
    await clickButton("Atualizar todos");

    expect(container.querySelector(".workspaceContent")?.textContent ?? "").toContain(
      "hotfix bulk",
    );
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").toContain("Por Ana");

    await clickFilterButton("Corrigidos");
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").toContain(
      "corrigido oculto",
    );
    expect(container.querySelector(".workspaceContent")?.textContent ?? "").not.toContain(
      "hotfix bulk",
    );
  });

  it("opens an image preview from evidence thumbnails without showing caption fields", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await clickButton("Testes");
    await clickElement(getToggleByControlPrefix("test-details"));

    expect(container.textContent ?? "").not.toContain("Legenda da imagem");

    const previewButton = container.querySelector<HTMLButtonElement>(".imagePreviewButton");
    await clickElement(previewButton);

    await waitForDialogText("Pre-visualizacao da imagem");
    expect(document.body.textContent ?? "").toContain("legacy-image.png");
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
  it("shows a contextual OT outline with only test items", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();

    expect(container.querySelector<HTMLElement>(".documentOutline")).toBeNull();

    await clickButton("Testes");

    const outline = container.querySelector<HTMLElement>(".documentOutline");

    const outlineText = outline?.textContent ?? "";

    expect(outline).toBeTruthy();
    expect(outlineText).toContain("ndice");
    expect(outlineText).toContain("Testes");
    expect(outlineText).not.toContain("Passo a passo");
    expect(outlineText).not.toContain("Permiss");
    expect(outlineText).not.toContain("Revis");
    expect(outlineText).not.toContain("MA (Macro A)");
    expect(outlineText).toContain("teste sem status");
    expect(outlineText).toContain("Pendente");
  });

  it("opens the OT test path from the outline", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();

    expect(container.querySelector("#test-card-macro-a-micro-at-test-missing-status")).toBeNull();

    await clickButton("Testes");
    await clickOutlineItem("teste sem status");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 100));
    });

    expect(container.textContent ?? "").toContain("Status rapido");
    expect(document.activeElement?.id).toBe("test-card-macro-a-micro-at--test-missing-status");
  });

  it("hides and restores the document outline from the visible top button", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createReviewDraft()));

    await renderApp();
    await clickButton("Testes");

    expect(container.querySelector<HTMLElement>(".documentOutline")).toBeTruthy();

    await clickAriaButtonAt("Mais ações do documento");
    await clickMenuItem("Ocultar indice");

    expect(container.querySelector<HTMLElement>(".documentOutline")).toBeNull();
    expect(window.localStorage.getItem("create-ot:outline-hidden")).toBe("true");

    await act(async () => {
      root.unmount();
    });
    root = createRoot(container);

    await renderApp();
    await clickButton("Testes");

    expect(container.querySelector<HTMLElement>(".documentOutline")).toBeNull();
  });

  it("shows and opens the OT DOCX preview without replacing import or export actions", async () => {
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();

    const tabs = Array.from(container.querySelectorAll<HTMLElement>('[role="tab"]')).map(
      (tab) => tab.textContent,
    );

    expect(tabs).toContain("Prévia DOCX");
    expect(container.textContent ?? "").toContain("Importar DOCX");
    expect(container.textContent ?? "").toContain("Exportar DOCX");
    expect(container.querySelector(".docxPreviewPage")).toBeNull();

    await clickButton("Documento");
    await setFieldValue(getInputByLabel("Tela"), "Tela performatica");

    await clickButton("Prévia DOCX");

    const previewText = container.querySelector(".docxPreviewPage")?.textContent ?? "";
    expect(previewText).toContain("OBSERVABILIDADE DE TESTES");
    expect(previewText).toContain("Tela performatica");
    expect(container.textContent ?? "").toContain("Atualizada");

    await clickOutlineItem("teste com varios status");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 80));
    });

    expect(document.activeElement?.id).toBe(
      "ot-preview-test-macro-a-micro-at--test-multiple-status",
    );
  });

  it("shows and navigates the TEA outline with activity and subtopic targets", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createMixedTeaPendingDraft()));

    await renderApp();
    await clickControl("TEA");

    expect(container.querySelector<HTMLElement>(".documentOutline")).toBeNull();

    await clickButton("Atividades");
    const outline = container.querySelector<HTMLElement>(".documentOutline");

    expect(outline?.textContent ?? "").not.toContain("Documento TEA");
    expect(outline?.textContent ?? "").not.toContain("Imagem geral");
    expect(outline?.textContent ?? "").toContain("Atividade com pendência");
    expect(outline?.textContent ?? "").not.toContain("Bloco 1 - Texto");
    expect(outline?.textContent ?? "").not.toContain("1 bloco");
    expect(outline?.textContent ?? "").toContain("Pendente");

    await clickButton("Recolher todos");
    expect(
      getToggleByControlId("tea-content-block-pending-empty-text-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("false");

    vi.clearAllMocks();

    await clickOutlineItem("Atividade com pendência");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 100));
    });

    expect(
      getToggleByControlId("tea-activity-pending-activity-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("false");
    expect(HTMLElement.prototype.scrollIntoView).not.toHaveBeenCalled();

    await clickButton("Expandir todos");
    vi.clearAllMocks();
    await clickOutlineItem("Atividade com pendência");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 100));
    });

    expect(
      getToggleByControlId("tea-activity-pending-activity-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("true");
    expect(
      getToggleByControlId("tea-composer-pending-activity-panel")?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("true");
    expect(document.activeElement?.id).toBe("tea-activity-pending-activity");
    expect(HTMLElement.prototype.scrollIntoView).toHaveBeenCalledWith(
      expect.objectContaining({ block: "start" }),
    );
    const activeOutlineItems = getActiveOutlineItems();
    expect(activeOutlineItems).toHaveLength(1);
    expect(activeOutlineItems[0]?.textContent ?? "").toContain("Atividade com pendência");
  });

  it("updates the active TEA outline item while scrolling through activities and subtopics", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(function (
      this: HTMLElement,
    ) {
      const id = this.id;

      if (id === "tea-section-activities") {
        return createRect(-320);
      }

      if (id === "tea-activity-tea-activity") {
        return createRect(-120);
      }

      if (id === "tea-subactivity-tea-subactivity") {
        return createRect(96);
      }

      return createRect(900);
    });

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    await act(async () => {
      window.dispatchEvent(new Event("scroll"));
      await new Promise((resolve) => setTimeout(resolve, 40));
    });

    const activeOutlineItems = getActiveOutlineItems();

    expect(activeOutlineItems).toHaveLength(1);
    expect(activeOutlineItems[0]?.textContent ?? "").toContain("Subtopico com imagem");
  });

  it("shows and opens the TEA DOCX preview without replacing import or export actions", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await clickControl("TEA");

    const tabs = Array.from(container.querySelectorAll<HTMLElement>('[role="tab"]')).map(
      (tab) => tab.textContent,
    );

    expect(tabs).toContain("Prévia DOCX");
    expect(container.textContent ?? "").toContain("Importar DOCX");
    expect(container.textContent ?? "").toContain("Exportar DOCX");

    await clickButton("Prévia DOCX");

    const previewText = container.querySelector(".docxPreviewPage")?.textContent ?? "";
    expect(previewText).toContain("Termo de Entrega de Atividade");
    expect(previewText).toContain("Atividade com imagens");
    const outline = container.querySelector<HTMLElement>(".documentOutline");
    expect(outline?.textContent ?? "").toContain("Atividade com imagens");
    expect(outline?.textContent ?? "").toContain("Subtopico com imagem");
    expect(outline?.textContent ?? "").not.toContain("Prévia DOCX");

    await clickOutlineItem("Subtopico com imagem");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 80));
    });

    expect(document.activeElement?.id).toBe("tea-preview-subactivity-tea-subactivity");
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
    expect(text).not.toContain("Legenda da imagem");
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

  it("copies a TEA subtopic to the selected target activity", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftForSubActivityCopy()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    await clickAriaButtonAt("Mais ações do subtópico");
    await waitForBodyText("Copiar subtópico");
    await clickMenuItem("Copiar subtópico");
    await waitForDialogText("Atividade destino");
    await selectComboboxOption("Atividade destino", "2.3 Atividade destino B");
    await confirmDialog("Copiar subtópico");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 120));
    });

    const activityCards = Array.from(container.querySelectorAll<HTMLElement>(".teaActivityCard"));
    const destinationActivity = activityCards.find((activityCard) =>
      activityCard.textContent?.includes("2.3 Atividade destino B"),
    );

    expect(container.querySelectorAll(".teaSubActivityCard")).toHaveLength(2);
    expect(destinationActivity?.textContent ?? "").toContain("2.3.1 Subtopico origem");
    expect(destinationActivity?.textContent ?? "").toContain("Texto copiado");
    expect(destinationActivity?.textContent ?? "").toContain("Item copiado");
    expect(destinationActivity?.textContent ?? "").toContain("copy-source-image.png");
  });

  it("copies multiple TEA subtopics to one target activity", async () => {
    window.localStorage.setItem(
      teaDraftKey,
      JSON.stringify(createTeaDraftForBulkSubActivityCopy()),
    );

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    await clickButton("Copiar subt");
    await waitForDialogText("Selecionar todos");
    await selectComboboxOption("Atividade destino", "2.3 Atividade destino B");
    await clickDialogButton("Copiar subt");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 120));
    });

    const activityCards = Array.from(container.querySelectorAll<HTMLElement>(".teaActivityCard"));
    const destinationActivity = activityCards.find((activityCard) =>
      activityCard.textContent?.includes("2.3 Atividade destino B"),
    );

    expect(container.querySelectorAll(".teaSubActivityCard")).toHaveLength(4);
    expect(destinationActivity?.textContent ?? "").toContain("2.3.1 Subtopico origem");
    expect(destinationActivity?.textContent ?? "").toContain("2.3.2 Segundo subtopico");
    expect(destinationActivity?.textContent ?? "").toContain("Texto do segundo subtopico");
  });

  it("disables subtopic copy confirmation when there is no other activity", async () => {
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createSingleActivitySubtopicDraft()));

    await renderApp();
    await clickControl("TEA");
    await clickButton("Atividades");

    await clickAriaButtonAt("Mais ações do subtópico");
    await waitForBodyText("Copiar subtópico");
    await clickMenuItem("Copiar subtópico");
    await waitForDialogText("Crie outra atividade");

    const dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');
    const copyButton = Array.from(dialog?.querySelectorAll<HTMLButtonElement>("button") ?? [])
      .find((button) => button.textContent?.includes("Copiar subtópico"));

    expect(copyButton).toBeTruthy();
    expect(copyButton?.disabled).toBe(true);
  });

  it("shows a TEA-only floating action to scroll back to the top", async () => {
    const scrollTo = vi.fn();

    Object.defineProperty(window, "scrollTo", {
      writable: true,
      value: scrollTo,
    });

    await renderApp();

    expect(container.querySelector<HTMLButtonElement>('button[aria-label="Voltar ao topo"]')).toBeNull();

    await clickControl("TEA");
    await clickElement(container.querySelector<HTMLButtonElement>('button[aria-label="Voltar ao topo"]'));

    expect(scrollTo).toHaveBeenCalledWith({ top: 0, behavior: "smooth" });
  });
});

describe("App loading feedback", () => {
  it("shows global loading while exporting a DOCX", async () => {
    const exportTask = createDeferred<void>();
    appMocks.exportOtDocument.mockReturnValue(exportTask.promise);
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await waitForNoBodyText("Preparando imagens...");

    await clickElement(getButton("Exportar DOCX"));
    await waitForBodyText("Exportando DOCX...");

    expect(appMocks.exportOtDocument).toHaveBeenCalledTimes(1);
    expect(container.querySelector(".loadingFeedback--global")?.textContent ?? "").toContain(
      "Exportando DOCX...",
    );
    expect(getButton("Exportar DOCX")?.disabled).toBe(true);

    await act(async () => {
      exportTask.resolve();
      await exportTask.promise;
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    await waitForNoBodyText("Exportando DOCX...");
  });

  it("shows local loading while processing evidence images", async () => {
    const optimizeTask = createDeferred<{
      dataUrl: string;
      width: number;
      height: number;
      originalBytes: number;
      savedBytes: number;
      optimized: boolean;
    }>();
    appMocks.optimizeImageFile.mockReturnValue(optimizeTask.promise);
    window.localStorage.setItem(draftKey, JSON.stringify(createCompleteReviewDraft()));

    await renderApp();
    await waitForNoBodyText("Preparando imagens...");
    await clickButton("Testes");
    await clickElement(getToggleByControlPrefix("test-details"));
    await uploadFile('input[accept="image/*"]', "evidencia.png", "image/png");

    await waitForBodyText("Processando imagens...");
    expect(container.querySelector<HTMLElement>(".evidencePanel")?.getAttribute("aria-busy")).toBe(
      "true",
    );

    await act(async () => {
      optimizeTask.resolve({
        dataUrl: "data:image/png;base64,Tk9WQQ==",
        width: 12,
        height: 12,
        originalBytes: 20,
        savedBytes: 10,
        optimized: true,
      });
      await optimizeTask.promise;
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    await waitForNoBodyText("Processando imagens...");
    expect(container.textContent ?? "").toContain("evidencia.png");
  });

  it("keeps the import confirmation modal loading until images are persisted", async () => {
    const persistTask = createDeferred<OtDocument>();
    const importResult = createOtImportResult();
    appMocks.parseDocxFile.mockResolvedValue(importResult);

    await renderApp();
    await waitForNoBodyText("Preparando imagens...");
    appMocks.persistEmbeddedEvidenceImages.mockReturnValue(persistTask.promise);
    await uploadFile(
      'input[accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"]',
      "importado.docx",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    await waitForDialogText("Substituir rascunho");

    await clickDialogButton("Substituir rascunho");
    const confirmButton = getDialogButton("Substituir rascunho");

    expect(confirmButton?.disabled).toBe(true);
    expect(document.body.querySelector('[role="dialog"]')).toBeTruthy();

    await act(async () => {
      persistTask.resolve(importResult.document);
      await persistTask.promise;
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    await waitForNoDialog();
    expect(appMocks.persistEmbeddedEvidenceImages).toHaveBeenCalledWith(importResult.document);
  });

  it("shows loading while copying a TEA subtopic with images", async () => {
    const copyTask = createDeferred<void>();
    appMocks.saveEvidenceImageDataBatch.mockReturnValue(copyTask.promise);
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftForSubActivityCopy()));

    await renderApp();
    await waitForNoBodyText("Preparando imagens...");
    await clickControl("TEA");
    await clickButton("Atividades");

    await clickAriaButtonAt("Mais ações do subtópico");
    await clickMenuItem("Copiar subtópico");
    await waitForDialogText("Atividade destino");
    await clickDialogButton("Copiar subtópico");

    expect(getDialogButton("Copiar subtópico")?.disabled).toBe(true);

    await act(async () => {
      copyTask.resolve();
      await copyTask.promise;
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    await waitForNoDialog();
    expect(container.querySelectorAll(".teaSubActivityCard")).toHaveLength(2);
  });

  it("shows compact loading and prevents duplicate TEA block clicks", async () => {
    const duplicateTask = createDeferred<void>();
    appMocks.saveEvidenceImageDataBatch.mockReturnValue(duplicateTask.promise);
    window.localStorage.setItem(teaDraftKey, JSON.stringify(createTeaDraftWithBlockImages()));

    await renderApp();
    await waitForNoBodyText("Preparando imagens...");
    await clickControl("TEA");
    await clickButton("Atividades");

    await clickAriaButtonAt("Mais ações do bloco", 1);
    await clickMenuItem("Duplicar bloco");
    await waitForBodyText("Duplicando bloco...");

    const blockActions = container.querySelectorAll<HTMLButtonElement>(
      'button[aria-label="Mais ações do bloco"]',
    );
    expect(blockActions[1]?.disabled).toBe(true);
    expect(container.querySelector(".loadingFeedback--compact")?.textContent ?? "").toContain(
      "Duplicando bloco...",
    );

    await act(async () => {
      duplicateTask.resolve();
      await duplicateTask.promise;
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    await waitForNoBodyText("Duplicando bloco...");
    expect(getTeaBlockTexts()).toHaveLength(4);
    expect(appMocks.saveEvidenceImageDataBatch).toHaveBeenCalledTimes(1);
  });
});

async function renderApp(): Promise<void> {
  await act(async () => {
    root.render(
      <MantineProvider theme={theme} defaultColorScheme="light">
        <App />
      </MantineProvider>,
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
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

async function waitForDialogText(text: string): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    const dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');

    if ((dialog?.textContent ?? "").includes(text)) {
      return;
    }

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
  }

  const dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');

  expect(dialog?.textContent ?? "").toContain(text);
}

async function waitForNoBodyText(text: string): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    if (!(document.body.textContent ?? "").includes(text)) {
      return;
    }

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
  }

  expect(document.body.textContent ?? "").not.toContain(text);
}

async function waitForNoDialog(): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    if (!document.body.querySelector('[role="dialog"]')) {
      return;
    }

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
  }

  expect(document.body.querySelector('[role="dialog"]')).toBeNull();
}

function getDialogButton(label: string): HTMLButtonElement | undefined {
  const dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');

  return Array.from(dialog?.querySelectorAll<HTMLButtonElement>("button") ?? []).find(
    (element) => element.textContent?.includes(label),
  );
}

async function uploadFile(selector: string, fileName: string, type: string): Promise<void> {
  const input = document.body.querySelector<HTMLInputElement>(selector);

  expect(input).toBeTruthy();

  const file = new File(["conteudo"], fileName, { type });
  const files = [file] as unknown as FileList;
  Object.defineProperty(files, "item", {
    configurable: true,
    value: (index: number) => files[index] ?? null,
  });
  Object.defineProperty(input, "files", {
    configurable: true,
    value: files,
  });

  await act(async () => {
    input?.dispatchEvent(new Event("change", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
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

  for (let attempt = 0; !item && attempt < 10; attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
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

async function selectComboboxOption(label: string, optionLabel: string): Promise<void> {
  const dialog = document.body.querySelector<HTMLElement>('[role="dialog"]');
  const input = Array.from(
    dialog?.querySelectorAll<HTMLInputElement>("input") ?? [],
  ).find(
    (element) =>
      element.getAttribute("aria-label") === label ||
      element.getAttribute("role") === "combobox" ||
      element.closest(".mantine-Select-root")?.textContent?.includes(label),
  );

  expect(input).toBeTruthy();

  await act(async () => {
    input?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    input?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });

  let option = getComboboxOption(optionLabel);

  for (let attempt = 0; !option && attempt < 5; attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    option = getComboboxOption(optionLabel);
  }

  expect(option).toBeTruthy();

  await act(async () => {
    option?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function getComboboxOption(label: string): HTMLElement | undefined {
  return Array.from(
    document.body.querySelectorAll<HTMLElement>('[role="option"]'),
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

function getInputByLabel(label: string): HTMLInputElement {
  const input = Array.from(container.querySelectorAll<HTMLInputElement>("input")).find(
    (candidate) =>
      candidate.closest(".mantine-InputWrapper-root")?.textContent?.includes(label),
  );

  expect(input).toBeTruthy();
  return input as HTMLInputElement;
}

async function setFieldValue(
  field: HTMLInputElement | HTMLTextAreaElement,
  value: string,
): Promise<void> {
  const descriptor = Object.getOwnPropertyDescriptor(
    Object.getPrototypeOf(field),
    "value",
  );

  await act(async () => {
    descriptor?.set?.call(field, value);
    field.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
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

function getActiveOutlineItems(): HTMLElement[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>('.documentOutlineItem[data-active="true"]'),
  );
}

function createRect(top: number): DOMRect {
  return {
    x: 0,
    y: top,
    width: 100,
    height: 40,
    top,
    right: 100,
    bottom: top + 40,
    left: 0,
    toJSON: () => ({}),
  } as DOMRect;
}

function getReviewMetricValue(label: string): string | undefined {
  const metric = Array.from(container.querySelectorAll<HTMLElement>(".reviewMetric")).find(
    (element) => element.textContent?.includes(label),
  );

  return metric?.querySelectorAll("p")[1]?.textContent ?? undefined;
}

type Deferred<T> = {
  promise: Promise<T>;
  resolve: (value?: T | PromiseLike<T>) => void;
  reject: (reason?: unknown) => void;
};

function createDeferred<T>(): Deferred<T> {
  let resolve: Deferred<T>["resolve"] = () => undefined;
  let reject: Deferred<T>["reject"] = () => undefined;
  const promise = new Promise<T>((promiseResolve, promiseReject) => {
    resolve = promiseResolve as Deferred<T>["resolve"];
    reject = promiseReject;
  });

  return { promise, resolve, reject };
}

function createOtImportResult(): OtDocxImportResult {
  const documentData = createCompleteReviewDraft();

  return {
    kind: "ot",
    document: documentData,
    summary: {
      kind: "ot",
      screen: documentData.metadata.screen,
      accessSteps: documentData.accessSteps.length,
      permissionGroups: documentData.permissionGroups.length,
      selectedPermissions: 1,
      tests: 1,
      images: 2,
    },
    warnings: [],
    sourceName: "importado.docx",
  };
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

function createCorrectionDraft(): OtDocument {
  const documentData = createCompleteReviewDraft();
  const macro = documentData.permissionGroups[0];

  macro.microPermissions.push({
    id: "micro-co",
    code: "CO",
    label: "Consulta",
    selected: true,
  });

  documentData.permissionBlocks["macro-a:micro-at"] = {
    tests: [
      {
        id: "pending-correction",
        title: "pendente visivel",
        result: {
          ...createResult({ newIssue: true }),
          observations: "Falha pendente.",
          legacyImages: [createImage("legacy-pending")],
          newImages: [createImage("new-pending")],
        },
        correction: {
          corrected: false,
          hotfixTag: "",
          correctedBy: "",
          cloudStage: "none",
          beforeImages: [],
          afterImages: [],
        },
      },
    ],
  };

  documentData.permissionBlocks["macro-a:micro-co"] = {
    tests: [
      {
        id: "corrected-correction",
        title: "corrigido oculto",
        result: {
          ...createResult({ newIssue: true }),
          observations: "Falha corrigida.",
          legacyImages: [createImage("legacy-corrected")],
          newImages: [createImage("new-corrected")],
        },
        correction: {
          corrected: true,
          hotfixTag: "",
          correctedBy: "",
          cloudStage: "none",
          beforeImages: [createImage("before-corrected")],
          afterImages: [createImage("after-corrected")],
        },
      },
    ],
  };

  return documentData;
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

function createTeaDraftForSubActivityCopy(): TeaDocument {
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
        id: "copy-source-activity",
        title: "Atividade origem",
        blocks: [],
        subActivities: [
          {
            id: "copy-source-subactivity",
            title: "Subtopico origem",
            blocks: [
              {
                id: "copy-source-text",
                type: "text",
                text: "Texto copiado",
              },
              {
                id: "copy-source-list",
                type: "list",
                items: [{ id: "copy-source-item", text: "Item copiado" }],
              },
              {
                id: "copy-source-images",
                type: "images",
                images: [createImageWithoutData("copy-source-image", "Evidencia copiada")],
              },
            ],
          },
        ],
      },
      {
        id: "copy-target-a",
        title: "Atividade destino A",
        blocks: [],
        subActivities: [],
      },
      {
        id: "copy-target-b",
        title: "Atividade destino B",
        blocks: [],
        subActivities: [],
      },
    ],
  };
}

function createTeaDraftForBulkSubActivityCopy(): TeaDocument {
  const documentData = createTeaDraftForSubActivityCopy();
  const sourceActivity = documentData.activities[0];

  sourceActivity.subActivities.push({
    id: "copy-source-subactivity-two",
    title: "Segundo subtopico",
    blocks: [
      {
        id: "copy-source-two-text",
        type: "text",
        text: "Texto do segundo subtopico",
      },
    ],
  });

  return documentData;
}

function createSingleActivitySubtopicDraft(): TeaDocument {
  const documentData = createTeaDraftForSubActivityCopy();

  return {
    ...documentData,
    activities: documentData.activities.slice(0, 1),
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

function createImageWithoutData(id: string, label: string): EvidenceImage {
  const { dataUrl: _dataUrl, ...image } = createImage(id);

  return {
    ...image,
    label,
  };
}
