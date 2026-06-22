import { describe, expect, it } from "vitest";
import {
  buildOtPreviewModel,
  buildTeaPreviewModel,
  type DocxPreviewBlock,
  type DocxPreviewCell,
  type DocxPreviewRun,
} from "./docxPreviewModel";
import type { CheckKey, EvidenceImage, OtDocument, TeaDocument } from "./types";

describe("buildOtPreviewModel", () => {
  it("mirrors the exported OT section order, tables, checks, observations and missing images", () => {
    const model = buildOtPreviewModel(createOtPreviewDocument());
    const texts = model.blocks.map(blockText);

    expect(model).toMatchObject({
      kind: "ot",
      sectionId: "ot-section-preview",
    });
    expect(texts[0]).toContain("OBSERVABILIDADE DE TESTES");
    expect(indexOfText(texts, "Passo a Passo")).toBeLessThan(
      indexOfText(texts, "Tipos de Permissão"),
    );
    expect(indexOfText(texts, "Tipos de Permissão")).toBeLessThan(
      indexOfExactText(texts, "TESTES"),
    );

    expect(texts.some((text) => text.includes("Tela:") && text.includes("Documentos"))).toBe(
      true,
    );
    expect(texts.some((text) => text.includes("1.") && text.includes("Acessar"))).toBe(true);
    expect(texts.some((text) => text.includes("AO (Administrador)"))).toBe(true);
    expect(texts.some((text) => text.includes("AT (Atualização)"))).toBe(true);
    expect(texts.some((text) => text.includes("1 - Filtro") && text.includes("( X )"))).toBe(
      true,
    );
    expect(texts.some((text) => text.includes("Correcao:"))).toBe(true);
    expect(texts.some((text) => text.includes("Corrigido") && text.includes("Sim"))).toBe(true);
    expect(texts.some((text) => text.includes("Hotfix") && text.includes("hotfix 1.2.2"))).toBe(
      true,
    );
    expect(texts.some((text) => text.includes("Corrigido por") && text.includes("Gabriel"))).toBe(
      true,
    );
    expect(texts.some((text) => text.includes("Nuvem") && text.includes("Ate homolog"))).toBe(
      true,
    );
    expect(texts.some((text) => text.includes("legado.png nao encontrada"))).toBe(true);

    const testTable = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "table" }> =>
        block.type === "table" && blockText(block).includes("1 - Filtro"),
    );
    const testTableText = testTable ? blockText(testTable) : "";

    expect(testTable?.anchorId).toBe("ot-preview-test-macro-micro--test");
    expect(testTableText).toContain("Erro no legado");
    expect(testTableText).toContain("( X ) Erro no novo");
    expect(testTableText).toContain("(   ) Relatório de Erros");

    const observationsTable = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "table" }> =>
        block.type === "table" && blockText(block).includes("Observações:"),
    );
    expect(observationsTable?.rows[0][1].paragraphs.map(runsText)).toEqual([
      "Primeira linha",
      "Segunda linha",
    ]);

    const image = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "image" }> =>
        block.type === "image",
    );
    expect(image).toMatchObject({ width: 560, height: 280, alt: "Novo ok" });
  });

  it("derives the OT preview error report from Legado", () => {
    const documentData = createOtPreviewDocument();
    const checks = documentData.permissionBlocks["macro:micro"].tests[0].result.checks;
    checks.sameBehavior = false;
    checks.bothIssue = true;
    checks.newIssue = false;
    checks.errorReport = false;

    const model = buildOtPreviewModel(documentData);
    const testTable = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "table" }> =>
        block.type === "table" && blockText(block).includes("1 - Filtro"),
    );
    const testTableText = testTable ? blockText(testTable) : "";

    expect(testTableText).toContain("( X ) Erro no legado");
    expect(testTableText).toContain("(   ) Erro no novo");
    expect(testTableText).toContain("( X ) Relatório de Erros");
  });
});

describe("buildTeaPreviewModel", () => {
  it("mirrors TEA metadata, numbered sections, block order, bold markdown and image scaling", () => {
    const model = buildTeaPreviewModel(createTeaPreviewDocument());
    const texts = model.blocks.map(blockText);

    expect(model).toMatchObject({
      kind: "tea",
      sectionId: "tea-section-preview",
    });
    expect(texts[0]).toContain("Termo de Entrega de Atividade");
    expect(indexOfText(texts, "1. VISÃO GERAL")).toBeLessThan(
      indexOfText(texts, "2. ATIVIDADES REALIZADAS"),
    );
    expect(texts.some((text) => text.includes("Ordem de Serviço:") && text.includes("OS2171"))).toBe(
      true,
    );

    expect(indexOfText(texts, "Legenda antes")).toBeLessThan(indexOfText(texts, "Item entre"));
    expect(indexOfText(texts, "Item entre")).toBeLessThan(indexOfText(texts, "Texto final"));
    expect(texts.some((text) => text.includes("2.1 - Atividade:"))).toBe(true);
    expect(texts.some((text) => text.includes("2.1.1 - Modal:"))).toBe(true);

    const boldList = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "list" }> => block.type === "list",
    );
    expect(boldList?.items[0]).toEqual([
      { text: "Item ", bold: false },
      { text: "entre", bold: true },
    ]);

    const finalText = model.blocks.find(
      (block): block is Extract<DocxPreviewBlock, { type: "paragraph" }> =>
        block.type === "paragraph" && blockText(block) === "Texto final",
    );
    expect(finalText?.runs).toEqual([
      { text: "Texto ", bold: false },
      { text: "final", bold: true },
    ]);

    const images = model.blocks.filter(
      (block): block is Extract<DocxPreviewBlock, { type: "image" }> =>
        block.type === "image",
    );
    expect(images[0]).toMatchObject({ width: 576, height: 192, alt: "Imagem geral" });
  });
});

function blockText(block: DocxPreviewBlock): string {
  if (block.type === "paragraph") {
    return runsText(block.runs);
  }

  if (block.type === "table") {
    return block.rows.flatMap((row) => row.map(cellText)).join(" ");
  }

  if (block.type === "list") {
    return block.items.map(runsText).join(" ");
  }

  return block.label || block.alt;
}

function cellText(cell: DocxPreviewCell): string {
  return cell.paragraphs.map(runsText).join(" ");
}

function runsText(runs: DocxPreviewRun[]): string {
  return runs.map((run) => run.text).join("");
}

function indexOfText(texts: string[], expected: string): number {
  const index = texts.findIndex((text) => text.includes(expected));

  expect(index, `Expected to find "${expected}"`).toBeGreaterThanOrEqual(0);

  return index;
}

function indexOfExactText(texts: string[], expected: string): number {
  const index = texts.findIndex((text) => text === expected);

  expect(index, `Expected to find exact "${expected}"`).toBeGreaterThanOrEqual(0);

  return index;
}

function createOtPreviewDocument(): OtDocument {
  const checks = {
    sameBehavior: true,
    possibleIssue: false,
    bothIssue: false,
    newIssue: true,
    errorReport: false,
  } satisfies Record<CheckKey, boolean>;

  return {
    metadata: {
      screen: "Documentos Vencidos",
      responsible: "GABRIEL",
      date: "2026-06-09",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar documento.",
    accessSteps: [{ id: "step-1", text: "Acessar a tela" }],
    permissionGroups: [
      {
        id: "macro",
        code: "AO",
        label: "Administrador",
        selected: true,
        microPermissions: [
          {
            id: "micro",
            code: "AT",
            label: "Atualização",
            selected: true,
          },
        ],
      },
    ],
    permissionBlocks: {
      "macro:micro": {
        tests: [
          {
            id: "test",
            title: "Filtro",
            result: {
              checks,
              observations: "Primeira linha\nSegunda linha",
              legacyImages: [
                {
                  id: "missing",
                  label: "",
                  name: "legado.png",
                  width: 100,
                  height: 80,
                },
              ],
              newImages: [createImage("new", "Novo ok", 1000, 500)],
              errors: [],
            },
            correction: {
              corrected: true,
              hotfixTag: "hotfix 1.2.2",
              correctedBy: "Gabriel",
              cloudStage: "homolog",
              beforeImages: [createImage("before-fix", "Antes com erro", 100, 80)],
              afterImages: [createImage("after-fix", "Depois corrigido", 100, 80)],
            },
          },
        ],
      },
    },
  };
}

function createTeaPreviewDocument(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-04-28",
      author: "Gabriel Sousa",
    },
    overview: "Visão geral.",
    activityIntro: "Atividades realizadas:",
    activityImages: [createImage("general", "Imagem geral", 1200, 400)],
    activities: [
      {
        id: "activity",
        title: "Atividade",
        blocks: [
          {
            id: "image-before",
            type: "images",
            images: [createImage("before", "Legenda antes", 100, 80)],
          },
          {
            id: "list",
            type: "list",
            items: [{ id: "item", text: "Item **entre**" }],
          },
          {
            id: "text",
            type: "text",
            text: "Texto **final**",
          },
        ],
        subActivities: [
          {
            id: "sub",
            title: "Modal",
            blocks: [{ id: "sub-text", type: "text", text: "Texto da subatividade." }],
          },
        ],
      },
    ],
  };
}

function createImage(
  id: string,
  label: string,
  width: number,
  height: number,
): EvidenceImage {
  return {
    id,
    label,
    name: `${id}.png`,
    dataUrl: "data:image/png;base64,AAAA",
    width,
    height,
  };
}
