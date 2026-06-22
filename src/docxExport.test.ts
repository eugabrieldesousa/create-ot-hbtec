import { beforeEach, describe, expect, it, vi } from "vitest";
import { exportOtDocument, exportTeaDocument } from "./docxExport";
import type { CheckKey, EvidenceImage, OtDocument, TeaDocument } from "./types";

const docxState = vi.hoisted(() => ({
  documents: [] as Array<{ options: { sections: Array<{ children: unknown[] }> } }>,
  downloads: [] as string[],
}));
const jszipState = vi.hoisted(() => ({
  media: [new Uint8Array([0, 0, 0])] as Uint8Array[],
  loadAsync: vi.fn(),
}));

vi.mock("docx", () => {
  class Document {
    options: { sections: Array<{ children: unknown[] }> };

    constructor(options: { sections: Array<{ children: unknown[] }> }) {
      this.options = options;
      docxState.documents.push(this);
    }
  }

  class Paragraph {
    options: { children?: unknown[] };

    constructor(options: { children?: unknown[] }) {
      this.options = options;
    }
  }

  class TextRun {
    options: { bold?: boolean; size?: number; text?: string; underline?: unknown };

    constructor(options: { bold?: boolean; size?: number; text?: string; underline?: unknown }) {
      this.options = options;
    }
  }

  class ImageRun {
    options: unknown;

    constructor(options: unknown) {
      this.options = options;
    }
  }

  class Table {
    options: unknown;

    constructor(options: unknown) {
      this.options = options;
    }
  }

  class TableCell {
    options: unknown;

    constructor(options: unknown) {
      this.options = options;
    }
  }

  class TableRow {
    options: unknown;

    constructor(options: unknown) {
      this.options = options;
    }
  }

  return {
    AlignmentType: { CENTER: "center" },
    BorderStyle: { SINGLE: "single" },
    Document,
    HeadingLevel: {
      HEADING_1: "Heading1",
      HEADING_2: "Heading2",
      TITLE: "Title",
    },
    ImageRun,
    Packer: {
      toBlob: vi.fn(async () => new Blob(["docx"])),
    },
    Paragraph,
    ShadingType: { CLEAR: "clear" },
    Table,
    TableCell,
    TableLayoutType: { FIXED: "fixed" },
    TableRow,
    TextRun,
    UnderlineType: { SINGLE: "single" },
    VerticalAlign: { CENTER: "center" },
    WidthType: { DXA: "dxa", PERCENTAGE: "percentage" },
  };
});

vi.mock("jszip", () => {
  class JSZip {
    static async loadAsync(): Promise<{
      files: Record<string, { dir: boolean; async: () => Promise<Uint8Array> }>;
    }> {
      jszipState.loadAsync();

      return {
        files: Object.fromEntries(
          jszipState.media.map((bytes, index) => [
            `word/media/image-${index + 1}.png`,
            {
              dir: false,
              async: vi.fn(async () => bytes),
            },
          ]),
        ),
      };
    }
  }

  return { default: JSZip };
});

vi.mock("./imageStorage", () => ({
  hydrateDocumentImages: vi.fn(async (documentData) => documentData),
  hydrateTeaDocumentImages: vi.fn(async (documentData) => documentData),
}));

vi.mock("./imageOptimizer", () => ({
  optimizeImageDataUrl: vi.fn(async (dataUrl: string) => ({
    dataUrl,
    width: 100,
    height: 80,
    originalBytes: 4,
    savedBytes: 4,
    optimized: false,
  })),
}));

beforeEach(() => {
  docxState.documents.length = 0;
  docxState.downloads.length = 0;
  jszipState.media = [new Uint8Array([0, 0, 0])];
  jszipState.loadAsync.mockClear();

  Object.defineProperty(URL, "createObjectURL", {
    configurable: true,
    value: vi.fn(() => "blob:test"),
  });

  Object.defineProperty(URL, "revokeObjectURL", {
    configurable: true,
    value: vi.fn(),
  });

  Object.defineProperty(HTMLAnchorElement.prototype, "click", {
    configurable: true,
    value: vi.fn(function (this: HTMLAnchorElement) {
      docxState.downloads.push(this.download);
    }),
  });
});

describe("exportTeaDocument", () => {
  it("renders activity content blocks in the configured order", async () => {
    await exportTeaDocument(createTeaDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const texts = children.map(readParagraphText).filter(Boolean);

    expect(texts.indexOf("Legenda antes")).toBeLessThan(texts.indexOf("Item entre"));
    expect(texts.indexOf("Item entre")).toBeLessThan(texts.indexOf("Legenda depois"));
    expect(texts.indexOf("Legenda depois")).toBeLessThan(texts.indexOf("Texto final"));
  });

  it("underlines the main TEA headings", async () => {
    await exportTeaDocument(createTeaDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const headingRuns = children
      .filter((child) => {
        const text = readParagraphText(child);
        return text.startsWith("1.") || text === "2. ATIVIDADES REALIZADAS";
      })
      .flatMap(readParagraphRuns);

    expect(headingRuns).toHaveLength(2);
    expect(headingRuns.every((run) => run.options.underline)).toBe(true);
  });

  it("exports activity headings with larger underlined text", async () => {
    await exportTeaDocument(createTeaDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const activityHeading = children.find((child) =>
      readParagraphText(child).startsWith("2.1 - Atividade:"),
    );

    expect(readParagraphRuns(activityHeading)).toEqual([
      expect.objectContaining({
        options: expect.objectContaining({
          text: "2.1 - Atividade:",
          size: 28,
          underline: expect.anything(),
        }),
      }),
    ]);
  });

  it("renders double-asterisk text segments as bold runs", async () => {
    await exportTeaDocument(createTeaDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const textParagraph = children.find((child) => readParagraphText(child) === "Texto final");
    const listParagraph = children.find((child) => readParagraphText(child) === "Item entre");

    expect(readParagraphRuns(textParagraph)).toEqual([
      expect.objectContaining({ options: expect.objectContaining({ text: "Texto " }) }),
      expect.objectContaining({
        options: expect.objectContaining({ text: "final", bold: true }),
      }),
    ]);
    expect(readParagraphRuns(listParagraph)).toEqual([
      expect.objectContaining({ options: expect.objectContaining({ text: "Item " }) }),
      expect.objectContaining({
        options: expect.objectContaining({ text: "entre", bold: true }),
      }),
    ]);
  });

  it("blocks TEA export when an image block is empty", async () => {
    const documentData = createTeaDocumentForExport();
    const imageBlock = documentData.activities[0].blocks[0];

    if (imageBlock.type === "images") {
      imageBlock.images = [];
    }

    await expect(exportTeaDocument(documentData)).rejects.toMatchObject({
      name: "DocxExportImageError",
      problems: [
        expect.objectContaining({
          label: "Bloco de imagens vazio",
          documentKind: "tea",
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });

  it("blocks TEA export when an image has no data", async () => {
    const documentData = createTeaDocumentForExport();
    const imageBlock = documentData.activities[0].blocks[0];

    if (imageBlock.type === "images") {
      imageBlock.images = [createExportImageWithoutData("missing-tea", "Sem dados")];
    }

    await expect(exportTeaDocument(documentData)).rejects.toMatchObject({
      name: "DocxExportImageError",
      problems: [
        expect.objectContaining({
          label: "Imagem sem dados",
          documentKind: "tea",
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });

  it("reports invalid image data URLs with a specific message", async () => {
    const documentData = createTeaDocumentForExport();
    const imageBlock = documentData.activities[0].blocks[0];

    if (imageBlock.type === "images") {
      imageBlock.images = [
        {
          ...createExportImage("webp-image", "WebP"),
          dataUrl: "data:image/webp;base64,AAAA",
        },
      ];
    }

    await expect(exportTeaDocument(documentData)).rejects.toMatchObject({
      problems: [
        expect.objectContaining({
          label: "Imagem invalida",
          detail: expect.stringContaining("tipo de imagem nao suportado"),
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });

  it("reports when the generated DOCX is missing expected image media", async () => {
    jszipState.media = [];

    await expect(exportTeaDocument(createTeaDocumentForExport())).rejects.toMatchObject({
      problems: [
        expect.objectContaining({
          label: "Imagem nao encontrada no DOCX gerado",
          detail: expect.stringContaining("word/media"),
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });

  it("allows duplicated image content when the DOCX stores one deduplicated media file", async () => {
    const documentData = createTeaDocumentForExport();
    const imageBlock = documentData.activities[0].blocks[0];

    if (imageBlock.type === "images") {
      imageBlock.images = [
        createExportImage("duplicate-a", "Duplicada A"),
        createExportImage("duplicate-b", "Duplicada B"),
      ];
    }

    await exportTeaDocument(documentData);

    expect(jszipState.loadAsync).toHaveBeenCalledTimes(1);
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);
  });
});

describe("exportOtDocument", () => {
  it("preserves accents in exported OT file names", async () => {
    const documentData = createOtDocumentForExport();
    documentData.metadata.screen = "Situação do Beneficiário";

    await exportOtDocument(documentData);

    expect(docxState.downloads).toContain("OT - Situação do Beneficiário - 2026-06-08.docx");
  });

  it("exports current check labels without deriving error report from Novo", async () => {
    await exportOtDocument(createOtDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const testTable = children.find((child) =>
      readTableText(child).includes("Teste com observação"),
    );
    const rows = readTableRowsText(testTable);

    expect(rows).toContain("(   ) Erro no legado");
    expect(rows).toContain("( X ) Erro no novo");
    expect(rows).toContain("(   ) Relatório de Erros");
  });

  it("derives the error report row from Legado", async () => {
    const documentData = createOtDocumentForExport();
    const checks = documentData.permissionBlocks["macro:micro"].tests[0].result.checks;
    checks.sameBehavior = false;
    checks.bothIssue = true;
    checks.newIssue = false;
    checks.errorReport = false;

    await exportOtDocument(documentData);

    const children = docxState.documents[0].options.sections[0].children;
    const testTable = children.find((child) =>
      readTableText(child).includes("Teste com observação"),
    );
    const rows = readTableRowsText(testTable);

    expect(rows).toContain("( X ) Erro no legado");
    expect(rows).toContain("(   ) Erro no novo");
    expect(rows).toContain("( X ) Relatório de Erros");
  });

  it("preserves line breaks in test observations", async () => {
    await exportOtDocument(createOtDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const observationsTable = children.find((child) =>
      readTableText(child).includes("Observações:"),
    );

    expect(readSecondCellParagraphTexts(observationsTable)).toEqual([
      "Primeira linha",
      "Segunda linha",
      "Terceira linha",
    ]);
  });

  it("exports correction details for Novo tests", async () => {
    await exportOtDocument(createOtDocumentForExport());

    const children = docxState.documents[0].options.sections[0].children;
    const correctionTable = children.find((child) =>
      readTableText(child).includes("Corrigido"),
    );
    const paragraphTexts = children.map(readParagraphText).filter(Boolean);
    const correctionText = readFullTableText(correctionTable);

    expect(correctionText).toContain("Corrigido Sim");
    expect(correctionText).toContain("Hotfix hotfix 1.2.2");
    expect(correctionText).toContain("Corrigido por Gabriel");
    expect(correctionText).toContain("Nuvem Ate dev");
    expect(paragraphTexts).toContain("Antes (com erro):");
    expect(paragraphTexts).toContain("Depois (corrigido):");
  });

  it("blocks OT export when Legado or Novo images are empty", async () => {
    const documentData = createOtDocumentForExport();
    const test = documentData.permissionBlocks["macro:micro"].tests[0];
    test.result.legacyImages = [];

    await expect(exportOtDocument(documentData)).rejects.toMatchObject({
      problems: [
        expect.objectContaining({
          label: "Imagem do legado ausente",
          documentKind: "ot",
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });

  it("blocks OT export when an attached image has no data", async () => {
    const documentData = createOtDocumentForExport();
    const test = documentData.permissionBlocks["macro:micro"].tests[0];
    test.result.newImages = [createExportImageWithoutData("missing-ot", "Sem dados")];

    await expect(exportOtDocument(documentData)).rejects.toMatchObject({
      problems: [
        expect.objectContaining({
          label: "Imagem sem dados",
          documentKind: "ot",
        }),
      ],
    });
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();
  });
});

function readParagraphText(node: unknown): string {
  const children = (node as { options?: { children?: unknown[] } }).options?.children ?? [];

  return children
    .map((child) => (child as { options?: { text?: string } }).options?.text ?? "")
    .join("");
}

function readParagraphRuns(
  node: unknown,
): Array<{ options: { bold?: boolean; size?: number; text?: string; underline?: unknown } }> {
  return (
    (node as { options?: { children?: Array<{ options: { bold?: boolean; size?: number; text?: string; underline?: unknown } }> } })
      .options?.children ?? []
  );
}

function readTableText(node: unknown): string {
  return readTableCells(node)
    .flatMap((cell) => readCellParagraphs(cell).map(readParagraphText))
    .join(" ");
}

function readFullTableText(node: unknown): string {
  const rows = (node as { options?: { rows?: Array<{ options?: { children?: unknown[] } }> } })
    .options?.rows ?? [];

  return rows
    .flatMap((row) => row.options?.children ?? [])
    .flatMap((cell) => readCellParagraphs(cell).map(readParagraphText))
    .join(" ");
}

function readTableRowsText(node: unknown): string[] {
  const rows = (node as { options?: { rows?: Array<{ options?: { children?: unknown[] } }> } })
    .options?.rows ?? [];

  return rows.map((row) =>
    (row.options?.children ?? [])
      .flatMap((cell) => readCellParagraphs(cell).map(readParagraphText))
      .join(" "),
  );
}

function readSecondCellParagraphTexts(node: unknown): string[] {
  const secondCell = readTableCells(node)[1];

  return readCellParagraphs(secondCell).map(readParagraphText);
}

function readTableCells(node: unknown): unknown[] {
  const rows = (node as { options?: { rows?: Array<{ options?: { children?: unknown[] } }> } })
    .options?.rows ?? [];

  return rows[0]?.options?.children ?? [];
}

function readCellParagraphs(node: unknown): unknown[] {
  return (node as { options?: { children?: unknown[] } }).options?.children ?? [];
}

function createTeaDocumentForExport(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas - Novo Layout",
      date: "2026-06-08",
      author: "Gabriel Sousa",
    },
    overview: "Visao geral.",
    activityIntro: "Atividades realizadas:",
    activityImages: [],
    activities: [
      {
        id: "activity",
        title: "Atividade",
        blocks: [
          {
            id: "image-before",
            type: "images",
            images: [createExportImage("image-before", "Legenda antes")],
          },
          {
            id: "list-between",
            type: "list",
            items: [{ id: "item-between", text: "Item **entre**" }],
          },
          {
            id: "image-after",
            type: "images",
            images: [createExportImage("image-after", "Legenda depois")],
          },
          {
            id: "text-final",
            type: "text",
            text: "Texto **final**",
          },
        ],
        subActivities: [],
      },
    ],
  };
}

function createOtDocumentForExport(): OtDocument {
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
      date: "2026-06-08",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar observações.",
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
            title: "Teste com observação",
            result: {
              checks,
              observations: "Primeira linha\nSegunda linha\nTerceira linha",
              legacyImages: [createExportImage("legacy", "Legado")],
              newImages: [createExportImage("new", "Novo")],
            },
            correction: {
              corrected: true,
              hotfixTag: "hotfix 1.2.2",
              correctedBy: "Gabriel",
              cloudStage: "dev",
              beforeImages: [createExportImage("before-fix", "Antes")],
              afterImages: [createExportImage("after-fix", "Depois")],
            },
          },
        ],
      },
    },
  };
}

function createExportImage(id: string, label: string) {
  return {
    id,
    label,
    name: `${id}.png`,
    dataUrl: "data:image/png;base64,AAAA",
    width: 100,
    height: 80,
  };
}

function createExportImageWithoutData(id: string, label: string): EvidenceImage {
  const { dataUrl: _dataUrl, ...image } = createExportImage(id, label);
  return image;
}
