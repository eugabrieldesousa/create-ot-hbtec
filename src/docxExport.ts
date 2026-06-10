import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
} from "docx";
import {
  checkLabels,
  checkOrder,
  createEmptyTestCorrection,
  createPermissionKey,
} from "./defaultDocument";
import { hydrateDocumentImages, hydrateTeaDocumentImages } from "./imageStorage";
import { optimizeImageDataUrl } from "./imageOptimizer";
import type {
  EvidenceImage,
  OtDocument,
  PermissionBlock,
  PermissionBlockTest,
  PermissionGroup,
  PermissionItem,
  TeaActivity,
  TeaContentBlock,
  TeaDocument,
  TeaSubActivity,
  TestResult,
} from "./types";

type ExportableEvidenceImage = EvidenceImage & { dataUrl: string };
type TeaRunOptions = {
  bold?: boolean;
  italics?: boolean;
  font?: string;
  size?: number;
  underline?: boolean;
};

const PAGE_WIDTH_TWIPS = 12240;
const PAGE_HEIGHT_TWIPS = 15840;
const PAGE_MARGIN = {
  top: 1440,
  right: 1800,
  bottom: 1440,
  left: 1800,
};
const PAGE_CONTENT_WIDTH_TWIPS =
  PAGE_WIDTH_TWIPS - PAGE_MARGIN.left - PAGE_MARGIN.right;
const TEA_HEADER_TABLE_WIDTH_TWIPS = 8640;
const TEA_HEADER_COLUMN_WIDTHS = [2175, 6465];

const COLORS = {
  title: "111827",
  text: "1F2937",
  muted: "6B7280",
  border: "AEB7C2",
  borderSoft: "D8DEE6",
  sectionFill: "E7F0FA",
  labelFill: "F2F6FA",
  testFill: "F6FAFE",
  selectedFill: "EAF7EF",
  emptyFill: "F8FAFC",
};
const EXPORT_IMAGE_YIELD_INTERVAL = 3;

export async function exportOtDocument(documentData: OtDocument): Promise<void> {
  const hydratedDocument = await hydrateDocumentImages(documentData);
  const exportDocument = await optimizeOtDocumentImages(hydratedDocument);
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Arial",
            size: 22,
          },
          paragraph: {
            spacing: { after: 120 },
          },
        },
      },
      paragraphStyles: [
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            bold: true,
            size: 26,
            font: "Arial",
          },
          paragraph: {
            spacing: { before: 240, after: 120 },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: PAGE_WIDTH_TWIPS,
              height: PAGE_HEIGHT_TWIPS,
            },
            margin: PAGE_MARGIN,
          },
        },
        children: buildDocumentChildren(exportDocument),
      },
    ],
  });

  await downloadDocument(doc, createFileName(exportDocument));
}

export async function exportTeaDocument(documentData: TeaDocument): Promise<void> {
  const hydratedDocument = await hydrateTeaDocumentImages(documentData);
  const exportDocument = await optimizeTeaDocumentImages(hydratedDocument);
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Arial",
            size: 22,
          },
          paragraph: {
            spacing: { after: 200, line: 276 },
          },
        },
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            bold: true,
            size: 36,
            font: "Calibri",
          },
          paragraph: {
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          },
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            bold: true,
            italics: true,
            size: 28,
            font: "Calibri",
          },
          paragraph: {
            spacing: { before: 260, after: 120 },
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            bold: true,
            size: 28,
            font: "Calibri",
            underline: { type: UnderlineType.SINGLE },
          },
          paragraph: {
            spacing: { before: 220, after: 100 },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: PAGE_WIDTH_TWIPS,
              height: PAGE_HEIGHT_TWIPS,
            },
            margin: PAGE_MARGIN,
          },
        },
        children: buildTeaDocumentChildren(exportDocument),
      },
    ],
  });

  await downloadDocument(doc, createTeaFileName(exportDocument));
}

async function downloadDocument(doc: Document, fileName: string): Promise<void> {
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function optimizeOtDocumentImages(documentData: OtDocument): Promise<OtDocument> {
  return {
    ...documentData,
    permissionBlocks: Object.fromEntries(
      await Promise.all(
        Object.entries(documentData.permissionBlocks).map(async ([blockKey, block]) => [
          blockKey,
          {
            ...block,
            tests: await Promise.all(
              block.tests.map(async (test) => ({
                ...test,
                result: {
                  ...test.result,
                  legacyImages: await optimizeEvidenceImagesForExport(
                    test.result.legacyImages,
                  ),
                  newImages: await optimizeEvidenceImagesForExport(test.result.newImages),
                },
              })),
            ),
          },
        ]),
      ),
    ),
  };
}

async function optimizeTeaDocumentImages(documentData: TeaDocument): Promise<TeaDocument> {
  const optimizeTeaContentBlockImages = async (
    block: TeaContentBlock,
  ): Promise<TeaContentBlock> =>
    block.type === "images"
      ? {
          ...block,
          images: await optimizeEvidenceImagesForExport(block.images),
        }
      : block;

  return {
    ...documentData,
    activityImages: await optimizeEvidenceImagesForExport(documentData.activityImages),
    activities: await Promise.all(
      documentData.activities.map(async (activity) => ({
        ...activity,
        blocks: await Promise.all(activity.blocks.map(optimizeTeaContentBlockImages)),
        subActivities: await Promise.all(
          activity.subActivities.map(async (subActivity) => ({
            ...subActivity,
            blocks: await Promise.all(
              subActivity.blocks.map(optimizeTeaContentBlockImages),
            ),
          })),
        ),
      })),
    ),
  };
}

async function optimizeEvidenceImagesForExport(
  images: EvidenceImage[],
): Promise<EvidenceImage[]> {
  const optimizedImages: EvidenceImage[] = [];

  for (let index = 0; index < images.length; index += 1) {
    if (index > 0 && index % EXPORT_IMAGE_YIELD_INTERVAL === 0) {
      await yieldToBrowser();
    }

    optimizedImages.push(await optimizeEvidenceImageForExport(images[index]));
  }

  return optimizedImages;
}

async function optimizeEvidenceImageForExport(image: EvidenceImage): Promise<EvidenceImage> {
  if (!hasImageData(image) || image.optimized) {
    return image;
  }

  const optimized = await optimizeImageDataUrl(image.dataUrl);

  return {
    ...image,
    dataUrl: optimized.dataUrl,
    width: optimized.width,
    height: optimized.height,
    originalBytes: image.originalBytes ?? optimized.originalBytes,
    savedBytes: optimized.savedBytes,
    optimized: image.optimized || optimized.optimized,
  };
}

function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, 0);
  });
}

function buildTeaDocumentChildren(documentData: TeaDocument) {
  const children: (Paragraph | Table)[] = [
    teaTitle("Termo de Entrega de Atividade (TEA)"),
    teaMetadataTable(documentData),
    teaHeading1("1. VISÃO GERAL"),
    ...teaParagraphs(documentData.overview || " "),
    teaHeading1("2. ATIVIDADES REALIZADAS"),
    ...teaParagraphs(
      documentData.activityIntro ||
        "A seguir serao apresentadas, a nova interface e as suas funcionalidades:",
    ),
    ...teaEvidenceSection(documentData.activityImages),
  ];

  documentData.activities.forEach((activity, index) => {
    children.push(...teaActivitySection(activity, index + 1));
  });

  return children;
}

function teaActivitySection(activity: TeaActivity, index: number): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [
    teaHeading2(`2.${index} - ${activity.title || "Atividade"}:`),
    ...teaContentBlocks(activity.blocks),
  ];

  activity.subActivities.forEach((subActivity, subIndex) => {
    children.push(...teaSubActivitySection(subActivity, index, subIndex + 1));
  });

  return children;
}

function teaSubActivitySection(
  subActivity: TeaSubActivity,
  activityIndex: number,
  subIndex: number,
): (Paragraph | Table)[] {
  return [
    teaSubHeading(`2.${activityIndex}.${subIndex} - ${subActivity.title || "Subatividade"}:`),
    ...teaContentBlocks(subActivity.blocks),
  ];
}

function teaContentBlocks(blocks: TeaContentBlock[]): (Paragraph | Table)[] {
  return blocks.flatMap((block) => {
    if (block.type === "text") {
      return teaParagraphs(block.text);
    }

    if (block.type === "list") {
      return teaBullets(block.items);
    }

    return teaEvidenceSection(block.images);
  });
}

function teaMetadataTable(documentData: TeaDocument): Table {
  const rows = [
    ["Ordem de Serviço:", documentData.metadata.serviceOrder],
    ["Fase/Etapa", documentData.metadata.phase],
    ["Chamado:", documentData.metadata.ticket],
    ["Assunto:", documentData.metadata.subject],
    ["Data:", formatDisplayDate(documentData.metadata.date)],
    ["Elaborado por:", documentData.metadata.author],
  ];

  return new Table({
    width: { size: TEA_HEADER_TABLE_WIDTH_TWIPS, type: WidthType.DXA },
    columnWidths: TEA_HEADER_COLUMN_WIDTHS,
    layout: TableLayoutType.FIXED,
    borders: teaTableBorders(),
    rows: rows.map(
      ([label, value]) =>
        new TableRow({
          children: [
            teaCell(label, true, TEA_HEADER_COLUMN_WIDTHS[0]),
            teaCell(value || " ", false, TEA_HEADER_COLUMN_WIDTHS[1]),
          ],
        }),
    ),
  });
}

function teaTitle(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [teaRun(text, { bold: true, font: "Calibri", size: 36 })],
  });
}

function teaHeading1(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 120 },
    children: [
      teaRun(text, {
        bold: true,
        italics: true,
        font: "Calibri",
        size: 28,
        underline: true,
      }),
    ],
  });
}

function teaHeading2(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100 },
    children: [teaRun(text, { bold: true, font: "Calibri", size: 28, underline: true })],
  });
}

function teaSubHeading(text: string): Paragraph {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [teaRun(text, { bold: true })],
  });
}

function teaParagraphs(value: string): Paragraph[] {
  const paragraphs = value
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (paragraphs.length === 0) {
    return [];
  }

  return paragraphs.map(
    (text) =>
      new Paragraph({
        spacing: { after: 200, line: 276 },
        children: teaRunsFromMarkdown(text),
      }),
  );
}

function teaBullets(items: { text: string }[]): Paragraph[] {
  return items
    .map((item) => item.text.trim())
    .filter(Boolean)
    .map(
      (text) =>
        new Paragraph({
          bullet: { level: 0 },
          spacing: { after: 120, line: 276 },
          children: teaRunsFromMarkdown(text),
        }),
    );
}

function teaEvidenceSection(images: EvidenceImage[]): Paragraph[] {
  return images.flatMap((image) => {
    if (!hasImageData(image)) {
      return [];
    }

    const paragraphs: Paragraph[] = [];

    if (image.label.trim()) {
      paragraphs.push(
        new Paragraph({
          spacing: { before: 40, after: 80 },
          children: [teaRun(image.label.trim(), { bold: true })],
        }),
      );
    }

    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 180 },
        children: [createImageRun(image, { maxWidth: 576, maxHeight: 720 })],
      }),
    );

    return paragraphs;
  });
}

function buildDocumentChildren(documentData: OtDocument) {
  const selectedGroups = selectedPermissionGroups(documentData.permissionGroups);
  const children = [
    documentTitle("OBSERVABILIDADE DE TESTES (PERMISSÕES + NAVEGAÇÃO)"),
    labeledBox("Objetivo", documentData.objective || " "),
    metadataTable(documentData),
    sectionHeading("Passo a Passo para Acessar a Tela"),
    accessStepsTable(documentData),
    sectionHeading("Tipos de Permissão para Testar"),
    permissionSummaryTable(selectedGroups),
    sectionHeading("TESTES"),
  ];

  if (selectedGroups.length === 0) {
    children.push(emptyParagraph("Nenhuma permissão selecionada para gerar testes."));
  }

  for (const macro of selectedGroups) {
    for (const micro of macro.microPermissions) {
      children.push(...buildPermissionBlock(documentData, macro, micro));
    }
  }

  return children;
}

function accessStepsTable(documentData: OtDocument): Table {
  const steps = documentData.accessSteps.filter((step) => step.text.trim());

  if (steps.length === 0) {
    return simpleTable([[cell("Passo", true), cell("Nenhum passo informado.")]], [16, 84]);
  }

  return simpleTable(
    steps.map((step, index) => [
      cell(`${index + 1}.`, true, { alignment: AlignmentType.CENTER }),
      cell(step.text.trim()),
    ]),
    [10, 90],
  );
}

function permissionSummaryTable(groups: PermissionGroup[]): Table {
  if (groups.length === 0) {
    return simpleTable(
      [[cell("Permissões", true), cell("Nenhuma permissão selecionada.")]],
      [24, 76],
    );
  }

  return simpleTable(
    groups.flatMap((macro) => [
      [cell("Macro-permissão", true), cell(formatPermission(macro))],
      ...macro.microPermissions.map((micro) => [
        cell("Micro-permissão", true),
        cell(formatPermission(micro)),
      ]),
    ]),
    [28, 72],
  );
}

function metadataTable(documentData: OtDocument): Table {
  const rows = [
    ["Tela:", documentData.metadata.screen],
    ["Responsável pelo teste:", documentData.metadata.responsible],
    ["Data:", formatDisplayDate(documentData.metadata.date)],
    ["Ambiente:", documentData.metadata.environment],
    ["Elaborada por:", documentData.metadata.author],
  ];

  return simpleTable(
    rows.map(([label, value]) => [cell(label, true), cell(value || " ")]),
    [34, 66],
  );
}

function buildPermissionBlock(
  documentData: OtDocument,
  macro: PermissionGroup,
  micro: PermissionItem,
) {
  const key = createPermissionKey(macro.id, micro.id);
  const block = documentData.permissionBlocks[key] ?? createEmptyBlock();
  const children: (Paragraph | Table)[] = [
    permissionContextTable(macro, micro),
    spacerParagraph(),
  ];

  if (block.tests.length === 0) {
    children.push(emptyParagraph("Nenhum teste informado para esta permissão."));
    return children;
  }

  block.tests.forEach((test, index) => {
    children.push(testResultTable(index + 1, test));
    if (test.result.observations.trim()) {
      children.push(labeledBox("Observações", test.result.observations.trim()));
    }

    children.push(...evidenceSection("Legado:", test.result.legacyImages));
    children.push(...evidenceSection("Novo:", test.result.newImages));

    if (test.result.checks.newIssue) {
      children.push(...correctionSection(test));
    }
  });

  return children;
}

function permissionContextTable(macro: PermissionGroup, micro: PermissionItem): Table {
  return simpleTable(
    [
      [cell("MACRO-PERMISSÃO", true), cell(`Tipo de usuário: ${formatPermission(macro)}`)],
      [cell("MICRO-PERMISSÃO", true), cell(`Tipo de permissão: ${formatPermission(micro)}`)],
    ],
    [28, 72],
  );
}

function testResultTable(index: number, test: PermissionBlockTest): Table {
  return simpleTable(
    [
      [cell(`${index} - ${test.title || "Teste"}`, true, { columnSpan: 2, fill: COLORS.testFill })],
      ...checkOrder.map((key) => [
        cell(test.result.checks[key] ? "( X )" : "(   )", test.result.checks[key], {
          alignment: AlignmentType.CENTER,
          fill: test.result.checks[key] ? COLORS.selectedFill : undefined,
        }),
        cell(checkLabels[key], key === "bothIssue" && test.result.checks[key], {
          fill: test.result.checks[key] ? COLORS.selectedFill : undefined,
        }),
      ]),
    ],
    [14, 86],
  );
}

function correctionSection(test: PermissionBlockTest) {
  const correction = {
    ...createEmptyTestCorrection(),
    ...test.correction,
    beforeImages: test.correction?.beforeImages ?? [],
    afterImages: test.correction?.afterImages ?? [],
  };

  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [run("Correcao:", { bold: true, color: COLORS.title })],
    }),
    simpleTable(
      [
        [cell("Corrigido", true), cell(correction.corrected ? "Sim" : "Nao")],
        [cell("Hotfix", true), cell(correction.hotfixTag.trim() || " ")],
        [cell("Nuvem", true), cell(formatCloudStage(correction.cloudStage))],
      ],
      [28, 72],
    ),
    ...evidenceSection("Antes (com erro):", correction.beforeImages),
    ...evidenceSection("Depois (corrigido):", correction.afterImages),
  ];
}

function evidenceSection(label: string, images: EvidenceImage[]) {
  const children = [
    new Paragraph({
      spacing: { before: 80, after: 60 },
      children: [run(label, { bold: true, color: COLORS.title })],
    }),
  ];

  if (images.length === 0) {
    return children;
  }

  for (const image of images) {
    if (!hasImageData(image)) {
      children.push(emptyParagraph(`${image.name || "Imagem"} nao encontrada no rascunho.`));
      continue;
    }

    if (image.label.trim()) {
      children.push(
        new Paragraph({
          spacing: { before: 20, after: 60 },
          children: [run(image.label.trim(), { bold: true, color: COLORS.text })],
        }),
      );
    }

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 160 },
        children: [createImageRun(image)],
      }),
    );
  }

  return children;
}

function createImageRun(
  image: ExportableEvidenceImage,
  options: { maxWidth?: number; maxHeight?: number } = {},
): ImageRun {
  const maxWidth = options.maxWidth ?? 560;
  const maxHeight = options.maxHeight ?? 420;
  const width = image.width || maxWidth;
  const height = image.height || maxHeight;
  const scale = Math.min(maxWidth / width, maxHeight / height, 1);

  return new ImageRun({
    data: dataUrlToUint8Array(image.dataUrl),
    transformation: {
      width: Math.max(1, Math.round(width * scale)),
      height: Math.max(1, Math.round(height * scale)),
    },
    type: imageType(image.dataUrl),
  } as never);
}

function hasImageData(image: EvidenceImage): image is ExportableEvidenceImage {
  return typeof image.dataUrl === "string" && image.dataUrl.trim().length > 0;
}

function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = dataUrl.split(",")[1] ?? "";
  const binary = window.atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

function imageType(dataUrl: string): "png" | "jpg" | "gif" | "bmp" {
  if (dataUrl.startsWith("data:image/jpeg") || dataUrl.startsWith("data:image/jpg")) {
    return "jpg";
  }

  if (dataUrl.startsWith("data:image/gif")) {
    return "gif";
  }

  if (dataUrl.startsWith("data:image/bmp")) {
    return "bmp";
  }

  return "png";
}

type CellDefinition = {
  text: string;
  bold?: boolean;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  columnSpan?: number;
  fill?: string;
};

function documentTitle(text: string): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 220 },
    children: [run(text, { bold: true, italics: true, color: COLORS.title, size: 28 })],
  });
}

function sectionHeading(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 120 },
    shading: { type: ShadingType.CLEAR, fill: COLORS.sectionFill },
    children: [run(text, { bold: true, italics: true, color: COLORS.title })],
  });
}

function labeledBox(label: string, value: string): Table {
  return simpleTable([[cell(`${label}:`, true), cell(value || " ")]], [22, 78]);
}

function emptyParagraph(text: string): Paragraph {
  return new Paragraph({
    spacing: { after: 120 },
    shading: { type: ShadingType.CLEAR, fill: COLORS.emptyFill },
    children: [run(text, { italics: true, color: COLORS.muted })],
  });
}

function spacerParagraph(): Paragraph {
  return new Paragraph({
    spacing: { after: 80 },
    children: [run(" ", { size: 2 })],
  });
}

function simpleTable(rows: CellDefinition[][], columnWidths: number[]): Table {
  const columnWidthTwips = toTwipColumnWidths(columnWidths);

  return new Table({
    width: { size: PAGE_CONTENT_WIDTH_TWIPS, type: WidthType.DXA },
    columnWidths: columnWidthTwips,
    layout: TableLayoutType.FIXED,
    borders: tableBorders(),
    rows: rows.map(
      (row) => {
        let columnIndex = 0;

        return new TableRow({
          children: row.map((definition) => {
            const columnSpan = definition.columnSpan ?? 1;
            const width = columnWidthTwips
              .slice(columnIndex, columnIndex + columnSpan)
              .reduce((total, value) => total + value, 0);
            const fill = definition.fill ?? (definition.bold ? COLORS.labelFill : undefined);
            columnIndex += columnSpan;

            return new TableCell({
              width: { size: width || PAGE_CONTENT_WIDTH_TWIPS, type: WidthType.DXA },
              columnSpan: definition.columnSpan,
              margins: cellMargins(),
              verticalAlign: VerticalAlign.CENTER,
              shading: fill
                ? { type: ShadingType.CLEAR, fill, color: "auto" }
                : undefined,
              children: [
                ...cellParagraphTexts(definition.text).map(
                  (paragraphText) =>
                    new Paragraph({
                      alignment: definition.alignment ?? AlignmentType.LEFT,
                      spacing: { before: 0, after: 0 },
                      children: [
                        run(paragraphText || " ", {
                          bold: definition.bold,
                          color: definition.bold ? COLORS.title : COLORS.text,
                        }),
                      ],
                    }),
                ),
              ],
            });
          }),
        });
      },
    ),
  });
}

function cellParagraphTexts(text: string): string[] {
  const paragraphs = (text || " ").replace(/\r\n/g, "\n").split("\n");

  return paragraphs.length > 0 ? paragraphs : [" "];
}

function toTwipColumnWidths(widths: number[]): number[] {
  const total = widths.reduce((sum, width) => sum + width, 0) || 100;

  return widths.map((width) =>
    Math.max(1, Math.round((PAGE_CONTENT_WIDTH_TWIPS * width) / total)),
  );
}

function cell(
  text: string,
  bold = false,
  options: Omit<CellDefinition, "text" | "bold"> = {},
): CellDefinition {
  return {
    text,
    bold,
    ...options,
  };
}

function run(
  text: string,
  options: { bold?: boolean; italics?: boolean; color?: string; size?: number } = {},
): TextRun {
  return new TextRun({
    text,
    bold: options.bold,
    italics: options.italics,
    color: options.color ?? COLORS.text,
    font: "Arial",
    size: options.size ?? 22,
  });
}

function tableBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
    left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
    right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.borderSoft },
    insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.borderSoft },
  };
}

function teaCell(text: string, bold: boolean, width: number): TableCell {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins(),
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        spacing: { before: 0, after: 0 },
        children: [teaRun(text || " ", { bold })],
      }),
    ],
  });
}

function teaRun(
  text: string,
  options: TeaRunOptions = {},
): TextRun {
  return new TextRun({
    text,
    bold: options.bold,
    italics: options.italics,
    underline: options.underline ? { type: UnderlineType.SINGLE } : undefined,
    font: options.font ?? "Arial",
    size: options.size ?? 22,
  });
}

function teaRunsFromMarkdown(text: string, options: TeaRunOptions = {}): TextRun[] {
  const markerMatches = text.match(/\*\*/g) ?? [];

  if (markerMatches.length === 0 || markerMatches.length % 2 !== 0) {
    return [teaRun(text, options)];
  }

  const runs: TextRun[] = [];
  let cursor = 0;
  let bold = false;

  while (cursor < text.length) {
    const markerIndex = text.indexOf("**", cursor);

    if (markerIndex === -1) {
      const value = text.slice(cursor);
      if (value) {
        runs.push(teaRun(value, { ...options, bold: options.bold || bold }));
      }
      break;
    }

    const value = text.slice(cursor, markerIndex);
    if (value) {
      runs.push(teaRun(value, { ...options, bold: options.bold || bold }));
    }

    bold = !bold;
    cursor = markerIndex + 2;
  }

  return runs.length > 0 ? runs : [teaRun(text.replace(/\*\*/g, ""), options)];
}

function teaTableBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  };
}

function cellMargins() {
  return {
    top: 110,
    bottom: 110,
    left: 160,
    right: 160,
  };
}

function selectedPermissionGroups(groups: PermissionGroup[]): PermissionGroup[] {
  return groups
    .filter((macro) => macro.selected && macro.code.trim())
    .map((macro) => ({
      ...macro,
      microPermissions: selectedPermissions(macro.microPermissions),
    }))
    .filter((macro) => macro.microPermissions.length > 0);
}

function selectedPermissions(permissions: PermissionItem[]): PermissionItem[] {
  return permissions.filter((permission) => permission.selected && permission.code.trim());
}

function formatPermission(permission: PermissionItem): string {
  const code = permission.code.trim();
  const label = permission.label.trim();

  if (code && label) {
    return `${code} (${label})`;
  }

  return code || label || "Sem código";
}

function formatCloudStage(value: string): string {
  if (value === "dev") {
    return "Ate dev";
  }

  if (value === "homolog") {
    return "Ate homolog";
  }

  if (value === "production") {
    return "Ate producao";
  }

  return "Nao enviado";
}

function createEmptyBlock(): PermissionBlock {
  return { tests: [] };
}

function formatDisplayDate(value: string): string {
  if (!value) {
    return "";
  }

  const [year, month, day] = value.split("-");
  return year && month && day ? `${day}/${month}/${year}` : value;
}

function createFileName(documentData: OtDocument): string {
  const screen = sanitizeFileName(documentData.metadata.screen || "Documento");
  const date = documentData.metadata.date || new Date().toISOString().slice(0, 10);
  return `OT - ${screen} - ${date}.docx`;
}

function createTeaFileName(documentData: TeaDocument): string {
  const base = sanitizeFileName(
    documentData.metadata.subject || documentData.metadata.serviceOrder || "Documento",
  );
  const date = documentData.metadata.date || new Date().toISOString().slice(0, 10);
  return `TEA - ${base} - ${date}.docx`;
}

function sanitizeFileName(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
