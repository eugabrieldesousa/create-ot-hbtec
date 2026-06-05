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
  VerticalAlign,
  WidthType,
} from "docx";
import { checkLabels, checkOrder, createPermissionKey } from "./defaultDocument";
import type {
  EvidenceImage,
  OtDocument,
  PermissionBlock,
  PermissionBlockTest,
  PermissionGroup,
  PermissionItem,
  TestResult,
} from "./types";

const PAGE_WIDTH_TWIPS = 12240;
const PAGE_HEIGHT_TWIPS = 15840;
const PAGE_MARGIN = {
  top: 1440,
  right: 1800,
  bottom: 1440,
  left: 1800,
};

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

export async function exportOtDocument(documentData: OtDocument): Promise<void> {
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
        children: buildDocumentChildren(documentData),
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download = createFileName(documentData);
  anchor.click();
  URL.revokeObjectURL(url);
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
    if (test.mode === "idem") {
      children.push(
        idemTestTable(
          index + 1,
          test,
          getPermissionTestLabel(documentData, test.idemReferenceKey) ||
            "referência não informada",
        ),
      );
      return;
    }

    children.push(testResultTable(index + 1, test));
    if (test.result.observations.trim()) {
      children.push(labeledBox("Observações", test.result.observations.trim()));
    }

    children.push(...evidenceSection("Legado:", test.result.legacyImages));
    children.push(...evidenceSection("Novo:", test.result.newImages));
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

function idemTestTable(
  index: number,
  test: PermissionBlockTest,
  referenceLabel: string,
): Table {
  return simpleTable(
    [
      [cell(`${index} - ${test.title || "Teste"}`, true, { columnSpan: 2, fill: COLORS.testFill })],
      [cell("IDEM ao teste", true), cell(referenceLabel)],
    ],
    [24, 76],
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

function createImageRun(image: EvidenceImage): ImageRun {
  const maxWidth = 560;
  const maxHeight = 420;
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
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: tableBorders(),
    rows: rows.map(
      (row) =>
        new TableRow({
          children: row.map((definition, index) => {
            const columnSpan = definition.columnSpan ?? 1;
            const width = columnWidths
              .slice(index, index + columnSpan)
              .reduce((total, value) => total + value, 0);
            const fill = definition.fill ?? (definition.bold ? COLORS.labelFill : undefined);

            return new TableCell({
              width: { size: width || 100, type: WidthType.PERCENTAGE },
              columnSpan: definition.columnSpan,
              margins: cellMargins(),
              verticalAlign: VerticalAlign.CENTER,
              shading: fill
                ? { type: ShadingType.CLEAR, fill, color: "auto" }
                : undefined,
              children: [
                new Paragraph({
                  alignment: definition.alignment ?? AlignmentType.LEFT,
                  spacing: { before: 0, after: 0 },
                  children: [
                    run(definition.text || " ", {
                      bold: definition.bold,
                      color: definition.bold ? COLORS.title : COLORS.text,
                    }),
                  ],
                }),
              ],
            });
          }),
        }),
    ),
  });
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

function getPermissionTestLabel(
  documentData: OtDocument,
  referenceKey: string | undefined,
): string {
  if (!referenceKey) {
    return "";
  }

  for (const macro of documentData.permissionGroups) {
    for (const micro of macro.microPermissions) {
      const blockKey = createPermissionKey(macro.id, micro.id);
      const block = documentData.permissionBlocks[blockKey];

      for (const [index, test] of (block?.tests ?? []).entries()) {
        if (createTestReferenceKey(blockKey, test.id) === referenceKey) {
          return `${formatPermission(macro)} / ${formatPermission(micro)} / ${
            test.title.trim() || `Teste ${index + 1}`
          }`;
        }
      }
    }
  }

  return "";
}

function formatPermission(permission: PermissionItem): string {
  const code = permission.code.trim();
  const label = permission.label.trim();

  if (code && label) {
    return `${code} (${label})`;
  }

  return code || label || "Sem código";
}

function createEmptyBlock(): PermissionBlock {
  return { tests: [] };
}

function createTestReferenceKey(blockKey: string, testId: string): string {
  return `${blockKey}::${testId}`;
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

function sanitizeFileName(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
