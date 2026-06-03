import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
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
    heading("OBSERVABILIDADE DE TESTES (PERMISSÕES + NAVEGAÇÃO)", true),
    paragraph([
      run("Objetivo: ", { bold: true }),
      run(documentData.objective || " "),
    ]),
    metadataTable(documentData),
    heading("Passo a Passo para Acessar a Tela"),
    ...documentData.accessSteps
      .filter((step) => step.text.trim())
      .map((step, index) => paragraph([run(`${index + 1}. ${step.text}`)])),
    heading("Tipos de Permissão para Testar"),
    ...buildPermissionSummary(selectedGroups),
    heading("TESTES"),
  ];

  for (const macro of selectedGroups) {
    for (const micro of macro.microPermissions) {
      children.push(...buildPermissionBlock(documentData, macro, micro));
    }
  }

  return children;
}

function buildPermissionSummary(groups: PermissionGroup[]): Paragraph[] {
  return groups.flatMap((macro) => [
    paragraph([run("Macro-permissão: ", { bold: true }), run(formatPermission(macro))]),
    ...macro.microPermissions.map((micro) =>
      paragraph([run("Micro-permissão: ", { bold: true }), run(formatPermission(micro))]),
    ),
  ]);
}

function metadataTable(documentData: OtDocument): Table {
  const rows = [
    ["Tela:", documentData.metadata.screen],
    ["Responsável pelo teste:", documentData.metadata.responsible],
    ["Data:", formatDisplayDate(documentData.metadata.date)],
    ["Ambiente:", documentData.metadata.environment],
    ["Elaborada por:", documentData.metadata.author],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders(),
    rows: rows.map(
      ([label, value]) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: 34, type: WidthType.PERCENTAGE },
              margins: cellMargins(),
              verticalAlign: VerticalAlign.CENTER,
              children: [paragraph([run(label, { bold: true })])],
            }),
            new TableCell({
              width: { size: 66, type: WidthType.PERCENTAGE },
              margins: cellMargins(),
              verticalAlign: VerticalAlign.CENTER,
              children: [paragraph([run(value || " ")])],
            }),
          ],
        }),
    ),
  });
}

function buildPermissionBlock(
  documentData: OtDocument,
  macro: PermissionGroup,
  micro: PermissionItem,
) {
  const key = createPermissionKey(macro.id, micro.id);
  const block = documentData.permissionBlocks[key] ?? createEmptyBlock();
  const children = [
    paragraph([run("MACRO-PERMISSÃO", { bold: true, italics: true })]),
    heading(`Tipo de usuário: ${formatPermission(macro)}`),
    paragraph([run("MICRO-PERMISSÃO", { bold: true, italics: true })]),
    heading(`Tipo de permissão: ${formatPermission(micro)}`),
  ];

  block.tests.forEach((test, index) => {
    if (test.mode === "idem") {
      children.push(idemTestHeader(index + 1, test));
      children.push(
        paragraph([
          run(
            `IDEM ao teste: ${
              getPermissionTestLabel(documentData, test.idemReferenceKey) ||
              "referência não informada"
            }`,
            { bold: true },
          ),
        ]),
      );
      return;
    }

    children.push(testHeader(index + 1, test));
    if (test.result.observations.trim()) {
      children.push(paragraph([run(test.result.observations.trim(), { bold: true })]));
    }

    children.push(...evidenceSection("Legado:", test.result.legacyImages));
    children.push(...evidenceSection("Novo:", test.result.newImages));
  });

  return children;
}

function idemTestHeader(index: number, test: PermissionBlockTest): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 120 },
    children: [run(`${index} - ${test.title || "Teste"}  `, { bold: true })],
  });
}

function testHeader(index: number, test: PermissionBlockTest): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 120 },
    children: [
      run(`${index} - ${test.title || "Teste"}  `, { bold: true }),
      ...checkOrder.flatMap((key) => [
        run(`${test.result.checks[key] ? "( X )" : "(    )"} ${checkLabels[key]} `, {
          bold: key === "bothIssue" && test.result.checks[key],
        }),
      ]),
    ],
  });
}

function evidenceSection(label: string, images: EvidenceImage[]) {
  const children = [paragraph([run(label, { bold: true })])];

  if (images.length === 0) {
    return children;
  }

  for (const image of images) {
    if (image.label.trim()) {
      children.push(paragraph([run(image.label.trim(), { bold: true })]));
    }

    children.push(
      new Paragraph({
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

function heading(text: string, centered = false): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: centered ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [run(text, { bold: true, italics: true })],
  });
}

function paragraph(children: TextRun[]): Paragraph {
  return new Paragraph({
    children,
    spacing: { after: 120 },
  });
}

function run(
  text: string,
  options: { bold?: boolean; italics?: boolean } = {},
): TextRun {
  return new TextRun({
    text,
    bold: options.bold,
    italics: options.italics,
    font: "Arial",
    size: 22,
  });
}

function tableBorders() {
  return {
    top: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
    left: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
    right: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
    insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "666666" },
  };
}

function cellMargins() {
  return {
    top: 90,
    bottom: 90,
    left: 140,
    right: 140,
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
