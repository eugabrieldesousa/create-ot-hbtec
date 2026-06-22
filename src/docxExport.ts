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
  createPermissionKey,
  getEffectiveChecks,
} from "./defaultDocument";
import { mapWithConcurrency } from "./asyncUtils";
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
  TestError,
  TestResult,
} from "./types";

type ExportableEvidenceImage = EvidenceImage & { dataUrl: string };
export type DocxExportKind = "ot" | "tea";
export type DocxExportImageProblem = {
  label: string;
  detail: string;
  location: string;
  severity: "danger";
  documentKind: DocxExportKind;
};
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

export class DocxExportImageError extends Error {
  readonly documentKind: DocxExportKind;
  readonly problems: DocxExportImageProblem[];

  constructor(documentKind: DocxExportKind, problems: DocxExportImageProblem[]) {
    super("Exportacao bloqueada por problema nas imagens.");
    this.name = "DocxExportImageError";
    this.documentKind = documentKind;
    this.problems = problems;
  }
}

export function isDocxExportImageError(error: unknown): error is DocxExportImageError {
  return (
    error instanceof Error &&
    error.name === "DocxExportImageError" &&
    Array.isArray((error as Partial<DocxExportImageError>).problems)
  );
}

export async function exportOtDocument(documentData: OtDocument): Promise<void> {
  const hydratedDocument = await hydrateDocumentImages(documentData);
  assertNoExportImageProblems(validateOtDocumentImages(hydratedDocument));
  const exportDocument = await optimizeOtDocumentImages(hydratedDocument);
  const expectedImages = collectOtExportImages(exportDocument);
  assertNoExportImageProblems(validateExportableImageData(expectedImages, "ot"));
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

  await downloadDocument(doc, createFileName(exportDocument), expectedImages, "ot");
}

export async function exportTeaDocument(documentData: TeaDocument): Promise<void> {
  const hydratedDocument = await hydrateTeaDocumentImages(documentData);
  assertNoExportImageProblems(validateTeaDocumentImages(hydratedDocument));
  const exportDocument = await optimizeTeaDocumentImages(hydratedDocument);
  const expectedImages = collectTeaExportImages(exportDocument);
  assertNoExportImageProblems(validateExportableImageData(expectedImages, "tea"));
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

  await downloadDocument(doc, createTeaFileName(exportDocument), expectedImages, "tea");
}

async function downloadDocument(
  doc: Document,
  fileName: string,
  expectedImages: LocatedEvidenceImage[],
  documentKind: DocxExportKind,
): Promise<void> {
  const blob = await Packer.toBlob(doc);
  await verifyDocxImageMedia(blob, expectedImages, documentKind);
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

type LocatedEvidenceImage = {
  image: EvidenceImage;
  location: string;
};

type ParsedImageDataUrl = {
  bytes: Uint8Array;
  extension: "png" | "jpg" | "gif" | "bmp";
  signature: string;
};
type ZipFileMap = Record<
  string,
  { dir: boolean; async: (type: "uint8array") => Promise<Uint8Array> }
>;

function assertNoExportImageProblems(problems: DocxExportImageProblem[]): void {
  if (problems.length > 0) {
    throw new DocxExportImageError(problems[0].documentKind, problems);
  }
}

function validateOtDocumentImages(documentData: OtDocument): DocxExportImageProblem[] {
  const problems: DocxExportImageProblem[] = [];
  const selectedGroups = selectedPermissionGroups(documentData.permissionGroups);

  selectedGroups.forEach((macro) => {
    macro.microPermissions.forEach((micro) => {
      const block = documentData.permissionBlocks[createPermissionKey(macro.id, micro.id)] ??
        createEmptyBlock();

      block.tests.forEach((test, testIndex) => {
        const testLocation = formatOtTestLocation(macro, micro, test, testIndex);

        if (test.result.legacyImages.length === 0) {
          problems.push(createImageProblem("ot", {
            label: "Imagem do legado ausente",
            detail: "Adicione evidencia em Legado antes de exportar.",
            location: `${testLocation} > Legado`,
          }));
        }

        if (test.result.newImages.length === 0) {
          problems.push(createImageProblem("ot", {
            label: "Imagem do novo ausente",
            detail: "Adicione evidencia em Novo antes de exportar.",
            location: `${testLocation} > Novo`,
          }));
        }

        test.result.errors.forEach((error, errorIndex) => {
          if (error.images.length === 0) {
            problems.push(createImageProblem("ot", {
              label: "Print do erro ausente",
              detail: "Adicione ao menos um print do erro antes de exportar.",
              location:
                `${testLocation} > Erros encontrados > Erro ${errorIndex + 1} ` +
                `(${formatTestErrorOrigin(error.origin)})`,
            }));
          }
        });

        collectOtTestImages(test, testLocation).forEach((located) => {
          if (!hasImageData(located.image)) {
            problems.push(createImageProblem("ot", {
              label: "Imagem sem dados",
              detail:
                `${located.image.name || "Imagem"} nao foi carregada do rascunho. ` +
                "Recarregue ou substitua a imagem.",
              location: located.location,
            }));
          }
        });
      });
    });
  });

  return problems;
}

function validateTeaDocumentImages(documentData: TeaDocument): DocxExportImageProblem[] {
  const problems: DocxExportImageProblem[] = [];

  documentData.activityImages.forEach((image, imageIndex) => {
    if (!hasImageData(image)) {
      problems.push(createImageProblem("tea", {
        label: "Imagem geral sem dados",
        detail:
          `${image.name || "Imagem"} nao foi carregada do rascunho. ` +
          "Recarregue ou substitua a imagem.",
        location: `TEA > Imagem geral > Imagem ${imageIndex + 1}`,
      }));
    }
  });

  documentData.activities.forEach((activity, activityIndex) => {
    const activityLocation = `TEA > Atividade 2.${activityIndex + 1}`;

    activity.blocks.forEach((block, blockIndex) => {
      addTeaBlockImageProblems(problems, block, {
        location: `${activityLocation} > Bloco ${blockIndex + 1}`,
      });
    });

    activity.subActivities.forEach((subActivity, subActivityIndex) => {
      const subActivityLocation =
        `${activityLocation} > Subtopico 2.${activityIndex + 1}.${subActivityIndex + 1}`;

      subActivity.blocks.forEach((block, blockIndex) => {
        addTeaBlockImageProblems(problems, block, {
          location: `${subActivityLocation} > Bloco ${blockIndex + 1}`,
        });
      });
    });
  });

  return problems;
}

function addTeaBlockImageProblems(
  problems: DocxExportImageProblem[],
  block: TeaContentBlock,
  context: { location: string },
): void {
  if (block.type !== "images") {
    return;
  }

  if (block.images.length === 0) {
    problems.push(createImageProblem("tea", {
      label: "Bloco de imagens vazio",
      detail: "Adicione uma imagem ou remova o bloco antes de exportar.",
      location: context.location,
    }));
    return;
  }

  block.images.forEach((image, imageIndex) => {
    if (!hasImageData(image)) {
      problems.push(createImageProblem("tea", {
        label: "Imagem sem dados",
        detail:
          `${image.name || "Imagem"} nao foi carregada do rascunho. ` +
          "Recarregue ou substitua a imagem.",
        location: `${context.location} > Imagem ${imageIndex + 1}`,
      }));
    }
  });
}

function validateExportableImageData(
  expectedImages: LocatedEvidenceImage[],
  documentKind: DocxExportKind,
): DocxExportImageProblem[] {
  return expectedImages.flatMap(({ image, location }) => {
    if (!hasImageData(image)) {
      return [
        createImageProblem(documentKind, {
          label: "Imagem sem dados",
          detail: `${image.name || "Imagem"} nao tem dados para exportacao.`,
          location,
        }),
      ];
    }

    try {
      parseImageDataUrl(image.dataUrl);
      return [];
    } catch (error) {
      return [
        createImageProblem(documentKind, {
          label: "Imagem invalida",
          detail:
            `${image.name || "Imagem"} nao pode ser exportada: ` +
            getErrorMessage(error),
          location,
        }),
      ];
    }
  });
}

function collectOtExportImages(documentData: OtDocument): LocatedEvidenceImage[] {
  return selectedPermissionGroups(documentData.permissionGroups).flatMap((macro) =>
    macro.microPermissions.flatMap((micro) => {
      const block = documentData.permissionBlocks[createPermissionKey(macro.id, micro.id)] ??
        createEmptyBlock();

      return block.tests.flatMap((test, testIndex) =>
        collectOtTestImages(test, formatOtTestLocation(macro, micro, test, testIndex)),
      );
    }),
  );
}

function collectOtTestImages(
  test: PermissionBlockTest,
  testLocation: string,
): LocatedEvidenceImage[] {
  const images: LocatedEvidenceImage[] = [
    ...test.result.legacyImages.map((image, imageIndex) => ({
      image,
      location: `${testLocation} > Legado > Imagem ${imageIndex + 1}`,
    })),
    ...test.result.newImages.map((image, imageIndex) => ({
      image,
      location: `${testLocation} > Novo > Imagem ${imageIndex + 1}`,
    })),
    ...test.result.errors.flatMap((error, errorIndex) => [
      ...error.images.map((image, imageIndex) => ({
        image,
        location:
          `${testLocation} > Erros encontrados > Erro ${errorIndex + 1} ` +
          `(${formatTestErrorOrigin(error.origin)}) > Imagem ${imageIndex + 1}`,
      })),
      ...error.correction.beforeImages.map((image, imageIndex) => ({
        image,
        location:
          `${testLocation} > Erros encontrados > Erro ${errorIndex + 1} > ` +
          `Antes (com erro) > Imagem ${imageIndex + 1}`,
      })),
      ...error.correction.afterImages.map((image, imageIndex) => ({
        image,
        location:
          `${testLocation} > Erros encontrados > Erro ${errorIndex + 1} > ` +
          `Depois (corrigido) > Imagem ${imageIndex + 1}`,
      })),
    ]),
  ];

  return images;
}

function collectTeaExportImages(documentData: TeaDocument): LocatedEvidenceImage[] {
  return [
    ...documentData.activityImages.map((image, imageIndex) => ({
      image,
      location: `TEA > Imagem geral > Imagem ${imageIndex + 1}`,
    })),
    ...documentData.activities.flatMap((activity, activityIndex) => {
      const activityLocation = `TEA > Atividade 2.${activityIndex + 1}`;

      return [
        ...collectTeaBlockImages(activity.blocks, activityLocation),
        ...activity.subActivities.flatMap((subActivity, subActivityIndex) =>
          collectTeaBlockImages(
            subActivity.blocks,
            `${activityLocation} > Subtopico 2.${activityIndex + 1}.${subActivityIndex + 1}`,
          ),
        ),
      ];
    }),
  ];
}

function collectTeaBlockImages(
  blocks: TeaContentBlock[],
  parentLocation: string,
): LocatedEvidenceImage[] {
  return blocks.flatMap((block, blockIndex) =>
    block.type === "images"
      ? block.images.map((image, imageIndex) => ({
          image,
          location: `${parentLocation} > Bloco ${blockIndex + 1} > Imagem ${imageIndex + 1}`,
        }))
      : [],
  );
}

async function verifyDocxImageMedia(
  blob: Blob,
  expectedImages: LocatedEvidenceImage[],
  documentKind: DocxExportKind,
): Promise<void> {
  const expectedMedia = collectExpectedMedia(expectedImages);

  if (expectedMedia.length === 0) {
    return;
  }

  let zipFiles: ZipFileMap;

  try {
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    zipFiles = zip.files as ZipFileMap;
  } catch (error) {
    throw new DocxExportImageError(documentKind, [
      createImageProblem(documentKind, {
        label: "Falha ao verificar DOCX",
        detail:
          "Nao foi possivel abrir o DOCX gerado para conferir as imagens: " +
          getErrorMessage(error),
        location: "DOCX gerado",
      }),
    ]);
  }

  const mediaFiles = Object.entries(zipFiles).filter(
    ([path, file]) => path.startsWith("word/media/") && !file.dir,
  );
  const mediaSignatures = new Set(
    await mapWithConcurrency(mediaFiles, async ([, file]) =>
      bytesToBase64(await file.async("uint8array")),
    ),
  );
  const missingMedia = expectedMedia.filter((expected) => !mediaSignatures.has(expected.signature));

  if (missingMedia.length === 0) {
    return;
  }

  throw new DocxExportImageError(
    documentKind,
    missingMedia.map((expected) =>
      createImageProblem(documentKind, {
        label: "Imagem nao encontrada no DOCX gerado",
        detail:
          "A imagem foi processada, mas nao apareceu em word/media no arquivo final.",
        location: expected.locations.join(" | "),
      }),
    ),
  );
}

function collectExpectedMedia(
  expectedImages: LocatedEvidenceImage[],
): Array<{ signature: string; locations: string[] }> {
  const mediaBySignature = new Map<string, string[]>();

  expectedImages.forEach(({ image, location }) => {
    if (!hasImageData(image)) {
      return;
    }

    const parsed = parseImageDataUrl(image.dataUrl);
    const locations = mediaBySignature.get(parsed.signature) ?? [];
    locations.push(location);
    mediaBySignature.set(parsed.signature, locations);
  });

  return Array.from(mediaBySignature, ([signature, locations]) => ({ signature, locations }));
}

function parseImageDataUrl(dataUrl: string): ParsedImageDataUrl {
  const [header = "", payload = ""] = dataUrl.split(",");
  const trimmedHeader = header.trim();
  const typeMatch = /^data:image\/(png|jpe?g|gif|bmp);base64$/i.exec(trimmedHeader);

  if (!trimmedHeader.toLowerCase().startsWith("data:image/")) {
    throw new Error("o valor nao e uma data URL de imagem.");
  }

  if (!trimmedHeader.toLowerCase().includes(";base64")) {
    throw new Error("a imagem precisa estar em base64.");
  }

  if (!typeMatch) {
    throw new Error("tipo de imagem nao suportado. Use PNG, JPG, GIF ou BMP.");
  }

  const normalizedPayload = payload.replace(/\s/g, "");

  if (!normalizedPayload) {
    throw new Error("payload base64 vazio.");
  }

  try {
    const binary = window.atob(normalizedPayload);
    const bytes = new Uint8Array(binary.length);

    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }

    const extensionName = typeMatch[1].toLowerCase();

    return {
      bytes,
      extension: extensionName.startsWith("jp")
        ? "jpg"
        : (extensionName as ParsedImageDataUrl["extension"]),
      signature: bytesToBase64(bytes),
    };
  } catch {
    throw new Error("payload base64 invalido.");
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return window.btoa(binary);
}

function createImageProblem(
  documentKind: DocxExportKind,
  problem: Omit<DocxExportImageProblem, "documentKind" | "severity">,
): DocxExportImageProblem {
  return {
    ...problem,
    severity: "danger",
    documentKind,
  };
}

function formatOtTestLocation(
  macro: PermissionGroup,
  micro: PermissionItem,
  test: PermissionBlockTest,
  testIndex: number,
): string {
  return [
    "OT",
    formatPermission(macro),
    formatPermission(micro),
    test.title.trim() || `Teste ${testIndex + 1}`,
  ].join(" > ");
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error && error.message ? error.message : "erro desconhecido.";
}

async function optimizeOtDocumentImages(documentData: OtDocument): Promise<OtDocument> {
  return {
    ...documentData,
    permissionBlocks: Object.fromEntries(
      await mapWithConcurrency(
        Object.entries(documentData.permissionBlocks),
        async ([blockKey, block]) => [
          blockKey,
          {
            ...block,
            tests: await mapWithConcurrency(
              block.tests,
              async (test) => ({
                ...test,
                result: {
                  ...test.result,
                  legacyImages: await optimizeEvidenceImagesForExport(
                    test.result.legacyImages,
                  ),
                  newImages: await optimizeEvidenceImagesForExport(test.result.newImages),
                  errors: await mapWithConcurrency(test.result.errors, async (error) => ({
                    ...error,
                    images: await optimizeEvidenceImagesForExport(error.images),
                    correction: {
                      ...error.correction,
                      beforeImages: await optimizeEvidenceImagesForExport(
                        error.correction.beforeImages,
                      ),
                      afterImages: await optimizeEvidenceImagesForExport(
                        error.correction.afterImages,
                      ),
                    },
                  })),
                },
              }),
            ),
          },
        ],
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
    activities: await mapWithConcurrency(
      documentData.activities,
      async (activity) => ({
        ...activity,
        blocks: await mapWithConcurrency(activity.blocks, optimizeTeaContentBlockImages),
        subActivities: await mapWithConcurrency(
          activity.subActivities,
          async (subActivity) => ({
            ...subActivity,
            blocks: await mapWithConcurrency(subActivity.blocks, optimizeTeaContentBlockImages),
          }),
        ),
      }),
    ),
  };
}

async function optimizeEvidenceImagesForExport(
  images: EvidenceImage[],
): Promise<EvidenceImage[]> {
  return mapWithConcurrency(images, async (image, index) => {
    if (index > 0 && index % EXPORT_IMAGE_YIELD_INTERVAL === 0) {
      await yieldToBrowser();
    }

    return optimizeEvidenceImageForExport(image);
  });
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
    children.push(...testErrorsSection(test));
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
  const effectiveChecks = getEffectiveChecks(test.result.checks, test.result.errors);

  return simpleTable(
    [
      [cell(`${index} - ${test.title || "Teste"}`, true, { columnSpan: 2, fill: COLORS.testFill })],
      ...checkOrder.map((key) => [
        cell(effectiveChecks[key] ? "( X )" : "(   )", effectiveChecks[key], {
          alignment: AlignmentType.CENTER,
          fill: effectiveChecks[key] ? COLORS.selectedFill : undefined,
        }),
        cell(checkLabels[key], key === "bothIssue" && effectiveChecks[key], {
          fill: effectiveChecks[key] ? COLORS.selectedFill : undefined,
        }),
      ]),
    ],
    [14, 86],
  );
}

function testErrorsSection(test: PermissionBlockTest) {
  if (test.result.errors.length === 0) {
    return [];
  }

  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [run("Erros encontrados:", { bold: true, color: COLORS.title })],
    }),
    ...test.result.errors.flatMap((error, index) => testErrorSection(error, index)),
  ];
}

function testErrorSection(error: TestError, index: number) {
  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [
        run(`Erro ${index + 1} - ${formatTestErrorOrigin(error.origin)}`, {
          bold: true,
          color: COLORS.title,
        }),
      ],
    }),
    simpleTable(
      [
        [cell("Origem", true), cell(formatTestErrorOrigin(error.origin))],
        [cell("Observacao", true), cell(error.observation.trim() || " ")],
      ],
      [28, 72],
    ),
    ...evidenceSection("Prints do erro:", error.images),
    ...(error.origin === "new" ? correctionSection(error.correction) : []),
  ];
}

function correctionSection(correction: TestError["correction"]) {
  return [
    new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [run("Correcao:", { bold: true, color: COLORS.title })],
    }),
    simpleTable(
      [
        [cell("Corrigido", true), cell(correction.corrected ? "Sim" : "Nao")],
        [cell("Hotfix", true), cell(correction.hotfixTag.trim() || " ")],
        [cell("Corrigido por", true), cell(correction.correctedBy.trim() || " ")],
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
  const parsedImage = parseImageDataUrl(image.dataUrl);
  const maxWidth = options.maxWidth ?? 560;
  const maxHeight = options.maxHeight ?? 420;
  const width = image.width || maxWidth;
  const height = image.height || maxHeight;
  const scale = Math.min(maxWidth / width, maxHeight / height, 1);

  return new ImageRun({
    data: parsedImage.bytes,
    transformation: {
      width: Math.max(1, Math.round(width * scale)),
      height: Math.max(1, Math.round(height * scale)),
    },
    type: parsedImage.extension,
  } as never);
}

function hasImageData(image: EvidenceImage): image is ExportableEvidenceImage {
  return typeof image.dataUrl === "string" && image.dataUrl.trim().length > 0;
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

function formatTestErrorOrigin(origin: TestError["origin"]): string {
  return origin === "legacy" ? "Legado" : "Novo";
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
    .replace(/[\x00-\x1f\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
