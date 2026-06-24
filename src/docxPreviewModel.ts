import {
  checkLabels,
  checkOrder,
  createPermissionKey,
  getEffectiveChecks,
} from "./defaultDocument";
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
} from "./types";

export type DocxPreviewKind = "ot" | "tea";
export type DocxPreviewAlignment = "left" | "center";

export type DocxPreviewRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

export type DocxPreviewParagraph = {
  id: string;
  type: "paragraph";
  variant:
    | "ot-title"
    | "ot-heading"
    | "ot-empty"
    | "ot-label"
    | "ot-spacer"
    | "tea-title"
    | "tea-heading1"
    | "tea-heading2"
    | "tea-subheading"
    | "tea-text"
    | "tea-image-label";
  runs: DocxPreviewRun[];
  alignment?: DocxPreviewAlignment;
  fill?: string;
  anchorId?: string;
};

export type DocxPreviewCell = {
  paragraphs: DocxPreviewRun[][];
  bold?: boolean;
  alignment?: DocxPreviewAlignment;
  columnSpan?: number;
  fill?: string;
};

export type DocxPreviewTable = {
  id: string;
  type: "table";
  variant: "ot" | "tea";
  columnWidths: number[];
  rows: DocxPreviewCell[][];
  anchorId?: string;
};

export type DocxPreviewList = {
  id: string;
  type: "list";
  variant: "tea";
  items: DocxPreviewRun[][];
};

export type DocxPreviewImage = {
  id: string;
  type: "image";
  variant: "ot" | "tea";
  alt: string;
  label?: string;
  src: string;
  width: number;
  height: number;
};

export type DocxPreviewBlock =
  | DocxPreviewParagraph
  | DocxPreviewTable
  | DocxPreviewList
  | DocxPreviewImage;

export type DocxPreviewModel = {
  kind: DocxPreviewKind;
  sectionId: string;
  title: string;
  blocks: DocxPreviewBlock[];
};

const OT_IMAGE_MAX_WIDTH = 560;
const OT_IMAGE_MAX_HEIGHT = 420;
const TEA_IMAGE_MAX_WIDTH = 576;
const TEA_IMAGE_MAX_HEIGHT = 720;

const COLORS = {
  title: "111827",
  border: "AEB7C2",
  borderSoft: "D8DEE6",
  sectionFill: "E7F0FA",
  labelFill: "F2F6FA",
  testFill: "F6FAFE",
  selectedFill: "EAF7EF",
  emptyFill: "F8FAFC",
};

export function buildOtPreviewModel(documentData: OtDocument): DocxPreviewModel {
  const selectedGroups = selectedPermissionGroups(documentData.permissionGroups);
  const blocks: DocxPreviewBlock[] = [
    paragraph("ot-title", "OBSERVABILIDADE DE TESTES (PERMISSÕES + NAVEGAÇÃO)", {
      alignment: "center",
      runs: [{ text: "OBSERVABILIDADE DE TESTES (PERMISSÕES + NAVEGAÇÃO)", bold: true, italic: true }],
    }),
    labeledBox("Objetivo", documentData.objective || " "),
    metadataTable(documentData),
    paragraph("ot-heading", "Passo a Passo para Acessar a Tela", {
      fill: COLORS.sectionFill,
      runs: [{ text: "Passo a Passo para Acessar a Tela", bold: true, italic: true }],
    }),
    accessStepsTable(documentData),
    paragraph("ot-heading", "Tipos de Permissão para Testar", {
      fill: COLORS.sectionFill,
      runs: [{ text: "Tipos de Permissão para Testar", bold: true, italic: true }],
    }),
    permissionSummaryTable(selectedGroups),
    paragraph("ot-heading", "TESTES", {
      fill: COLORS.sectionFill,
      runs: [{ text: "TESTES", bold: true, italic: true }],
    }),
  ];

  if (selectedGroups.length === 0) {
    blocks.push(emptyParagraph("Nenhuma permissão selecionada para gerar testes."));
  }

  selectedGroups.forEach((macro) => {
    macro.microPermissions.forEach((micro) => {
      blocks.push(...buildPermissionBlock(documentData, macro, micro));
    });
  });

  return {
    kind: "ot",
    sectionId: "ot-section-preview",
    title: "Prévia DOCX da OT",
    blocks,
  };
}

export function buildTeaPreviewModel(documentData: TeaDocument): DocxPreviewModel {
  const blocks: DocxPreviewBlock[] = [
    paragraph("tea-title", "Termo de Entrega de Atividade (TEA)", {
      alignment: "center",
      runs: [{ text: "Termo de Entrega de Atividade (TEA)", bold: true }],
    }),
    teaMetadataTable(documentData),
    paragraph("tea-heading1", "1. VISÃO GERAL", {
      runs: [{ text: "1. VISÃO GERAL", bold: true, italic: true, underline: true }],
    }),
    ...teaParagraphs(documentData.overview || " "),
    paragraph("tea-heading1", "2. ATIVIDADES REALIZADAS", {
      runs: [
        { text: "2. ATIVIDADES REALIZADAS", bold: true, italic: true, underline: true },
      ],
    }),
    ...teaParagraphs(
      documentData.activityIntro ||
        "A seguir serao apresentadas, a nova interface e as suas funcionalidades:",
    ),
    ...teaEvidenceSection(documentData.activityImages),
  ];

  documentData.activities.forEach((activity, index) => {
    blocks.push(...teaActivitySection(activity, index + 1));
  });

  return {
    kind: "tea",
    sectionId: "tea-section-preview",
    title: "Prévia DOCX do TEA",
    blocks,
  };
}

function buildPermissionBlock(
  documentData: OtDocument,
  macro: PermissionGroup,
  micro: PermissionItem,
): DocxPreviewBlock[] {
  const key = createPermissionKey(macro.id, micro.id);
  const block = documentData.permissionBlocks[key] ?? createEmptyBlock();
  const blocks: DocxPreviewBlock[] = [
    permissionContextTable(macro, micro),
    spacerParagraph(),
  ];

  if (block.tests.length === 0) {
    blocks.push(emptyParagraph("Nenhum teste informado para esta permissão."));
    return blocks;
  }

  block.tests.forEach((test, index) => {
    const referenceKey = createPermissionTestReferenceKey(key, test.id);

    blocks.push(testResultTable(index + 1, test, referenceKey));

    if (test.result.observations.trim()) {
      blocks.push(labeledBox("Observações", test.result.observations.trim()));
    }

    if (test.result.errors.length === 0) {
      blocks.push(...evidenceSection("Legado:", test.result.legacyImages));
      blocks.push(...evidenceSection("Novo:", test.result.newImages));
    }
    blocks.push(...testErrorsSection(test));
  });

  return blocks;
}

function accessStepsTable(documentData: OtDocument): DocxPreviewTable {
  const steps = documentData.accessSteps.filter((step) => step.text.trim());

  if (steps.length === 0) {
    return simpleTable([[cell("Passo", true), cell("Nenhum passo informado.")]], [16, 84]);
  }

  return simpleTable(
    steps.map((step, index) => [
      cell(`${index + 1}.`, true, { alignment: "center" }),
      cell(step.text.trim()),
    ]),
    [10, 90],
  );
}

function permissionSummaryTable(groups: PermissionGroup[]): DocxPreviewTable {
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

function metadataTable(documentData: OtDocument): DocxPreviewTable {
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

function permissionContextTable(
  macro: PermissionGroup,
  micro: PermissionItem,
): DocxPreviewTable {
  return simpleTable(
    [
      [cell("MACRO-PERMISSÃO", true), cell(`Tipo de usuário: ${formatPermission(macro)}`)],
      [cell("MICRO-PERMISSÃO", true), cell(`Tipo de permissão: ${formatPermission(micro)}`)],
    ],
    [28, 72],
  );
}

function testResultTable(
  index: number,
  test: PermissionBlockTest,
  referenceKey: string,
): DocxPreviewTable {
  const effectiveChecks = getEffectiveChecks(test.result.checks, test.result.errors);

  return simpleTable(
    [
      [
        cell(`${index} - ${test.title || "Teste"}`, true, {
          columnSpan: 2,
          fill: COLORS.testFill,
        }),
      ],
      ...checkOrder.map((key) => [
        cell(effectiveChecks[key] ? "( X )" : "(   )", effectiveChecks[key], {
          alignment: "center",
          fill: effectiveChecks[key] ? COLORS.selectedFill : undefined,
        }),
        cell(checkLabels[key], key === "bothIssue" && effectiveChecks[key], {
          fill: effectiveChecks[key] ? COLORS.selectedFill : undefined,
        }),
      ]),
    ],
    [14, 86],
    { anchorId: `ot-preview-test-${toPreviewDomId(referenceKey)}` },
  );
}

function testErrorsSection(test: PermissionBlockTest): DocxPreviewBlock[] {
  if (test.result.errors.length === 0) {
    return [];
  }

  return [
    paragraph("ot-label", "Erros encontrados:", {
      runs: [{ text: "Erros encontrados:", bold: true }],
    }),
    ...test.result.errors.flatMap((error, index) => testErrorSection(error, index)),
  ];
}

function testErrorSection(error: TestError, index: number): DocxPreviewBlock[] {
  return [
    paragraph("ot-label", "Correcao:", {
      runs: [
        {
          text: `Erro ${index + 1} - ${formatTestErrorOrigin(error.origin)}`,
          bold: true,
        },
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
    ...(error.origin === "new" ? legacyReferenceSection(error.legacyReference) : []),
    ...(error.origin === "legacy" ? legacyNewStatusSection(error.newStatus) : []),
    ...(error.origin === "new" ? correctionSection(error.correction) : []),
  ];
}

function legacyReferenceSection(
  reference: TestError["legacyReference"],
): DocxPreviewBlock[] {
  if (!hasMeaningfulLegacyReference(reference)) {
    return [];
  }

  return [
    paragraph("ot-label", "Como e no legado que esta certo:", {
      runs: [{ text: "Como e no legado que esta certo:", bold: true }],
    }),
    simpleTable(
      [[cell("Descricao", true), cell(reference.description.trim() || " ")]],
      [28, 72],
    ),
    ...evidenceSection("Prints do legado correto:", reference.images),
  ];
}

function legacyNewStatusSection(status: TestError["newStatus"]): DocxPreviewBlock[] {
  return [
    paragraph("ot-label", "Situacao no novo:", {
      runs: [{ text: "Situacao no novo:", bold: true }],
    }),
    simpleTable(
      [[cell("Situacao no novo", true), cell(status.works ? "Funciona" : "Tambem precisa ajuste")]],
      [28, 72],
    ),
    ...(!status.works ? evidenceSection("Prints do erro no novo:", status.images) : []),
  ];
}

function correctionSection(correction: TestError["correction"]): DocxPreviewBlock[] {
  return [
    paragraph("ot-label", "Correcao:", {
      runs: [{ text: "Correcao:", bold: true }],
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

function evidenceSection(label: string, images: EvidenceImage[]): DocxPreviewBlock[] {
  const blocks: DocxPreviewBlock[] = [
    paragraph("ot-label", label, {
      runs: [{ text: label, bold: true }],
    }),
  ];

  images.forEach((image) => {
    if (!hasImageData(image)) {
      blocks.push(emptyParagraph(`${image.name || "Imagem"} nao encontrada no rascunho.`));
      return;
    }

    if (image.label.trim()) {
      blocks.push(
        paragraph("ot-label", image.label.trim(), {
          runs: [{ text: image.label.trim(), bold: true }],
        }),
      );
    }

    blocks.push(previewImage(image, "ot", OT_IMAGE_MAX_WIDTH, OT_IMAGE_MAX_HEIGHT));
  });

  return blocks;
}

function teaActivitySection(activity: TeaActivity, index: number): DocxPreviewBlock[] {
  const blocks: DocxPreviewBlock[] = [
    paragraph("tea-heading2", `2.${index} - ${activity.title || "Atividade"}:`, {
      anchorId: `tea-preview-activity-${toPreviewDomId(activity.id)}`,
      runs: [
        {
          text: `2.${index} - ${activity.title || "Atividade"}:`,
          bold: true,
          underline: true,
        },
      ],
    }),
    ...teaContentBlocks(activity.blocks),
  ];

  activity.subActivities.forEach((subActivity, subIndex) => {
    blocks.push(...teaSubActivitySection(subActivity, index, subIndex + 1));
  });

  return blocks;
}

function teaSubActivitySection(
  subActivity: TeaSubActivity,
  activityIndex: number,
  subIndex: number,
): DocxPreviewBlock[] {
  const heading = `2.${activityIndex}.${subIndex} - ${subActivity.title || "Subatividade"}:`;

  return [
    paragraph("tea-subheading", heading, {
      anchorId: `tea-preview-subactivity-${toPreviewDomId(subActivity.id)}`,
      runs: [{ text: heading, bold: true }],
    }),
    ...teaContentBlocks(subActivity.blocks),
  ];
}

function teaContentBlocks(blocks: TeaContentBlock[]): DocxPreviewBlock[] {
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

function teaMetadataTable(documentData: TeaDocument): DocxPreviewTable {
  const rows = [
    ["Ordem de Serviço:", documentData.metadata.serviceOrder],
    ["Fase/Etapa", documentData.metadata.phase],
    ["Chamado:", documentData.metadata.ticket],
    ["Assunto:", documentData.metadata.subject],
    ["Data:", formatDisplayDate(documentData.metadata.date)],
    ["Elaborado por:", documentData.metadata.author],
  ];

  return {
    id: createPreviewId("table", "tea-metadata"),
    type: "table",
    variant: "tea",
    columnWidths: [25.1736111111, 74.8263888889],
    rows: rows.map(([label, value]) => [cell(label, true), cell(value || " ")]),
  };
}

function teaParagraphs(value: string): DocxPreviewParagraph[] {
  const paragraphs = value
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  return paragraphs.map((text) =>
    paragraph("tea-text", text, { runs: teaRunsFromMarkdown(text) }),
  );
}

function teaBullets(items: { text: string }[]): DocxPreviewList[] {
  const listItems = items
    .map((item) => item.text.trim())
    .filter(Boolean)
    .map((text) => teaRunsFromMarkdown(text));

  return listItems.length > 0
    ? [
        {
          id: createPreviewId("list", listItems.map((runs) => runs.map((run) => run.text).join("")).join("-")),
          type: "list",
          variant: "tea",
          items: listItems,
        },
      ]
    : [];
}

function teaEvidenceSection(images: EvidenceImage[]): DocxPreviewBlock[] {
  return images.flatMap((image) => {
    if (!hasImageData(image)) {
      return [];
    }

    const blocks: DocxPreviewBlock[] = [];

    if (image.label.trim()) {
      blocks.push(
        paragraph("tea-image-label", image.label.trim(), {
          runs: [{ text: image.label.trim(), bold: true }],
        }),
      );
    }

    blocks.push(previewImage(image, "tea", TEA_IMAGE_MAX_WIDTH, TEA_IMAGE_MAX_HEIGHT));
    return blocks;
  });
}

function teaRunsFromMarkdown(text: string): DocxPreviewRun[] {
  const markerMatches = text.match(/\*\*/g) ?? [];

  if (markerMatches.length === 0 || markerMatches.length % 2 !== 0) {
    return [{ text }];
  }

  const runs: DocxPreviewRun[] = [];
  let cursor = 0;
  let bold = false;

  while (cursor < text.length) {
    const markerIndex = text.indexOf("**", cursor);

    if (markerIndex === -1) {
      const value = text.slice(cursor);
      if (value) {
        runs.push({ text: value, bold });
      }
      break;
    }

    const value = text.slice(cursor, markerIndex);
    if (value) {
      runs.push({ text: value, bold });
    }

    bold = !bold;
    cursor = markerIndex + 2;
  }

  return runs.length > 0 ? runs : [{ text: text.replace(/\*\*/g, "") }];
}

function labeledBox(label: string, value: string): DocxPreviewTable {
  return simpleTable([[cell(`${label}:`, true), cell(value || " ")]], [22, 78]);
}

function simpleTable(
  rows: DocxPreviewCell[][],
  columnWidths: number[],
  options: Pick<DocxPreviewTable, "anchorId"> = {},
): DocxPreviewTable {
  return {
    id: createPreviewId(
      "table",
      rows
        .flat()
        .flatMap((row) => row.paragraphs.flatMap((runs) => runs.map((run) => run.text)))
        .join("-"),
    ),
    type: "table",
    variant: "ot",
    columnWidths,
    rows,
    ...options,
  };
}

function cell(
  text: string,
  bold = false,
  options: Omit<DocxPreviewCell, "paragraphs" | "bold"> = {},
): DocxPreviewCell {
  return {
    paragraphs: cellParagraphRuns(text, bold),
    bold,
    fill: options.fill ?? (bold ? COLORS.labelFill : undefined),
    ...options,
  };
}

function cellParagraphRuns(text: string, bold: boolean): DocxPreviewRun[][] {
  const paragraphs = (text || " ").replace(/\r\n/g, "\n").split("\n");

  return (paragraphs.length > 0 ? paragraphs : [" "]).map((paragraphText) => [
    { text: paragraphText || " ", bold },
  ]);
}

function paragraph(
  variant: DocxPreviewParagraph["variant"],
  idSeed: string,
  options: Omit<DocxPreviewParagraph, "id" | "type" | "variant">,
): DocxPreviewParagraph {
  return {
    id: createPreviewId("paragraph", `${variant}-${idSeed}`),
    type: "paragraph",
    variant,
    ...options,
  };
}

function emptyParagraph(text: string): DocxPreviewParagraph {
  return paragraph("ot-empty", text, {
    fill: COLORS.emptyFill,
    runs: [{ text, italic: true }],
  });
}

function spacerParagraph(): DocxPreviewParagraph {
  return paragraph("ot-spacer", "spacer", { runs: [{ text: " " }] });
}

function previewImage(
  image: EvidenceImage & { dataUrl: string },
  variant: DocxPreviewImage["variant"],
  maxWidth: number,
  maxHeight: number,
): DocxPreviewImage {
  const width = image.width || maxWidth;
  const height = image.height || maxHeight;
  const scale = Math.min(maxWidth / width, maxHeight / height, 1);
  const displayWidth = Math.max(1, Math.round(width * scale));
  const displayHeight = Math.max(1, Math.round(height * scale));
  const label = image.label.trim();
  const alt = label || image.name || "Imagem do documento";

  return {
    id: createPreviewId("image", image.id),
    type: "image",
    variant,
    alt,
    label: label || undefined,
    src: image.dataUrl,
    width: displayWidth,
    height: displayHeight,
  };
}

function hasImageData(image: EvidenceImage): image is EvidenceImage & { dataUrl: string } {
  return typeof image.dataUrl === "string" && image.dataUrl.trim().length > 0;
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

function createPermissionTestReferenceKey(blockKey: string, testId: string): string {
  return `${blockKey}--${testId}`;
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

function hasMeaningfulLegacyReference(reference: TestError["legacyReference"]): boolean {
  return (
    reference.enabled ||
    Boolean(reference.description.trim()) ||
    reference.images.length > 0
  );
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

function createPreviewId(prefix: string, value: string): string {
  return `${prefix}-${value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase() || "item"}`;
}

function toPreviewDomId(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]/g, "-");
}
