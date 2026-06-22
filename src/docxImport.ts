import mammoth from "mammoth";
import {
  checkLabelAliases,
  checkOrder,
  createEmptyTestError,
  createEmptyTestCorrection,
  createEmptyTestResult,
  createPermissionKey,
} from "./defaultDocument";
import type {
  CheckKey,
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
  TestCorrection,
  TestError,
} from "./types";

export type DocxImportKind = "ot" | "tea";

export type OtDocxImportSummary = {
  kind: "ot";
  screen: string;
  accessSteps: number;
  permissionGroups: number;
  selectedPermissions: number;
  tests: number;
  images: number;
};

export type TeaDocxImportSummary = {
  kind: "tea";
  subject: string;
  activities: number;
  subActivities: number;
  blocks: number;
  images: number;
};

export type OtDocxImportResult = {
  kind: "ot";
  document: OtDocument;
  summary: OtDocxImportSummary;
  warnings: string[];
  sourceName: string;
};

export type TeaDocxImportResult = {
  kind: "tea";
  document: TeaDocument;
  summary: TeaDocxImportSummary;
  warnings: string[];
  sourceName: string;
};

export type DocxImportResult = OtDocxImportResult | TeaDocxImportResult;

type ParseOptions = {
  sourceName?: string;
  mammothMessages?: Array<{ message?: string; type?: string }>;
};

type Token = {
  text: string;
  markdownText?: string;
  cells?: string[];
  images: EvidenceImage[];
  section: SectionName;
  tagName?: string;
};

type SectionName = "document" | "steps" | "permissions" | "tests" | "other";

type PermissionDraft = Pick<PermissionItem, "code" | "label">;

type ParserState = {
  document: OtDocument;
  warnings: string[];
  currentMacro?: PermissionGroup;
  currentMicro?: PermissionItem;
  currentTest?: PermissionBlockTest;
  currentError?: TestError;
  currentEvidence?: "legacyImages" | "newImages";
  currentErrorEvidence?: boolean;
  currentCorrectionEvidence?: "beforeImages" | "afterImages";
  pendingImageLabel?: string;
  imageCount: number;
};

type TeaParserState = {
  document: TeaDocument;
  warnings: string[];
  currentSection: "document" | "overview" | "activityIntro" | "activities";
  currentActivity?: TeaActivity;
  currentSubActivity?: TeaSubActivity;
  imageCount: number;
  blockCount: number;
  itemCount: number;
  usedIds: Set<string>;
};

const metadataFields: Array<{
  field: keyof OtDocument["metadata"];
  aliases: string[];
}> = [
  { field: "screen", aliases: ["tela"] },
  { field: "responsible", aliases: ["responsavel pelo teste", "responsavel"] },
  { field: "date", aliases: ["data"] },
  { field: "environment", aliases: ["ambiente"] },
  { field: "author", aliases: ["elaborada por", "elaboradora por", "autor"] },
];

export async function parseDocxFile(
  file: File,
  kind: DocxImportKind = "ot",
): Promise<DocxImportResult> {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const options = {
    sourceName: file.name,
    mammothMessages: result.messages,
  };

  return kind === "tea" ? parseTeaHtml(result.value, options) : parseOtHtml(result.value, options);
}

export function parseOtHtml(html: string, options: ParseOptions = {}): OtDocxImportResult {
  const dom = new DOMParser().parseFromString(html, "text/html");
  const tokens = collectTokens(dom.body);
  const state = createParserState();

  applyMammothWarnings(state, options.mammothMessages);
  parseDocumentFields(tokens, state.document);
  parseAccessSteps(tokens, state.document);
  parseTestSection(tokens, state);
  parsePermissionSummary(tokens, state);
  finishDocument(state, options.sourceName ?? "documento-importado.docx");

  return {
    kind: "ot",
    document: state.document,
    summary: buildImportSummary(state.document),
    warnings: state.warnings,
    sourceName: options.sourceName ?? "documento-importado.docx",
  };
}

export function parseTeaHtml(html: string, options: ParseOptions = {}): TeaDocxImportResult {
  const dom = new DOMParser().parseFromString(html, "text/html");
  const tokens = collectTokens(dom.body);
  const sourceName = options.sourceName ?? "documento-importado.docx";
  const state = createTeaParserState();

  applyMammothWarnings(state, options.mammothMessages);
  parseTeaTokens(tokens, state);
  finishTeaDocument(state, sourceName);

  return {
    kind: "tea",
    document: state.document,
    summary: buildTeaImportSummary(state.document),
    warnings: state.warnings,
    sourceName,
  };
}

function createParserState(): ParserState {
  const today = new Date().toISOString().slice(0, 10);

  return {
    document: {
      metadata: {
        screen: "",
        responsible: "GABRIEL",
        date: today,
        environment: "LOCAL + DESENVOLVIMENTO",
        author: "GABRIEL",
      },
      objective: "",
      accessSteps: [],
      permissionGroups: [],
      permissionBlocks: {},
    },
    warnings: [],
    imageCount: 0,
  };
}

function createTeaParserState(): TeaParserState {
  const today = new Date().toISOString().slice(0, 10);

  return {
    document: {
      metadata: {
        serviceOrder: "",
        phase: "",
        ticket: "",
        subject: "",
        date: today,
        author: "",
      },
      overview: "",
      activityIntro: "",
      activityImages: [],
      activities: [],
    },
    warnings: [],
    currentSection: "document",
    imageCount: 0,
    blockCount: 0,
    itemCount: 0,
    usedIds: new Set(),
  };
}

function collectTokens(root: HTMLElement): Token[] {
  const tokens: Token[] = [];
  let section: SectionName = "document";

  Array.from(root.children).forEach((child) => {
    if (child instanceof HTMLTableElement) {
      Array.from(child.rows).forEach((row) => {
        const cells = Array.from(row.cells).map((cell) => cleanText(cell.textContent ?? ""));
        const text = cleanText(cells.join(" "));
        section = sectionFromText(text, section);
        tokens.push({
          text,
          cells,
          images: collectImages(row),
          section,
          tagName: "tr",
        });
      });
      return;
    }

    if (child instanceof HTMLOListElement || child instanceof HTMLUListElement) {
      Array.from(child.querySelectorAll("li")).forEach((item) => {
        const text = cleanText(item.textContent ?? "");
        const markdownText = markdownTextFromElement(item);
        section = sectionFromText(text, section);
        tokens.push({
          text,
          markdownText,
          images: collectImages(item),
          section,
          tagName: "li",
        });
      });
      return;
    }

    const text = cleanText(child.textContent ?? "");
    const markdownText = markdownTextFromElement(child);
    section = sectionFromText(text, section);

    tokens.push({
      text,
      markdownText,
      images: collectImages(child),
      section,
      tagName: child.tagName.toLowerCase(),
    });
  });

  return tokens.filter((token) => token.text || token.images.length > 0);
}

function collectImages(element: Element): EvidenceImage[] {
  return Array.from(element.querySelectorAll("img"))
    .map((image, index) => image.getAttribute("src") ?? "")
    .filter((source) => source.startsWith("data:image/"))
    .map((source, index) => ({
      id: createImportId("image", `${Date.now()}-${index}-${source.slice(0, 24)}`),
      label: "",
      name: `imagem-importada-${index + 1}`,
      dataUrl: source,
      width: 560,
      height: 320,
    }));
}

function markdownTextFromElement(element: Element): string {
  return cleanText(
    Array.from(element.childNodes)
      .map((node) => markdownTextFromNode(node))
      .join(""),
  );
}

function markdownTextFromNode(node: ChildNode, inheritedBold = false): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? "";
  }

  if (!(node instanceof Element)) {
    return "";
  }

  const isBold = node.tagName.toLowerCase() === "strong" || node.tagName.toLowerCase() === "b";
  const childText = Array.from(node.childNodes)
    .map((child) => markdownTextFromNode(child, inheritedBold || isBold))
    .join("");

  if (!isBold || inheritedBold) {
    return childText;
  }

  return wrapMarkdownBold(childText);
}

function wrapMarkdownBold(value: string): string {
  const leadingSpace = value.match(/^\s*/)?.[0] ?? "";
  const trailingSpace = value.match(/\s*$/)?.[0] ?? "";
  const text = value.trim();

  return text ? `${leadingSpace}**${text}**${trailingSpace}` : value;
}

function sectionFromText(text: string, current: SectionName): SectionName {
  const normalized = normalizeText(text);

  if (normalized.includes("passo a passo")) {
    return "steps";
  }

  if (normalized.includes("tipos de permissao") || normalized.includes("permissoes para testar")) {
    return "permissions";
  }

  if (normalized === "testes") {
    return "tests";
  }

  return current;
}

function applyMammothWarnings(
  state: { warnings: string[] },
  messages: Array<{ message?: string; type?: string }> | undefined,
): void {
  messages?.forEach((message) => {
    if (message.message && !isIgnorableMammothWarning(message.message)) {
      state.warnings.push(`Conversao DOCX: ${message.message}`);
    }
  });
}

function isIgnorableMammothWarning(message: string): boolean {
  return message === "Unrecognised paragraph style: 'Title' (Style ID: Title)";
}

function parseDocumentFields(tokens: Token[], documentData: OtDocument): void {
  tokens.forEach((token) => {
    if (token.cells && token.cells.length >= 2) {
      const [label, ...rest] = token.cells;
      assignLabeledField(documentData, label, rest.join(" "));
      return;
    }

    const field = splitLabelValue(token.text);
    if (field) {
      assignLabeledField(documentData, field.label, field.value);
    }
  });
}

function assignLabeledField(documentData: OtDocument, label: string, rawValue: string): void {
  const normalizedLabel = normalizeText(label);
  const value = cleanText(rawValue);

  if (!value) {
    return;
  }

  if (normalizedLabel.includes("objetivo")) {
    documentData.objective = value;
    return;
  }

  metadataFields.forEach(({ field, aliases }) => {
    if (aliases.some((alias) => normalizedLabel.includes(alias))) {
      documentData.metadata[field] = field === "date" ? normalizeDate(value) : value;
    }
  });
}

function parseAccessSteps(tokens: Token[], documentData: OtDocument): void {
  const steps: string[] = [];

  tokens
    .filter((token) => token.section === "steps")
    .forEach((token) => {
      if (isSectionHeading(token.text)) {
        return;
      }

      if (token.cells && token.cells.length >= 2 && /^\d+\.?$/.test(token.cells[0])) {
        addStepText(steps, token.cells.slice(1).join(" "));
        return;
      }

      splitPossibleSteps(token.text).forEach((step) => addStepText(steps, step));
    });

  documentData.accessSteps = uniqueTexts(steps).map((text, index) => ({
    id: `step-import-${index + 1}`,
    text,
  }));
}

function parseTeaTokens(tokens: Token[], state: TeaParserState): void {
  tokens.forEach((token) => {
    if (token.cells && token.cells.length >= 2 && assignTeaMetadata(token, state.document)) {
      return;
    }

    if (parseTeaMainSection(token, state)) {
      return;
    }

    if (parseTeaSubActivityHeading(token, state)) {
      appendTeaImagesToCurrentTarget(state, token.images);
      return;
    }

    if (parseTeaActivityHeading(token, state)) {
      appendTeaImagesToCurrentTarget(state, token.images);
      return;
    }

    appendTeaContentToken(state, token);
  });
}

function assignTeaMetadata(token: Token, documentData: TeaDocument): boolean {
  const [label = "", ...rest] = token.cells ?? [];

  return assignTeaMetadataField(documentData, label, rest.join(" "));
}

function assignTeaMetadataField(
  documentData: TeaDocument,
  label: string,
  rawValue: string,
): boolean {
  const normalizedLabel = normalizeText(label);
  const value = cleanText(rawValue);

  if (!value) {
    return false;
  }

  if (normalizedLabel.includes("ordem de servico")) {
    documentData.metadata.serviceOrder = value;
    return true;
  }

  if (normalizedLabel.includes("fase") || normalizedLabel.includes("etapa")) {
    documentData.metadata.phase = value;
    return true;
  }

  if (normalizedLabel.includes("chamado")) {
    documentData.metadata.ticket = value;
    return true;
  }

  if (normalizedLabel.includes("assunto")) {
    documentData.metadata.subject = value;
    return true;
  }

  if (normalizedLabel.includes("data")) {
    documentData.metadata.date = normalizeDate(value);
    return true;
  }

  if (normalizedLabel.includes("elaborado por")) {
    documentData.metadata.author = value;
    return true;
  }

  return false;
}

function parseTeaMainSection(token: Token, state: TeaParserState): boolean {
  const normalized = normalizeText(token.text);

  if (/^1\.?\s+visao geral/.test(normalized)) {
    state.currentSection = "overview";
    state.currentActivity = undefined;
    state.currentSubActivity = undefined;
    return true;
  }

  if (/^2\.?\s+atividades realizadas/.test(normalized)) {
    state.currentSection = "activityIntro";
    state.currentActivity = undefined;
    state.currentSubActivity = undefined;
    return true;
  }

  return false;
}

function parseTeaActivityHeading(token: Token, state: TeaParserState): boolean {
  const heading = parseTeaNumberedHeading(token.text);

  if (!heading || heading.level !== "activity") {
    return false;
  }

  const activity: TeaActivity = {
    id: createTeaImportId(
      state,
      "tea-activity",
      `${state.document.activities.length + 1}-${heading.title}`,
    ),
    title: heading.title,
    blocks: [],
    subActivities: [],
  };

  state.document.activities.push(activity);
  state.currentSection = "activities";
  state.currentActivity = activity;
  state.currentSubActivity = undefined;

  return true;
}

function parseTeaSubActivityHeading(token: Token, state: TeaParserState): boolean {
  const heading = parseTeaNumberedHeading(token.text);

  if (!heading || heading.level !== "subActivity") {
    return false;
  }

  const activity =
    state.currentActivity ??
    ensureTeaActivity(state, `Atividade ${heading.activityIndex}`);
  const subActivity: TeaSubActivity = {
    id: createTeaImportId(
      state,
      "tea-subactivity",
      `${activity.id}-${activity.subActivities.length + 1}-${heading.title}`,
    ),
    title: heading.title,
    blocks: [],
  };

  activity.subActivities.push(subActivity);
  state.currentSection = "activities";
  state.currentActivity = activity;
  state.currentSubActivity = subActivity;

  return true;
}

function parseTeaNumberedHeading(
  text: string,
):
  | { level: "activity"; activityIndex: number; title: string }
  | { level: "subActivity"; activityIndex: number; subIndex: number; title: string }
  | null {
  const cleaned = cleanText(text);
  const subActivity = cleaned.match(/^2\.(\d+)\.(\d+)\s*[-\u2013\u2014]\s*(.+)$/);

  if (subActivity) {
    return {
      level: "subActivity",
      activityIndex: Number(subActivity[1]),
      subIndex: Number(subActivity[2]),
      title: cleanTeaHeadingTitle(subActivity[3]),
    };
  }

  const activity = cleaned.match(/^2\.(\d+)\s*[-\u2013\u2014]\s*(.+)$/);

  if (activity) {
    return {
      level: "activity",
      activityIndex: Number(activity[1]),
      title: cleanTeaHeadingTitle(activity[2]),
    };
  }

  return null;
}

function cleanTeaHeadingTitle(value: string): string {
  return cleanText(value).replace(/:$/, "").trim();
}

function appendTeaContentToken(state: TeaParserState, token: Token): void {
  const text = cleanText(token.markdownText ?? token.text);

  if (state.currentSection === "overview") {
    appendTeaDocumentText(state.document, "overview", text);
    return;
  }

  if (state.currentSection === "activityIntro") {
    appendTeaDocumentText(state.document, "activityIntro", text);
    state.document.activityImages.push(...createTeaImages(state, token.images));
    return;
  }

  if (state.currentSection !== "activities") {
    return;
  }

  ensureTeaContentTarget(state);

  if (token.tagName === "li") {
    appendTeaListItem(state, text);
  } else {
    appendTeaTextBlock(state, text);
  }

  appendTeaImagesToCurrentTarget(state, token.images);
}

function appendTeaDocumentText(
  documentData: TeaDocument,
  field: "overview" | "activityIntro",
  text: string,
): void {
  if (!text) {
    return;
  }

  documentData[field] = [documentData[field], text].filter(Boolean).join("\n");
}

function appendTeaTextBlock(state: TeaParserState, text: string): void {
  if (!text) {
    return;
  }

  getTeaCurrentBlocks(state).push({
    id: createTeaBlockId(state, "text"),
    type: "text",
    text,
  });
}

function appendTeaListItem(state: TeaParserState, text: string): void {
  if (!text) {
    return;
  }

  const blocks = getTeaCurrentBlocks(state);
  const lastBlock = blocks[blocks.length - 1];
  const item = {
    id: createTeaItemId(state, text),
    text,
  };

  if (lastBlock?.type === "list") {
    lastBlock.items.push(item);
    return;
  }

  blocks.push({
    id: createTeaBlockId(state, "list"),
    type: "list",
    items: [item],
  });
}

function appendTeaImagesToCurrentTarget(
  state: TeaParserState,
  images: EvidenceImage[],
): void {
  if (images.length === 0) {
    return;
  }

  ensureTeaContentTarget(state);

  getTeaCurrentBlocks(state).push({
    id: createTeaBlockId(state, "images"),
    type: "images",
    images: createTeaImages(state, images),
  });
}

function createTeaImages(state: TeaParserState, images: EvidenceImage[]): EvidenceImage[] {
  return images.map((image) => {
    state.imageCount += 1;

    return {
      ...image,
      id: createTeaImportId(
        state,
        "tea-image",
        `${state.imageCount}-${image.dataUrl ?? image.name}`,
      ),
      name: `imagem-importada-${state.imageCount}`,
      label: image.label ?? "",
    };
  });
}

function ensureTeaContentTarget(state: TeaParserState): TeaActivity {
  return state.currentActivity ?? ensureTeaActivity(state, "Atividade importada");
}

function ensureTeaActivity(state: TeaParserState, title: string): TeaActivity {
  const activity: TeaActivity = {
    id: createTeaImportId(
      state,
      "tea-activity",
      `${state.document.activities.length + 1}-${title}`,
    ),
    title,
    blocks: [],
    subActivities: [],
  };

  state.document.activities.push(activity);
  state.currentSection = "activities";
  state.currentActivity = activity;
  state.currentSubActivity = undefined;

  return activity;
}

function getTeaCurrentBlocks(state: TeaParserState): TeaContentBlock[] {
  return (state.currentSubActivity ?? state.currentActivity ?? ensureTeaContentTarget(state))
    .blocks;
}

function createTeaBlockId(state: TeaParserState, type: TeaContentBlock["type"]): string {
  state.blockCount += 1;

  return createTeaImportId(state, "tea-block", `${state.blockCount}-${type}`);
}

function createTeaItemId(state: TeaParserState, text: string): string {
  state.itemCount += 1;

  return createTeaImportId(state, "tea-item", `${state.itemCount}-${text}`);
}

function parsePermissionSummary(tokens: Token[], state: ParserState): void {
  const summaryMacros: PermissionDraft[] = [];
  const summaryMicros: PermissionDraft[] = [];
  let reading: "macro" | "micro" | undefined;

  tokens
    .filter((token) => token.section === "permissions")
    .forEach((token) => {
      const normalized = normalizeText(token.text);

      if (normalized.includes("macro-permiss")) {
        reading = "macro";
      }

      if (normalized.includes("micro-permiss")) {
        reading = "micro";
      }

      if (token.cells && token.cells.length >= 2) {
        const label = normalizeText(token.cells[0]);
        const value = token.cells.slice(1).join(" ");

        if (label.includes("macro-permiss")) {
          summaryMacros.push(parsePermission(value));
        }

        if (label.includes("micro-permiss")) {
          summaryMicros.push(parsePermission(value));
        }

        return;
      }

      splitPermissionCandidates(token.text).forEach((candidate) => {
        if (reading === "macro") {
          summaryMacros.push(parsePermission(candidate));
        }

        if (reading === "micro") {
          summaryMicros.push(parsePermission(candidate));
        }
      });
    });

  if (state.document.permissionGroups.length === 0) {
    uniquePermissions(summaryMacros).forEach((macro) => {
      const macroGroup = ensureMacro(state, macro);
      uniquePermissions(summaryMicros).forEach((micro) => {
        ensureMicro(state, macroGroup, micro);
      });
    });
  }
}

function parseTestSection(tokens: Token[], state: ParserState): void {
  tokens
    .filter((token) => token.section === "tests")
    .forEach((token) => {
      if (parsePermissionContext(token, state)) {
        return;
      }

      if (isSectionHeading(token.text)) {
        return;
      }

      const startedTest = parseTestStart(token, state);
      parseChecks(token, state.currentTest);
      parseObservation(token, state.currentTest);
      const parsedError = parseTestError(token, state);
      const parsedCorrection = parseCorrection(token, state);
      parseEvidence(token, state);

      if (
        !startedTest &&
        !parsedError &&
        !parsedCorrection &&
        !state.currentCorrectionEvidence &&
        !state.currentErrorEvidence
      ) {
        parseFreeObservation(token, state.currentTest, state.currentEvidence);
      }
    });
}

function parsePermissionContext(token: Token, state: ParserState): boolean {
  const normalized = normalizeText(token.text);

  if (token.cells && token.cells.length >= 2) {
    const label = normalizeText(token.cells[0]);
    const value = token.cells.slice(1).join(" ");

    if (label.includes("macro-permissao") || normalizeText(value).includes("tipo de usuario")) {
      state.currentMacro = ensureMacro(state, parsePermission(value));
      state.currentMicro = undefined;
      resetCurrentOtTest(state);
      return true;
    }

    if (label.includes("micro-permissao") || normalizeText(value).includes("tipo de permissao")) {
      const macro = state.currentMacro ?? ensureMacro(state, { code: "MACRO", label: "" });
      state.currentMicro = ensureMicro(state, macro, parsePermission(value));
      resetCurrentOtTest(state);
      return true;
    }
  }

  if (normalized.includes("tipo de usuario")) {
    state.currentMacro = ensureMacro(state, parsePermission(token.text));
    state.currentMicro = undefined;
    resetCurrentOtTest(state);
    return true;
  }

  if (normalized.includes("tipo de permissao")) {
    const macro = state.currentMacro ?? ensureMacro(state, { code: "MACRO", label: "" });
    state.currentMicro = ensureMicro(state, macro, parsePermission(token.text));
    resetCurrentOtTest(state);
    return true;
  }

  return false;
}

function resetCurrentOtTest(state: ParserState): void {
  state.currentTest = undefined;
  state.currentError = undefined;
  state.currentEvidence = undefined;
  state.currentErrorEvidence = undefined;
  state.currentCorrectionEvidence = undefined;
  state.pendingImageLabel = undefined;
}

function parseTestStart(token: Token, state: ParserState): boolean {
  const match = token.text.match(/(?:^|\s)(\d{1,3})\s*[-–]\s*([^()]+)/);

  if (!match) {
    return false;
  }

  const title = cleanText(match[2]).replace(/\s{2,}.*/, "").trim();
  const macro = state.currentMacro ?? ensureMacro(state, { code: "MACRO", label: "" });
  const micro = state.currentMicro ?? ensureMicro(state, macro, { code: "MICRO", label: "" });
  const block = ensureBlock(state.document, macro, micro);
  const test: PermissionBlockTest = {
    id: createImportId("test", `${macro.id}-${micro.id}-${block.tests.length + 1}-${title}`),
    title,
    result: createEmptyTestResult(),
    correction: createEmptyTestCorrection(),
  };

  block.tests.push(test);
  state.currentTest = test;
  state.currentError = undefined;
  state.currentEvidence = undefined;
  state.currentErrorEvidence = undefined;
  state.currentCorrectionEvidence = undefined;
  state.pendingImageLabel = undefined;

  return true;
}

function parseChecks(token: Token, test: PermissionBlockTest | undefined): void {
  if (!test) {
    return;
  }

  if (token.cells && token.cells.length >= 2) {
    const marker = token.cells[0];
    const label = token.cells.slice(1).join(" ");
    const key = checkKeyFromText(label);

    if (key) {
      test.result.checks[key] = hasCheckedMarker(marker);
      return;
    }
  }

  checkOrder.forEach((key) => {
    const normalizedText = normalizeText(token.text);
    const matchingLabel = checkLabelAliases[key].find((label) =>
      normalizedText.includes(normalizeText(label)),
    );

    if (!matchingLabel) {
      return;
    }

    const labelIndex = normalizedText.indexOf(normalizeText(matchingLabel));
    const beforeLabel = token.text.slice(Math.max(0, labelIndex - 18), labelIndex);
    test.result.checks[key] = hasCheckedMarker(beforeLabel);
  });
}

function parseObservation(token: Token, test: PermissionBlockTest | undefined): void {
  if (!test || !token.cells || token.cells.length < 2) {
    return;
  }

  const label = normalizeText(token.cells[0]);

  if (label.includes("observacoes") || label.includes("observacao")) {
    appendObservation(test, token.cells.slice(1).join(" "));
  }
}

function parseTestError(token: Token, state: ParserState): boolean {
  if (!state.currentTest) {
    return false;
  }

  const normalized = normalizeText(token.text);

  if (normalized === "erros encontrados" || normalized === "erros encontrados:") {
    state.currentEvidence = undefined;
    state.currentErrorEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = undefined;
    return true;
  }

  const headingMatch = token.text.match(/^\s*Erro\s+\d+\s*[-–]\s*(.+)$/i);

  if (headingMatch) {
    const origin = parseTestErrorOrigin(headingMatch[1]);
    const error = createEmptyTestError(
      createImportId("error", `${state.currentTest.id}-${state.currentTest.result.errors.length + 1}`),
      origin,
    );

    state.currentTest.result.errors.push(error);
    state.currentError = error;
    state.currentEvidence = undefined;
    state.currentErrorEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = undefined;
    return true;
  }

  if (!state.currentError || !token.cells || token.cells.length < 2) {
    return false;
  }

  const label = normalizeText(token.cells[0]);
  const value = cleanText(token.cells.slice(1).join(" "));

  if (label.includes("origem")) {
    state.currentError.origin = parseTestErrorOrigin(value);
    return true;
  }

  if (label.includes("observacao") || label.includes("observacoes")) {
    state.currentError.observation = value;
    return true;
  }

  return false;
}

function parseCorrection(token: Token, state: ParserState): boolean {
  if (!state.currentTest) {
    return false;
  }

  const normalized = normalizeText(token.text);
  const correction = state.currentError?.correction ?? getImportCorrection(state.currentTest);

  if (normalized === "correcao:" || normalized === "correcao") {
    state.currentEvidence = undefined;
    state.currentErrorEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = undefined;
    return true;
  }

  if (token.cells && token.cells.length >= 2) {
    const label = normalizeText(token.cells[0]);
    const value = cleanText(token.cells.slice(1).join(" "));

    if (label.includes("corrigido por")) {
      correction.correctedBy = value;
      return true;
    }

    if (label.includes("corrigido")) {
      correction.corrected = parseBooleanAnswer(value);
      return true;
    }

    if (label.includes("hotfix")) {
      correction.hotfixTag = value;
      return true;
    }

    if (label.includes("nuvem")) {
      correction.cloudStage = parseCloudStageText(value);
      return true;
    }
  }

  return isCorrectionEvidenceHeading(token.text);
}

function parseEvidence(token: Token, state: ParserState): void {
  const normalized = normalizeText(token.text);

  if (isCorrectionBeforeHeading(token.text)) {
    state.currentCorrectionEvidence = "beforeImages";
    state.currentEvidence = undefined;
    state.currentErrorEvidence = undefined;
    state.pendingImageLabel = extractCorrectionEvidenceLabel(token.text, "antes");
  } else if (isCorrectionAfterHeading(token.text)) {
    state.currentCorrectionEvidence = "afterImages";
    state.currentEvidence = undefined;
    state.currentErrorEvidence = undefined;
    state.pendingImageLabel = extractCorrectionEvidenceLabel(token.text, "depois");
  } else if (isErrorEvidenceHeading(token.text)) {
    state.currentErrorEvidence = true;
    state.currentEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = extractErrorEvidenceLabel(token.text);
  } else if (normalized.includes("legado")) {
    state.currentEvidence = "legacyImages";
    state.currentErrorEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = extractEvidenceLabel(token.text, "legado");
  } else if (normalized.includes("novo")) {
    state.currentEvidence = "newImages";
    state.currentErrorEvidence = undefined;
    state.currentCorrectionEvidence = undefined;
    state.pendingImageLabel = extractEvidenceLabel(token.text, "novo");
  }

  if (!state.currentTest || token.images.length === 0) {
    if (
      (state.currentEvidence || state.currentCorrectionEvidence) &&
      token.text &&
      !isEvidenceHeading(token.text)
    ) {
      state.pendingImageLabel = token.text;
    }
    return;
  }

  if (state.currentCorrectionEvidence) {
    const field = state.currentCorrectionEvidence;
    const correction = state.currentError?.correction ?? getImportCorrection(state.currentTest);
    const label = extractCorrectionEvidenceLabel(
      token.text,
      field === "beforeImages" ? "antes" : "depois",
    );

    token.images.forEach((image) => {
      state.imageCount += 1;
      correction[field].push({
        ...image,
        id: createImportId("image", `${state.imageCount}-${image.dataUrl ?? image.name}`),
        name: `imagem-importada-${state.imageCount}`,
        label: label || state.pendingImageLabel || "",
      });
    });

    state.pendingImageLabel = undefined;
    return;
  }

  if (state.currentErrorEvidence && state.currentError) {
    const label = extractErrorEvidenceLabel(token.text);

    token.images.forEach((image) => {
      state.imageCount += 1;
      state.currentError?.images.push({
        ...image,
        id: createImportId("image", `${state.imageCount}-${image.dataUrl ?? image.name}`),
        name: `imagem-importada-${state.imageCount}`,
        label: label || state.pendingImageLabel || "",
      });
    });

    state.pendingImageLabel = undefined;
    return;
  }

  if (!state.currentEvidence) {
    return;
  }

  const label = extractEvidenceLabel(
    token.text,
    state.currentEvidence === "legacyImages" ? "legado" : "novo",
  );

  token.images.forEach((image) => {
    state.imageCount += 1;
    state.currentTest?.result[state.currentEvidence as "legacyImages" | "newImages"].push({
      ...image,
      id: createImportId("image", `${state.imageCount}-${image.dataUrl ?? image.name}`),
      name: `imagem-importada-${state.imageCount}`,
      label: label || state.pendingImageLabel || "",
    });
  });

  state.pendingImageLabel = undefined;
}

function parseFreeObservation(
  token: Token,
  test: PermissionBlockTest | undefined,
  evidence: "legacyImages" | "newImages" | undefined,
): void {
  if (!test || evidence || token.images.length > 0 || token.cells) {
    return;
  }

  const text = cleanText(
    token.text
      .replace(/\(\s*(?:x)?\s*\)/gi, "")
      .replace(/Legado:?.*/i, "")
      .replace(/Novo:?.*/i, ""),
  );

  if (
    text &&
    !isSectionHeading(text) &&
    !checkKeyFromText(text) &&
    !/\d{1,3}\s*[-–]/.test(text)
  ) {
    appendObservation(test, text);
  }
}

function getImportCorrection(test: PermissionBlockTest): TestCorrection {
  test.correction = {
    ...createEmptyTestCorrection(),
    ...test.correction,
    beforeImages: test.correction?.beforeImages ?? [],
    afterImages: test.correction?.afterImages ?? [],
  };

  return test.correction;
}

function parseBooleanAnswer(value: string): boolean {
  const normalized = normalizeText(value);

  if (normalized.includes("sim") || hasCheckedMarker(value)) {
    return true;
  }

  if (normalized.includes("nao") || normalized.includes("não")) {
    return false;
  }

  return false;
}

function parseCloudStageText(value: string): TestCorrection["cloudStage"] {
  const normalized = normalizeText(value);

  if (normalized.includes("producao") || normalized.includes("production")) {
    return "production";
  }

  if (normalized.includes("homolog")) {
    return "homolog";
  }

  if (normalized.includes("dev") || normalized.includes("desenvolvimento")) {
    return "dev";
  }

  return "none";
}

function parseTestErrorOrigin(value: string): TestError["origin"] {
  const normalized = normalizeText(value);

  if (normalized.includes("legado")) {
    return "legacy";
  }

  return "new";
}

function isCorrectionEvidenceHeading(text: string): boolean {
  return isCorrectionBeforeHeading(text) || isCorrectionAfterHeading(text);
}

function isErrorEvidenceHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return (
    normalized === "prints do erro" ||
    normalized === "print do erro" ||
    normalized.startsWith("prints do erro ") ||
    normalized.startsWith("print do erro ") ||
    normalized === "evidencias do erro" ||
    normalized.startsWith("evidencias do erro ")
  );
}

function isCorrectionBeforeHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return normalized === "antes" || normalized.startsWith("antes ");
}

function isCorrectionAfterHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return (
    normalized === "depois" ||
    normalized.startsWith("depois ") ||
    normalized === "corrigido" ||
    normalized.startsWith("corrigido ")
  );
}

function finishDocument(state: ParserState, sourceName: string): void {
  if (!state.document.metadata.screen) {
    state.document.metadata.screen = screenNameFromFile(sourceName);
    state.warnings.push("Tela nao encontrada no DOCX; usei o nome do arquivo.");
  }

  if (!state.document.objective) {
    state.warnings.push("Objetivo nao encontrado no DOCX.");
  }

  if (state.document.accessSteps.length === 0) {
    state.warnings.push("Nenhum passo de acesso reconhecido.");
  }

  if (selectedPermissionCount(state.document.permissionGroups) === 0) {
    state.warnings.push("Nenhuma permissao reconhecida.");
  }

  if (buildImportSummary(state.document).tests === 0) {
    state.warnings.push("Nenhum teste reconhecido.");
  }
}

function finishTeaDocument(state: TeaParserState, sourceName: string): void {
  if (!state.document.metadata.subject) {
    state.document.metadata.subject = teaSubjectFromFile(sourceName);
    state.warnings.push("Assunto nao encontrado no DOCX; usei o nome do arquivo.");
  }

  if (!state.document.metadata.serviceOrder) {
    state.warnings.push("Ordem de servico nao encontrada no DOCX.");
  }

  if (!state.document.metadata.author) {
    state.warnings.push("Elaborado por nao encontrado no DOCX.");
  }

  if (!state.document.overview.trim()) {
    state.warnings.push("Visao geral nao encontrada no DOCX.");
  }

  if (!state.document.activityIntro.trim()) {
    state.warnings.push("Texto inicial de atividades nao encontrado no DOCX.");
  }

  if (state.document.activities.length === 0) {
    state.warnings.push("Nenhuma atividade reconhecida.");
  }
}

function ensureMacro(state: ParserState, permission: PermissionDraft): PermissionGroup {
  const parsed = normalizePermissionDraft(permission, "MACRO");
  const existing = state.document.permissionGroups.find(
    (group) => normalizeText(group.code) === normalizeText(parsed.code),
  );

  if (existing) {
    if (!existing.label && parsed.label) {
      existing.label = parsed.label;
    }
    return existing;
  }

  const group: PermissionGroup = {
    id: createImportId("macro", parsed.code),
    code: parsed.code,
    label: parsed.label,
    selected: true,
    microPermissions: [],
  };

  state.document.permissionGroups.push(group);
  return group;
}

function ensureMicro(
  state: ParserState,
  macro: PermissionGroup,
  permission: PermissionDraft,
): PermissionItem {
  const parsed = normalizePermissionDraft(permission, "MICRO");
  const existing = macro.microPermissions.find(
    (micro) => normalizeText(micro.code) === normalizeText(parsed.code),
  );

  if (existing) {
    if (!existing.label && parsed.label) {
      existing.label = parsed.label;
    }
    ensureBlock(state.document, macro, existing);
    return existing;
  }

  const micro: PermissionItem = {
    id: createImportId("micro", `${macro.code}-${parsed.code}`),
    code: parsed.code,
    label: parsed.label,
    selected: true,
  };

  macro.microPermissions.push(micro);
  ensureBlock(state.document, macro, micro);
  return micro;
}

function ensureBlock(
  documentData: OtDocument,
  macro: PermissionGroup,
  micro: PermissionItem,
): PermissionBlock {
  const key = createPermissionKey(macro.id, micro.id);

  documentData.permissionBlocks[key] ??= { tests: [] };
  return documentData.permissionBlocks[key];
}

function parsePermission(value: string): PermissionDraft {
  const withoutLabel = value
    .replace(/tipo de usu[aá]rio:*/i, "")
    .replace(/tipo de permiss[aã]o:*/i, "")
    .trim();
  const match = withoutLabel.match(/([A-Za-z0-9_-]+)\s*(?:\(([^)]+)\))?/);

  if (!match) {
    return { code: withoutLabel || "PERM", label: "" };
  }

  return {
    code: match[1],
    label: cleanText(match[2] ?? ""),
  };
}

function splitPermissionCandidates(text: string): string[] {
  return cleanText(text)
    .split(/\n|•|;|,/)
    .map((value) => value.replace(/^(macro|micro)-permiss[oõ]es?:?/i, "").trim())
    .filter((value) => /^[A-Za-z0-9_-]+(?:\s*\([^)]+\))?$/.test(value));
}

function uniquePermissions(permissions: PermissionDraft[]): PermissionDraft[] {
  const seen = new Set<string>();

  return permissions.filter((permission) => {
    const key = normalizeText(permission.code);
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizePermissionDraft(permission: PermissionDraft, fallbackCode: string): PermissionDraft {
  return {
    code: cleanText(permission.code || fallbackCode).toUpperCase(),
    label: cleanText(permission.label),
  };
}

function splitPossibleSteps(text: string): string[] {
  const cleaned = cleanText(text);

  if (!cleaned || isSectionHeading(cleaned)) {
    return [];
  }

  const numberedMatches = Array.from(
    cleaned.matchAll(/\d{1,2}[.)-]\s*(.*?)(?=\s+\d{1,2}[.)-]\s*|$)/g),
  );

  if (numberedMatches.length > 0) {
    return numberedMatches.map((match) => cleanText(match[1]));
  }

  return [cleaned.replace(/^\d{1,2}[.)-]\s*/, "")];
}

function addStepText(steps: string[], text: string): void {
  const cleaned = cleanText(text).replace(/^\d{1,2}[.)-]\s*/, "");

  if (cleaned && !isSectionHeading(cleaned)) {
    steps.push(cleaned);
  }
}

function splitLabelValue(text: string): { label: string; value: string } | undefined {
  const match = text.match(/^([^:]{3,60}):\s*(.+)$/);

  return match ? { label: match[1], value: match[2] } : undefined;
}

function checkKeyFromText(text: string): CheckKey | undefined {
  const normalized = normalizeText(text);

  return checkOrder.find((key) =>
    checkLabelAliases[key].some((label) => normalized.includes(normalizeText(label))),
  );
}

function hasCheckedMarker(text: string): boolean {
  return /\(\s*(?:x|X|✓|✔)\s*\)/.test(text) || /\bX\b/.test(text);
}

function appendObservation(test: PermissionBlockTest, value: string): void {
  const text = cleanText(value);

  if (!text) {
    return;
  }

  test.result.observations = cleanText(
    [test.result.observations, text].filter(Boolean).join("\n"),
  );
}

function extractEvidenceLabel(text: string, heading: "legado" | "novo"): string {
  const cleaned = cleanText(text.replace(new RegExp(`${heading}:?`, "i"), ""));

  if (!cleaned || normalizeText(cleaned) === heading) {
    return "";
  }

  return cleaned;
}

function extractCorrectionEvidenceLabel(text: string, heading: "antes" | "depois"): string {
  const normalized = normalizeText(text);
  const cleaned = cleanText(
    normalized === heading || normalized.startsWith(`${heading} `)
      ? text
          .replace(new RegExp(`^${heading}\\s*(?:\\([^)]*\\))?:?`, "i"), "")
          .replace(/^corrigido:?/i, "")
      : text,
  );

  if (!cleaned || isCorrectionEvidenceHeading(cleaned)) {
    return "";
  }

  return cleaned;
}

function extractErrorEvidenceLabel(text: string): string {
  const cleaned = cleanText(text.replace(/^prints?\s+do\s+erro\s*:?\s*/i, ""));

  if (!cleaned || isErrorEvidenceHeading(cleaned)) {
    return "";
  }

  return cleaned;
}

function isEvidenceHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return (
    normalized === "legado" ||
    normalized === "legado:" ||
    normalized === "novo" ||
    normalized === "novo:" ||
    isErrorEvidenceHeading(text) ||
    isCorrectionEvidenceHeading(text)
  );
}

function isSectionHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return (
    normalized.includes("passo a passo") ||
    normalized.includes("tipos de permissao") ||
    normalized === "testes" ||
    normalized.includes("macro-permissao") ||
    normalized.includes("micro-permissao")
  );
}

function buildImportSummary(documentData: OtDocument): OtDocxImportSummary {
  const blocks = Object.values(documentData.permissionBlocks);
  const tests = blocks.reduce((total, block) => total + block.tests.length, 0);
  const images = blocks.reduce(
    (total, block) =>
      total +
      block.tests.reduce(
        (testTotal, test) =>
          testTotal +
          test.result.legacyImages.length +
          test.result.newImages.length +
          (test.correction?.beforeImages.length ?? 0) +
          (test.correction?.afterImages.length ?? 0),
        0,
      ),
    0,
  );

  return {
    kind: "ot",
    screen: documentData.metadata.screen,
    accessSteps: documentData.accessSteps.length,
    permissionGroups: documentData.permissionGroups.length,
    selectedPermissions: selectedPermissionCount(documentData.permissionGroups),
    tests,
    images,
  };
}

function buildTeaImportSummary(documentData: TeaDocument): TeaDocxImportSummary {
  const blocks = getTeaContentBlocks(documentData);

  return {
    kind: "tea",
    subject: documentData.metadata.subject,
    activities: documentData.activities.length,
    subActivities: documentData.activities.reduce(
      (total, activity) => total + activity.subActivities.length,
      0,
    ),
    blocks: blocks.length,
    images:
      documentData.activityImages.length +
      blocks.reduce(
        (total, block) => total + (block.type === "images" ? block.images.length : 0),
        0,
      ),
  };
}

function getTeaContentBlocks(documentData: TeaDocument): TeaContentBlock[] {
  return documentData.activities.flatMap((activity) => [
    ...activity.blocks,
    ...activity.subActivities.flatMap((subActivity) => subActivity.blocks),
  ]);
}

function selectedPermissionCount(groups: PermissionGroup[]): number {
  return groups.reduce(
    (total, group) =>
      total + (group.selected ? group.microPermissions.filter((micro) => micro.selected).length : 0),
    0,
  );
}

function uniqueTexts(values: string[]): string[] {
  const seen = new Set<string>();

  return values.filter((value) => {
    const key = normalizeText(value);
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizeDate(value: string): string {
  const trimmed = value.trim();
  const brazilian = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (brazilian) {
    const [, day, month, year] = brazilian;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }

  return trimmed;
}

function screenNameFromFile(fileName: string): string {
  return fileName
    .replace(/\.docx$/i, "")
    .replace(/^OT\s*-\s*/i, "")
    .trim();
}

function teaSubjectFromFile(fileName: string): string {
  return fileName
    .replace(/\.docx$/i, "")
    .replace(/^TEA\s*-\s*/i, "")
    .trim();
}

function createTeaImportId(state: TeaParserState, prefix: string, value: string): string {
  const baseId = createImportId(prefix, value);
  let id = baseId;
  let suffix = 2;

  while (state.usedIds.has(id)) {
    id = `${baseId}-${suffix}`;
    suffix += 1;
  }

  state.usedIds.add(id);
  return id;
}

function createImportId(prefix: string, value: string): string {
  const slug = normalizeText(value)
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);

  return `${prefix}-${slug || "item"}`;
}

function cleanText(value: string): string {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeText(value: string): string {
  return cleanText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}
