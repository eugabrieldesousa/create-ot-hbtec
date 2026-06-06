import mammoth from "mammoth";
import {
  checkLabels,
  checkOrder,
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
} from "./types";

export type DocxImportSummary = {
  screen: string;
  accessSteps: number;
  permissionGroups: number;
  selectedPermissions: number;
  tests: number;
  images: number;
};

export type DocxImportResult = {
  document: OtDocument;
  summary: DocxImportSummary;
  warnings: string[];
  sourceName: string;
};

type ParseOptions = {
  sourceName?: string;
  mammothMessages?: Array<{ message?: string; type?: string }>;
};

type Token = {
  text: string;
  cells?: string[];
  images: EvidenceImage[];
  section: SectionName;
};

type SectionName = "document" | "steps" | "permissions" | "tests" | "other";

type PermissionDraft = Pick<PermissionItem, "code" | "label">;

type ParserState = {
  document: OtDocument;
  warnings: string[];
  currentMacro?: PermissionGroup;
  currentMicro?: PermissionItem;
  currentTest?: PermissionBlockTest;
  currentEvidence?: "legacyImages" | "newImages";
  pendingImageLabel?: string;
  imageCount: number;
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

export async function parseDocxFile(file: File): Promise<DocxImportResult> {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });

  return parseOtHtml(result.value, {
    sourceName: file.name,
    mammothMessages: result.messages,
  });
}

export function parseOtHtml(html: string, options: ParseOptions = {}): DocxImportResult {
  const dom = new DOMParser().parseFromString(html, "text/html");
  const tokens = collectTokens(dom.body);
  const state = createParserState();

  applyMammothWarnings(state, options.mammothMessages);
  parseDocumentFields(tokens, state.document);
  parseAccessSteps(tokens, state.document);
  parsePermissionSummary(tokens, state);
  parseTestSection(tokens, state);
  finishDocument(state, options.sourceName ?? "documento-importado.docx");

  return {
    document: state.document,
    summary: buildImportSummary(state.document),
    warnings: state.warnings,
    sourceName: options.sourceName ?? "documento-importado.docx",
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
        });
      });
      return;
    }

    if (child instanceof HTMLOListElement || child instanceof HTMLUListElement) {
      Array.from(child.querySelectorAll("li")).forEach((item) => {
        const text = cleanText(item.textContent ?? "");
        section = sectionFromText(text, section);
        tokens.push({
          text,
          images: collectImages(item),
          section,
        });
      });
      return;
    }

    const text = cleanText(child.textContent ?? "");
    section = sectionFromText(text, section);

    tokens.push({
      text,
      images: collectImages(child),
      section,
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
  state: ParserState,
  messages: Array<{ message?: string; type?: string }> | undefined,
): void {
  messages?.forEach((message) => {
    if (message.message) {
      state.warnings.push(`Conversao DOCX: ${message.message}`);
    }
  });
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
      parseEvidence(token, state);

      if (!startedTest) {
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
      state.currentTest = undefined;
      return true;
    }

    if (label.includes("micro-permissao") || normalizeText(value).includes("tipo de permissao")) {
      const macro = state.currentMacro ?? ensureMacro(state, { code: "MACRO", label: "" });
      state.currentMicro = ensureMicro(state, macro, parsePermission(value));
      state.currentTest = undefined;
      return true;
    }
  }

  if (normalized.includes("tipo de usuario")) {
    state.currentMacro = ensureMacro(state, parsePermission(token.text));
    state.currentMicro = undefined;
    state.currentTest = undefined;
    return true;
  }

  if (normalized.includes("tipo de permissao")) {
    const macro = state.currentMacro ?? ensureMacro(state, { code: "MACRO", label: "" });
    state.currentMicro = ensureMicro(state, macro, parsePermission(token.text));
    state.currentTest = undefined;
    return true;
  }

  return false;
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
  };

  block.tests.push(test);
  state.currentTest = test;
  state.currentEvidence = undefined;
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
    const label = checkLabels[key];
    const labelIndex = normalizeText(token.text).indexOf(normalizeText(label));

    if (labelIndex === -1) {
      return;
    }

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

function parseEvidence(token: Token, state: ParserState): void {
  const normalized = normalizeText(token.text);

  if (normalized.includes("legado")) {
    state.currentEvidence = "legacyImages";
    state.pendingImageLabel = extractEvidenceLabel(token.text, "legado");
  }

  if (normalized.includes("novo")) {
    state.currentEvidence = "newImages";
    state.pendingImageLabel = extractEvidenceLabel(token.text, "novo");
  }

  if (!state.currentTest || !state.currentEvidence || token.images.length === 0) {
    if (state.currentEvidence && token.text && !isEvidenceHeading(token.text)) {
      state.pendingImageLabel = token.text;
    }
    return;
  }

  const label = extractEvidenceLabel(token.text, state.currentEvidence === "legacyImages" ? "legado" : "novo");

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

  return checkOrder.find((key) => normalized.includes(normalizeText(checkLabels[key])));
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

function isEvidenceHeading(text: string): boolean {
  const normalized = normalizeText(text);

  return normalized === "legado" || normalized === "legado:" || normalized === "novo" || normalized === "novo:";
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

function buildImportSummary(documentData: OtDocument): DocxImportSummary {
  const blocks = Object.values(documentData.permissionBlocks);
  const tests = blocks.reduce((total, block) => total + block.tests.length, 0);
  const images = blocks.reduce(
    (total, block) =>
      total +
      block.tests.reduce(
        (testTotal, test) =>
          testTotal + test.result.legacyImages.length + test.result.newImages.length,
        0,
      ),
    0,
  );

  return {
    screen: documentData.metadata.screen,
    accessSteps: documentData.accessSteps.length,
    permissionGroups: documentData.permissionGroups.length,
    selectedPermissions: selectedPermissionCount(documentData.permissionGroups),
    tests,
    images,
  };
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
