import {
  createDefaultDocument,
  createEmptyTestCorrection,
  createEmptyTestResult,
  createPermissionKey,
} from "./defaultDocument";
import { createDefaultTeaDocument } from "./defaultTeaDocument";
import {
  deleteEvidenceImageDataBatch,
  stripImageDataFromDocument,
  stripImageDataFromTeaDocument,
} from "./imageStorage";
import type {
  AccessStep,
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
  TeaTextItem,
  TestCorrection,
  TestError,
  TestErrorOrigin,
  TestResult,
} from "./types";

const STORAGE_KEY = "create-ot-draft-v3";
const TEA_STORAGE_KEY = "create-tea-draft-v1";
const LEGACY_STORAGE_KEYS = ["create-ot-draft-v2"];

type LegacyTestDefinition = {
  id: string;
  title: string;
};

type LegacyPermissionBlock = {
  results?: Record<string, TestResult>;
  tests?: PermissionBlockTest[];
};

type DraftCandidate = Partial<OtDocument> & {
  macroPermissions?: PermissionItem[];
  microPermissions?: PermissionItem[];
  tests?: LegacyTestDefinition[];
  permissionBlocks?: Record<string, LegacyPermissionBlock>;
};

export function loadDraft(): OtDocument {
  const fallback = createDefaultDocument();

  try {
    const saved = findSavedDraft();
    if (!saved) {
      return fallback;
    }

    return normalizeDraft(JSON.parse(saved), fallback);
  } catch {
    return fallback;
  }
}

export function saveDraft(document: OtDocument): void {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(stripImageDataFromDocument(document)));
}

export async function clearDraft(): Promise<void> {
  const imageIds = getOtImageIds(loadDraft());

  window.localStorage.removeItem(STORAGE_KEY);
  LEGACY_STORAGE_KEYS.forEach((key) => window.localStorage.removeItem(key));

  await deleteEvidenceImages(imageIds);
}

export function loadTeaDraft(): TeaDocument {
  const fallback = createDefaultTeaDocument();

  try {
    const saved = window.localStorage.getItem(TEA_STORAGE_KEY);
    if (!saved) {
      return fallback;
    }

    return normalizeTeaDraft(JSON.parse(saved), fallback);
  } catch {
    return fallback;
  }
}

export function saveTeaDraft(document: TeaDocument): void {
  window.localStorage.setItem(
    TEA_STORAGE_KEY,
    JSON.stringify(stripImageDataFromTeaDocument(document)),
  );
}

export async function clearTeaDraft(): Promise<void> {
  const imageIds = getTeaImageIds(loadTeaDraft());

  window.localStorage.removeItem(TEA_STORAGE_KEY);
  await deleteEvidenceImages(imageIds);
}

function findSavedDraft(): string | null {
  const current = window.localStorage.getItem(STORAGE_KEY);
  if (current) {
    return current;
  }

  for (const key of LEGACY_STORAGE_KEYS) {
    const legacy = window.localStorage.getItem(key);
    if (legacy) {
      return legacy;
    }
  }

  return null;
}

function normalizeDraft(value: unknown, fallback: OtDocument): OtDocument {
  const candidate = value as DraftCandidate;
  const permissionGroups = normalizePermissionGroups(candidate, fallback);

  return {
    metadata: {
      ...fallback.metadata,
      ...(candidate.metadata ?? {}),
    },
    objective:
      typeof candidate.objective === "string" ? candidate.objective : fallback.objective,
    accessSteps: normalizeAccessSteps(candidate.accessSteps, fallback.accessSteps),
    permissionGroups,
    permissionBlocks: normalizePermissionBlocks(
      candidate.permissionBlocks,
      candidate.tests,
      permissionGroups,
    ),
  };
}

function normalizeAccessSteps(
  steps: OtDocument["accessSteps"] | unknown,
  fallback: AccessStep[],
): AccessStep[] {
  if (!Array.isArray(steps) || steps.length === 0) {
    return fallback;
  }

  return steps.map((step, index) => {
    if (typeof step === "string") {
      return { id: `step-${index + 1}`, text: step };
    }

    const candidate = step as Partial<AccessStep>;
    return {
      id: textOrFallback(candidate.id, `step-${index + 1}`),
      text: textOrFallback(candidate.text, ""),
    };
  });
}

function normalizePermissionGroups(
  candidate: DraftCandidate,
  fallback: OtDocument,
): PermissionGroup[] {
  if (Array.isArray(candidate.permissionGroups) && candidate.permissionGroups.length > 0) {
    return candidate.permissionGroups.map((group, index) => ({
      ...normalizePermissionItem(group, `macro-${index + 1}`),
      microPermissions: normalizePermissionItems(group.microPermissions, "micro"),
    }));
  }

  if (Array.isArray(candidate.macroPermissions) && candidate.macroPermissions.length > 0) {
    const legacyMicros = normalizePermissionItems(candidate.microPermissions, "micro");

    return candidate.macroPermissions.map((macro, index) => ({
      ...normalizePermissionItem(macro, `macro-${index + 1}`),
      microPermissions: legacyMicros.map((micro) => ({ ...micro })),
    }));
  }

  return fallback.permissionGroups;
}

function normalizePermissionItems(
  items: PermissionItem[] | undefined,
  prefix: string,
): PermissionItem[] {
  if (!Array.isArray(items)) {
    return [];
  }

  return items.map((item, index) => normalizePermissionItem(item, `${prefix}-${index + 1}`));
}

function normalizePermissionItem(
  item: Partial<PermissionItem>,
  fallbackId: string,
): PermissionItem {
  return {
    id: textOrFallback(item.id, fallbackId),
    code: textOrFallback(item.code, ""),
    label: textOrFallback(item.label, ""),
    selected: typeof item.selected === "boolean" ? item.selected : true,
  };
}

function normalizePermissionBlocks(
  blocks: Record<string, LegacyPermissionBlock> | undefined,
  legacyTests: LegacyTestDefinition[] | undefined,
  groups: PermissionGroup[],
): Record<string, PermissionBlock> {
  const nextBlocks: Record<string, PermissionBlock> = {};
  const activeKeys = new Set(
    groups.flatMap((macro) =>
      macro.microPermissions.map((micro) => createPermissionKey(macro.id, micro.id)),
    ),
  );

  for (const key of activeKeys) {
    nextBlocks[key] = normalizePermissionBlock(blocks?.[key], legacyTests);
  }

  if (blocks) {
    Object.entries(blocks).forEach(([key, block]) => {
      if (!nextBlocks[key]) {
        nextBlocks[key] = normalizePermissionBlock(block, legacyTests);
      }
    });
  }

  return nextBlocks;
}

function normalizePermissionBlock(
  block: LegacyPermissionBlock | undefined,
  legacyTests: LegacyTestDefinition[] | undefined,
): PermissionBlock {
  if (!block) {
    return {
      tests: normalizeLegacyTests(legacyTests, {}),
    };
  }

  if (Array.isArray(block.tests)) {
    return {
      tests: normalizeBlockTests(block.tests),
    };
  }

  return {
    tests: normalizeLegacyTests(legacyTests, block.results ?? {}),
  };
}

function normalizeBlockTests(tests: PermissionBlockTest[]): PermissionBlockTest[] {
  return tests.map((test, index) => ({
    id: textOrFallback(test.id, `test-${index + 1}`),
    title: textOrFallback(test.title, ""),
    result: normalizeTestResult(
      test.result,
      test.correction,
      textOrFallback(test.id, `test-${index + 1}`),
    ),
    correction: normalizeTestCorrection(test.correction),
  }));
}

function normalizeLegacyTests(
  tests: LegacyTestDefinition[] | undefined,
  results: Record<string, TestResult>,
): PermissionBlockTest[] {
  if (!Array.isArray(tests)) {
    return [];
  }

  return tests.map((test, index) => ({
    id: textOrFallback(test.id, `test-${index + 1}`),
    title: textOrFallback(test.title, ""),
    result: normalizeTestResult(results[test.id]),
    correction: createEmptyTestCorrection(),
  }));
}

function normalizeTestResult(
  result: TestResult | undefined,
  legacyCorrection?: TestCorrection,
  testId = "test",
): TestResult {
  const fallback = createEmptyTestResult();

  if (!result) {
    return fallback;
  }

  const checks = {
    ...fallback.checks,
    ...(result.checks ?? {}),
  };
  const normalizedResult = {
    checks,
    observations: textOrFallback(result.observations, ""),
    legacyImages: normalizeEvidenceImages(result.legacyImages),
    newImages: normalizeEvidenceImages(result.newImages),
    errors: [] as TestError[],
  };

  normalizedResult.errors = normalizeTestErrors(
    {
      ...result,
      ...normalizedResult,
    },
    legacyCorrection,
    testId,
  );

  return normalizedResult;
}

function normalizeTestErrors(
  result: TestResult,
  legacyCorrection: TestCorrection | undefined,
  testId: string,
): TestError[] {
  const candidateErrors = (result as Partial<TestResult>).errors;

  if (Array.isArray(candidateErrors)) {
    return candidateErrors.map((error, index) =>
      normalizeTestError(error, `${testId}-error-${index + 1}`),
    );
  }

  const migratedErrors: TestError[] = [];
  const checks = result.checks ?? createEmptyTestResult().checks;

  if (checks.bothIssue || checks.errorReport) {
    migratedErrors.push({
      id: `${testId}-legacy-error`,
      origin: "legacy",
      observation: textOrFallback(result.observations, ""),
      images: normalizeEvidenceImages(result.legacyImages),
      correction: createEmptyTestCorrection(),
    });
  }

  if (checks.newIssue) {
    migratedErrors.push({
      id: `${testId}-new-error`,
      origin: "new",
      observation: textOrFallback(result.observations, ""),
      images: normalizeEvidenceImages(result.newImages),
      correction: normalizeTestCorrection(legacyCorrection),
    });
  }

  return migratedErrors;
}

function normalizeTestError(error: TestError, fallbackId: string): TestError {
  const origin = normalizeTestErrorOrigin(error.origin);

  return {
    id: textOrFallback(error.id, fallbackId),
    origin,
    observation: textOrFallback(error.observation, ""),
    images: normalizeEvidenceImages(error.images),
    correction: normalizeTestCorrection(error.correction),
  };
}

function normalizeTestErrorOrigin(origin: unknown): TestErrorOrigin {
  return origin === "legacy" ? "legacy" : "new";
}

function normalizeTestCorrection(correction: TestCorrection | undefined): TestCorrection {
  const fallback = createEmptyTestCorrection();

  if (!correction) {
    return fallback;
  }

  const cloudStage = correction.cloudStage;

  return {
    corrected: typeof correction.corrected === "boolean" ? correction.corrected : false,
    beforeImages: normalizeEvidenceImages(correction.beforeImages),
    afterImages: normalizeEvidenceImages(correction.afterImages),
    hotfixTag: textOrFallback(correction.hotfixTag, ""),
    correctedBy: textOrFallback(correction.correctedBy, ""),
    cloudStage:
      cloudStage === "dev" || cloudStage === "homolog" || cloudStage === "production"
        ? cloudStage
        : "none",
  };
}

function normalizeEvidenceImages(images: unknown): EvidenceImage[] {
  if (!Array.isArray(images)) {
    return [];
  }

  return images.map((image, index) => {
    const candidate = image as Partial<EvidenceImage>;

    return {
      id: textOrFallback(candidate.id, `image-${index + 1}`),
      label: textOrFallback(candidate.label, ""),
      name: textOrFallback(candidate.name, "imagem"),
      dataUrl: textOrUndefined(candidate.dataUrl),
      width: numberOrFallback(candidate.width, 560),
      height: numberOrFallback(candidate.height, 320),
      originalBytes: numberOrUndefined(candidate.originalBytes),
      savedBytes: numberOrUndefined(candidate.savedBytes),
      optimized: typeof candidate.optimized === "boolean" ? candidate.optimized : undefined,
    };
  });
}

function normalizeTeaDraft(value: unknown, fallback: TeaDocument): TeaDocument {
  const candidate = value as Partial<TeaDocument>;

  return {
    metadata: {
      ...fallback.metadata,
      ...(candidate.metadata ?? {}),
    },
    overview: textOrFallback(candidate.overview, fallback.overview),
    activityIntro: textOrFallback(candidate.activityIntro, fallback.activityIntro),
    activityImages: normalizeEvidenceImages(candidate.activityImages),
    activities: normalizeTeaActivities(candidate.activities, fallback.activities),
  };
}

function normalizeTeaActivities(
  activities: unknown,
  fallback: TeaActivity[],
): TeaActivity[] {
  if (!Array.isArray(activities) || activities.length === 0) {
    return fallback;
  }

  return activities.map((activity, index) => {
    const candidate = activity as Partial<TeaActivity> & LegacyTeaContentFields;

    return {
      id: textOrFallback(candidate.id, `tea-activity-${index + 1}`),
      title: textOrFallback(candidate.title, ""),
      blocks: normalizeTeaContentBlocks(candidate.blocks, candidate, `tea-activity-${index + 1}`),
      subActivities: normalizeTeaSubActivities(candidate.subActivities, index),
    };
  });
}

function normalizeTeaSubActivities(
  subActivities: unknown,
  activityIndex: number,
): TeaSubActivity[] {
  if (!Array.isArray(subActivities)) {
    return [];
  }

  return subActivities.map((subActivity, index) => {
    const candidate = subActivity as Partial<TeaSubActivity> & LegacyTeaContentFields;

    return {
      id: textOrFallback(candidate.id, `tea-sub-${activityIndex + 1}-${index + 1}`),
      title: textOrFallback(candidate.title, ""),
      blocks: normalizeTeaContentBlocks(
        candidate.blocks,
        candidate,
        `tea-sub-${activityIndex + 1}-${index + 1}`,
      ),
    };
  });
}

type LegacyTeaContentFields = {
  description?: unknown;
  items?: unknown;
  images?: unknown;
};

function normalizeTeaContentBlocks(
  blocks: unknown,
  legacyFields: LegacyTeaContentFields,
  prefix: string,
): TeaContentBlock[] {
  if (Array.isArray(blocks)) {
    return blocks.flatMap((block, index) => {
      const normalized = normalizeTeaContentBlock(block, index, prefix);
      return normalized ? [normalized] : [];
    });
  }

  return normalizeLegacyTeaContentBlocks(legacyFields, prefix);
}

function normalizeTeaContentBlock(
  block: unknown,
  index: number,
  prefix: string,
): TeaContentBlock | null {
  const candidate = block as Partial<TeaContentBlock> & {
    type?: unknown;
    text?: unknown;
    items?: unknown;
    images?: unknown;
  };
  const id = textOrFallback(candidate.id, `${prefix}-block-${index + 1}`);

  if (candidate.type === "text") {
    return {
      id,
      type: "text",
      text: textOrFallback(candidate.text, ""),
    };
  }

  if (candidate.type === "list") {
    return {
      id,
      type: "list",
      items: normalizeTeaTextItems(candidate.items, `${prefix}-item-${index + 1}`),
    };
  }

  if (candidate.type === "images") {
    return {
      id,
      type: "images",
      images: normalizeEvidenceImages(candidate.images),
    };
  }

  return null;
}

function normalizeLegacyTeaContentBlocks(
  legacyFields: LegacyTeaContentFields,
  prefix: string,
): TeaContentBlock[] {
  const blocks: TeaContentBlock[] = [];
  const description = textOrFallback(legacyFields.description, "");
  const items = normalizeTeaTextItems(legacyFields.items, `${prefix}-item`);
  const images = normalizeEvidenceImages(legacyFields.images);

  if (description.trim()) {
    blocks.push({
      id: `${prefix}-text`,
      type: "text",
      text: description,
    });
  }

  if (items.some((item) => item.text.trim())) {
    blocks.push({
      id: `${prefix}-list`,
      type: "list",
      items,
    });
  }

  if (images.length > 0) {
    blocks.push({
      id: `${prefix}-images`,
      type: "images",
      images,
    });
  }

  return blocks;
}

function normalizeTeaTextItems(items: TeaTextItem[] | unknown, prefix: string): TeaTextItem[] {
  if (!Array.isArray(items)) {
    return [];
  }

  return items.map((item, index) => {
    if (typeof item === "string") {
      return { id: `${prefix}-${index + 1}`, text: item };
    }

    const candidate = item as Partial<TeaTextItem>;

    return {
      id: textOrFallback(candidate.id, `${prefix}-${index + 1}`),
      text: textOrFallback(candidate.text, ""),
    };
  });
}

async function deleteEvidenceImages(imageIds: string[]): Promise<void> {
  try {
    await deleteEvidenceImageDataBatch(imageIds);
  } catch {
    // A falta do IndexedDB nao deve impedir a limpeza do rascunho leve.
  }
}

function getOtImageIds(document: OtDocument): string[] {
  return Object.values(document.permissionBlocks).flatMap((block) =>
    block.tests.flatMap((test) => [
      ...test.result.legacyImages.map((image) => image.id),
      ...test.result.newImages.map((image) => image.id),
      ...test.result.errors.flatMap((error) => [
        ...error.images.map((image) => image.id),
        ...error.correction.beforeImages.map((image) => image.id),
        ...error.correction.afterImages.map((image) => image.id),
      ]),
      ...(test.correction?.beforeImages ?? []).map((image) => image.id),
      ...(test.correction?.afterImages ?? []).map((image) => image.id),
    ]),
  );
}

function getTeaImageIds(document: TeaDocument): string[] {
  return [
    ...document.activityImages.map((image) => image.id),
    ...document.activities.flatMap((activity) => [
      ...activity.blocks.flatMap(getTeaBlockImageIds),
      ...activity.subActivities.flatMap((subActivity) =>
        subActivity.blocks.flatMap(getTeaBlockImageIds),
      ),
    ]),
  ];
}

function getTeaBlockImageIds(block: TeaContentBlock): string[] {
  return block.type === "images" ? block.images.map((image) => image.id) : [];
}

function textOrFallback(value: unknown, fallback: string): string {
  return typeof value === "string" ? value : fallback;
}

function textOrUndefined(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value : undefined;
}

function numberOrFallback(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function numberOrUndefined(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}
