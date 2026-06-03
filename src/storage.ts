import { createDefaultDocument, createEmptyTestResult, createPermissionKey } from "./defaultDocument";
import type {
  AccessStep,
  OtDocument,
  PermissionBlock,
  PermissionBlockTest,
  PermissionGroup,
  PermissionItem,
  TestResult,
} from "./types";

const STORAGE_KEY = "create-ot-draft-v3";
const LEGACY_STORAGE_KEYS = ["create-ot-draft-v2"];

type LegacyTestDefinition = {
  id: string;
  title: string;
};

type LegacyPermissionBlock = {
  status?: "complete" | "idem" | "no-access";
  mode?: "test" | "idem";
  idemReferenceKey?: string;
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
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(document));
}

export function clearDraft(): void {
  window.localStorage.removeItem(STORAGE_KEY);
  LEGACY_STORAGE_KEYS.forEach((key) => window.localStorage.removeItem(key));
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

  return clearInvalidIdemReferences(nextBlocks);
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
      tests: normalizeBlockTests(block.tests, block),
    };
  }

  return {
    tests: normalizeLegacyTests(legacyTests, block.results ?? {}, block),
  };
}

function normalizeBlockTests(
  tests: PermissionBlockTest[],
  block: LegacyPermissionBlock,
): PermissionBlockTest[] {
  const legacyMode = block.status === "idem" || block.mode === "idem" ? "idem" : "test";

  return tests.map((test, index) => ({
    id: textOrFallback(test.id, `test-${index + 1}`),
    title: textOrFallback(test.title, ""),
    mode: test.mode ?? legacyMode,
    idemReferenceKey: textOrUndefined(test.idemReferenceKey),
    result: normalizeTestResult(test.result),
  }));
}

function normalizeLegacyTests(
  tests: LegacyTestDefinition[] | undefined,
  results: Record<string, TestResult>,
  block?: LegacyPermissionBlock,
): PermissionBlockTest[] {
  if (!Array.isArray(tests)) {
    return [];
  }

  const mode = block?.status === "idem" || block?.mode === "idem" ? "idem" : "test";

  return tests.map((test, index) => ({
    id: textOrFallback(test.id, `test-${index + 1}`),
    title: textOrFallback(test.title, ""),
    mode,
    result: normalizeTestResult(results[test.id]),
  }));
}

function normalizeTestResult(result: TestResult | undefined): TestResult {
  const fallback = createEmptyTestResult();

  if (!result) {
    return fallback;
  }

  return {
    checks: {
      ...fallback.checks,
      ...(result.checks ?? {}),
    },
    observations: textOrFallback(result.observations, ""),
    legacyImages: Array.isArray(result.legacyImages) ? result.legacyImages : [],
    newImages: Array.isArray(result.newImages) ? result.newImages : [],
  };
}

function clearInvalidIdemReferences(
  blocks: Record<string, PermissionBlock>,
): Record<string, PermissionBlock> {
  const validKeys = new Set(
    Object.entries(blocks).flatMap(([blockKey, block]) =>
      block.tests.map((test) => createTestReferenceKey(blockKey, test.id)),
    ),
  );

  return Object.fromEntries(
    Object.entries(blocks).map(([key, block]) => [
      key,
      {
        ...block,
        tests: block.tests.map((test) => {
          if (
            test.idemReferenceKey &&
            test.idemReferenceKey !== createTestReferenceKey(key, test.id) &&
            validKeys.has(test.idemReferenceKey)
          ) {
            return test;
          }

          return { ...test, idemReferenceKey: undefined };
        }),
      },
    ]),
  );
}

function textOrFallback(value: unknown, fallback: string): string {
  return typeof value === "string" ? value : fallback;
}

function textOrUndefined(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value : undefined;
}

function createTestReferenceKey(blockKey: string, testId: string): string {
  return `${blockKey}::${testId}`;
}
