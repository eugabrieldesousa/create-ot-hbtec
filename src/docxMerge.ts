import { createPermissionKey } from "./defaultDocument";
import type {
  EvidenceImage,
  OtDocument,
  PermissionBlockTest,
  PermissionGroup,
  PermissionItem,
  TeaActivity,
  TeaContentBlock,
  TeaDocument,
  TeaSubActivity,
  TestCorrection,
  TestError,
  TestResult,
} from "./types";

export type MergeInsertPosition =
  | { mode: "end" }
  | { mode: "before" | "after"; targetId: string };

export type TeaDocxMergeSelection = {
  activityIds: string[];
  subActivityIds: string[];
  activityPosition: MergeInsertPosition;
  subActivityTargetActivityId: string | null;
  subActivityPosition: MergeInsertPosition;
};

export type OtDocxMergeTarget =
  | { kind: "existing"; macroId: string; microId: string }
  | { kind: "new" };

export type OtDocxMergeGroupSelection = {
  sourceMacroId: string;
  sourceMicroId: string;
  testIds: string[];
  target: OtDocxMergeTarget;
  position: MergeInsertPosition;
};

export type OtDocxMergeSelection = {
  groups: OtDocxMergeGroupSelection[];
};

export type TeaDocxMergeResult = {
  document: TeaDocument;
  insertedActivityIds: string[];
  insertedSubActivityIds: string[];
};

export type OtDocxMergeResult = {
  document: OtDocument;
  insertedBlockKeys: string[];
  insertedTestReferences: Array<{ blockKey: string; testId: string }>;
};

type IdFactory = () => string;

export function applyTeaDocxMerge(
  current: TeaDocument,
  source: TeaDocument,
  selection: TeaDocxMergeSelection,
  createId: IdFactory = defaultCreateId,
): TeaDocxMergeResult {
  const selectedActivityIds = new Set(selection.activityIds);
  const selectedSubActivityIds = new Set(selection.subActivityIds);
  const selectedActivities = source.activities.filter((activity) =>
    selectedActivityIds.has(activity.id),
  );
  const looseSubActivities = source.activities.flatMap((activity) =>
    selectedActivityIds.has(activity.id)
      ? []
      : activity.subActivities.filter((subActivity) => selectedSubActivityIds.has(subActivity.id)),
  );

  const clonedActivities = selectedActivities.map((activity) =>
    cloneTeaActivity(activity, selectedSubActivityIds, createId),
  );
  const clonedLooseSubActivities = looseSubActivities.map((subActivity) =>
    cloneTeaSubActivity(subActivity, createId),
  );
  let nextActivities = insertItems(current.activities, clonedActivities, selection.activityPosition);

  if (selection.subActivityTargetActivityId && clonedLooseSubActivities.length > 0) {
    nextActivities = nextActivities.map((activity) =>
      activity.id === selection.subActivityTargetActivityId
        ? {
            ...activity,
            subActivities: insertItems(
              activity.subActivities,
              clonedLooseSubActivities,
              selection.subActivityPosition,
            ),
          }
        : activity,
    );
  }

  return {
    document: {
      ...current,
      activities: nextActivities,
    },
    insertedActivityIds: clonedActivities.map((activity) => activity.id),
    insertedSubActivityIds: [
      ...clonedActivities.flatMap((activity) =>
        activity.subActivities.map((subActivity) => subActivity.id),
      ),
      ...clonedLooseSubActivities.map((subActivity) => subActivity.id),
    ],
  };
}

export function applyOtDocxMerge(
  current: OtDocument,
  source: OtDocument,
  selection: OtDocxMergeSelection,
  createId: IdFactory = defaultCreateId,
): OtDocxMergeResult {
  let nextDocument = cloneOtDocumentShell(current);
  const insertedBlockKeys: string[] = [];
  const insertedTestReferences: Array<{ blockKey: string; testId: string }> = [];

  selection.groups.forEach((groupSelection) => {
    const sourceMacro = source.permissionGroups.find(
      (macro) => macro.id === groupSelection.sourceMacroId,
    );
    const sourceMicro = sourceMacro?.microPermissions.find(
      (micro) => micro.id === groupSelection.sourceMicroId,
    );

    if (!sourceMacro || !sourceMicro || groupSelection.testIds.length === 0) {
      return;
    }

    const sourceBlock =
      source.permissionBlocks[createPermissionKey(sourceMacro.id, sourceMicro.id)];
    const selectedTestIds = new Set(groupSelection.testIds);
    const selectedTests =
      sourceBlock?.tests.filter((test) => selectedTestIds.has(test.id)) ?? [];

    if (selectedTests.length === 0) {
      return;
    }

    const target = ensureOtMergeTarget(
      nextDocument,
      sourceMacro,
      sourceMicro,
      groupSelection.target,
      createId,
    );
    const targetBlock = nextDocument.permissionBlocks[target.blockKey] ?? { tests: [] };
    const clonedTests = selectedTests.map((test) => clonePermissionBlockTest(test, createId));

    nextDocument = {
      ...nextDocument,
      permissionBlocks: {
        ...nextDocument.permissionBlocks,
        [target.blockKey]: {
          tests: insertItems(targetBlock.tests, clonedTests, groupSelection.position),
        },
      },
    };

    if (!insertedBlockKeys.includes(target.blockKey)) {
      insertedBlockKeys.push(target.blockKey);
    }

    clonedTests.forEach((test) => {
      insertedTestReferences.push({ blockKey: target.blockKey, testId: test.id });
    });
  });

  return {
    document: nextDocument,
    insertedBlockKeys,
    insertedTestReferences,
  };
}

export function findMatchingOtMergeTarget(
  current: OtDocument,
  sourceMacro: PermissionItem,
  sourceMicro: PermissionItem,
): OtDocxMergeTarget {
  const matchingMacro = current.permissionGroups.find(
    (macro) => normalizeCode(macro.code) === normalizeCode(sourceMacro.code),
  );
  const matchingMicro = matchingMacro?.microPermissions.find(
    (micro) => normalizeCode(micro.code) === normalizeCode(sourceMicro.code),
  );

  if (matchingMacro && matchingMicro) {
    return {
      kind: "existing",
      macroId: matchingMacro.id,
      microId: matchingMicro.id,
    };
  }

  return { kind: "new" };
}

function ensureOtMergeTarget(
  documentData: OtDocument,
  sourceMacro: PermissionGroup,
  sourceMicro: PermissionItem,
  target: OtDocxMergeTarget,
  createId: IdFactory,
): { blockKey: string } {
  if (target.kind === "existing") {
    const macro = documentData.permissionGroups.find(
      (candidate) => candidate.id === target.macroId,
    );
    const micro = macro?.microPermissions.find((candidate) => candidate.id === target.microId);

    if (macro && micro) {
      macro.selected = true;
      micro.selected = true;
      const blockKey = createPermissionKey(macro.id, micro.id);
      documentData.permissionBlocks[blockKey] = documentData.permissionBlocks[blockKey] ?? {
        tests: [],
      };
      return { blockKey };
    }
  }

  const existingMacro = documentData.permissionGroups.find(
    (candidate) => normalizeCode(candidate.code) === normalizeCode(sourceMacro.code),
  );

  if (existingMacro) {
    const microId = createId();
    const micro = clonePermissionItem(sourceMicro, microId);
    existingMacro.selected = true;
    existingMacro.microPermissions = [...existingMacro.microPermissions, micro];
    const blockKey = createPermissionKey(existingMacro.id, micro.id);
    documentData.permissionBlocks[blockKey] = { tests: [] };
    return { blockKey };
  }

  const macroId = createId();
  const microId = createId();
  const macro: PermissionGroup = {
    ...clonePermissionItem(sourceMacro, macroId),
    microPermissions: [clonePermissionItem(sourceMicro, microId)],
  };
  documentData.permissionGroups = [...documentData.permissionGroups, macro];
  const blockKey = createPermissionKey(macro.id, microId);
  documentData.permissionBlocks[blockKey] = { tests: [] };
  return { blockKey };
}

function cloneOtDocumentShell(documentData: OtDocument): OtDocument {
  return {
    ...documentData,
    metadata: { ...documentData.metadata },
    accessSteps: documentData.accessSteps.map((step) => ({ ...step })),
    permissionGroups: documentData.permissionGroups.map((macro) => ({
      ...macro,
      microPermissions: macro.microPermissions.map((micro) => ({ ...micro })),
    })),
    permissionBlocks: Object.fromEntries(
      Object.entries(documentData.permissionBlocks).map(([key, block]) => [
        key,
        { tests: block.tests },
      ]),
    ),
  };
}

function clonePermissionItem(item: PermissionItem, id: string): PermissionItem {
  return {
    id,
    code: item.code,
    label: item.label,
    selected: true,
  };
}

function clonePermissionBlockTest(
  test: PermissionBlockTest,
  createId: IdFactory,
): PermissionBlockTest {
  return {
    ...test,
    id: createId(),
    result: cloneTestResult(test.result, createId),
    correction: test.correction ? cloneTestCorrection(test.correction, createId) : undefined,
  };
}

function cloneTestResult(result: TestResult, createId: IdFactory): TestResult {
  return {
    checks: { ...result.checks },
    observations: result.observations,
    legacyImages: result.legacyImages.map((image) => cloneEvidenceImage(image, createId)),
    newImages: result.newImages.map((image) => cloneEvidenceImage(image, createId)),
    errors: result.errors.map((error) => cloneTestError(error, createId)),
  };
}

function cloneTestError(error: TestError, createId: IdFactory): TestError {
  return {
    ...error,
    id: createId(),
    images: error.images.map((image) => cloneEvidenceImage(image, createId)),
    correction: cloneTestCorrection(error.correction, createId),
  };
}

function cloneTestCorrection(correction: TestCorrection, createId: IdFactory): TestCorrection {
  return {
    ...correction,
    beforeImages: correction.beforeImages.map((image) => cloneEvidenceImage(image, createId)),
    afterImages: correction.afterImages.map((image) => cloneEvidenceImage(image, createId)),
  };
}

function cloneTeaActivity(
  activity: TeaActivity,
  selectedSubActivityIds: Set<string>,
  createId: IdFactory,
): TeaActivity {
  const selectedSubActivities = activity.subActivities.filter((subActivity) =>
    selectedSubActivityIds.has(subActivity.id),
  );
  const subActivities =
    selectedSubActivities.length > 0 ? selectedSubActivities : activity.subActivities;

  return {
    ...activity,
    id: createId(),
    blocks: activity.blocks.map((block) => cloneTeaContentBlock(block, createId)),
    subActivities: subActivities.map((subActivity) => cloneTeaSubActivity(subActivity, createId)),
  };
}

function cloneTeaSubActivity(
  subActivity: TeaSubActivity,
  createId: IdFactory,
): TeaSubActivity {
  return {
    ...subActivity,
    id: createId(),
    blocks: subActivity.blocks.map((block) => cloneTeaContentBlock(block, createId)),
  };
}

function cloneTeaContentBlock(block: TeaContentBlock, createId: IdFactory): TeaContentBlock {
  if (block.type === "text") {
    return {
      ...block,
      id: createId(),
    };
  }

  if (block.type === "list") {
    return {
      ...block,
      id: createId(),
      items: block.items.map((item) => ({
        ...item,
        id: createId(),
      })),
    };
  }

  return {
    ...block,
    id: createId(),
    images: block.images.map((image) => cloneEvidenceImage(image, createId)),
  };
}

function cloneEvidenceImage(image: EvidenceImage, createId: IdFactory): EvidenceImage {
  return {
    ...image,
    id: createId(),
  };
}

function insertItems<T extends { id: string }>(
  items: T[],
  insertedItems: T[],
  position: MergeInsertPosition,
): T[] {
  if (insertedItems.length === 0) {
    return items;
  }

  if (position.mode === "end") {
    return [...items, ...insertedItems];
  }

  const targetIndex = items.findIndex((item) => item.id === position.targetId);

  if (targetIndex < 0) {
    return [...items, ...insertedItems];
  }

  const insertIndex = position.mode === "before" ? targetIndex : targetIndex + 1;

  return [
    ...items.slice(0, insertIndex),
    ...insertedItems,
    ...items.slice(insertIndex),
  ];
}

function normalizeCode(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function defaultCreateId(): string {
  return crypto.randomUUID();
}
