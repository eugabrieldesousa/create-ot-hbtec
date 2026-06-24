import type {
  CheckKey,
  OtDocument,
  TestCorrection,
  TestError,
  TestErrorLegacyReference,
  TestErrorNewStatus,
  TestResult,
} from "./types";

export const checkLabels: Record<CheckKey, string> = {
  sameBehavior: "Funcionou legado e novo estão com o mesmo comportamento",
  possibleIssue: "Possível problema de lógica ou regra de negócio",
  bothIssue: "Erro no legado",
  newIssue: "Erro no novo",
  errorReport: "Relatório de Erros",
};

export const checkLabelAliases: Record<CheckKey, string[]> = {
  sameBehavior: [checkLabels.sameBehavior],
  possibleIssue: [
    checkLabels.possibleIssue,
    "Possível problema/Comportamento estranho/Ambiguidade/incoerência",
  ],
  bothIssue: [checkLabels.bothIssue, "Problema em AMBOS"],
  newIssue: [checkLabels.newIssue, "Problema só no NOVO"],
  errorReport: [checkLabels.errorReport, "Relatorio de Erros"],
};

export const checkOrder: CheckKey[] = [
  "sameBehavior",
  "possibleIssue",
  "bothIssue",
  "newIssue",
  "errorReport",
];

export function createEmptyTestResult(): TestResult {
  return {
    checks: {
      sameBehavior: false,
      possibleIssue: false,
      bothIssue: false,
      newIssue: false,
      errorReport: false,
    },
    observations: "",
    legacyImages: [],
    newImages: [],
    errors: [],
  };
}

export function getEffectiveChecks(
  checks: Record<CheckKey, boolean>,
  errors?: TestError[],
): Record<CheckKey, boolean> {
  if (errors && errors.length > 0) {
    const hasLegacyError = errors.some((error) => error.origin === "legacy");
    const hasNewError = errors.some((error) => error.origin === "new");

    return {
      ...checks,
      sameBehavior: checks.sameBehavior && errors.length === 0,
      bothIssue: hasLegacyError,
      newIssue: hasNewError,
      errorReport: hasLegacyError,
    };
  }

  return {
    ...checks,
    errorReport: checks.bothIssue,
  };
}

export function createEmptyTestCorrection(): TestCorrection {
  return {
    corrected: false,
    beforeImages: [],
    afterImages: [],
    hotfixTag: "",
    correctedBy: "",
    cloudStage: "none",
  };
}

export function createEmptyTestErrorLegacyReference(): TestErrorLegacyReference {
  return {
    enabled: false,
    description: "",
    images: [],
  };
}

export function createEmptyTestErrorNewStatus(): TestErrorNewStatus {
  return {
    works: false,
    images: [],
  };
}

export function createEmptyTestError(id: string, origin: TestError["origin"]): TestError {
  return {
    id,
    origin,
    observation: "",
    images: [],
    legacyReference: createEmptyTestErrorLegacyReference(),
    newStatus: createEmptyTestErrorNewStatus(),
    correction: createEmptyTestCorrection(),
  };
}

export function createPermissionKey(macroId: string, microId: string): string {
  return `${macroId}:${microId}`;
}

export function createDefaultDocument(): OtDocument {
  const today = new Date().toISOString().slice(0, 10);

  return {
    metadata: {
      screen: "",
      responsible: "GABRIEL",
      date: today,
      environment: "LOCAL + DESENVOLVIMENTO",
      author: "GABRIEL",
    },
    objective:
      "Garantir que a tela esteja funcionando corretamente considerando navegação, regras de negócio, estados, e principalmente permissões (micro e macro).",
    accessSteps: [{ id: "step-default", text: "" }],
    permissionGroups: [
      {
        id: "macro-ao",
        code: "AO",
        label: "Administrador Geral",
        selected: true,
        microPermissions: [
          {
            id: "micro-at",
            code: "AT",
            label: "Atualização",
            selected: true,
          },
        ],
      },
    ],
    permissionBlocks: {
      "macro-ao:micro-at": { tests: [] },
    },
  };
}
