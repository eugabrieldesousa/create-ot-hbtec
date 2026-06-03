export type AccessStep = {
  id: string;
  text: string;
};

export type PermissionTestMode = "test" | "idem";

export type CheckKey =
  | "sameBehavior"
  | "possibleIssue"
  | "bothIssue"
  | "newIssue"
  | "errorReport";

export type EvidenceImage = {
  id: string;
  label: string;
  name: string;
  dataUrl: string;
  width: number;
  height: number;
};

export type PermissionItem = {
  id: string;
  code: string;
  label: string;
  selected: boolean;
};

export type PermissionGroup = PermissionItem & {
  microPermissions: PermissionItem[];
};

export type TestResult = {
  checks: Record<CheckKey, boolean>;
  observations: string;
  legacyImages: EvidenceImage[];
  newImages: EvidenceImage[];
};

export type PermissionBlockTest = {
  id: string;
  title: string;
  mode: PermissionTestMode;
  idemReferenceKey?: string;
  result: TestResult;
};

export type PermissionBlock = {
  tests: PermissionBlockTest[];
};

export type OtDocument = {
  metadata: {
    screen: string;
    responsible: string;
    date: string;
    environment: string;
    author: string;
  };
  objective: string;
  accessSteps: AccessStep[];
  permissionGroups: PermissionGroup[];
  permissionBlocks: Record<string, PermissionBlock>;
};
