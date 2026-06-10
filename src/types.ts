export type AccessStep = {
  id: string;
  text: string;
};

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
  dataUrl?: string;
  width: number;
  height: number;
  originalBytes?: number;
  savedBytes?: number;
  optimized?: boolean;
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

export type CorrectionCloudStage = "none" | "dev" | "homolog" | "production";

export type TestCorrection = {
  corrected: boolean;
  beforeImages: EvidenceImage[];
  afterImages: EvidenceImage[];
  hotfixTag: string;
  correctedBy: string;
  cloudStage: CorrectionCloudStage;
};

export type PermissionBlockTest = {
  id: string;
  title: string;
  result: TestResult;
  correction?: TestCorrection;
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

export type TeaTextItem = {
  id: string;
  text: string;
};

export type TeaTextBlock = {
  id: string;
  type: "text";
  text: string;
};

export type TeaListBlock = {
  id: string;
  type: "list";
  items: TeaTextItem[];
};

export type TeaImagesBlock = {
  id: string;
  type: "images";
  images: EvidenceImage[];
};

export type TeaContentBlock = TeaTextBlock | TeaListBlock | TeaImagesBlock;

export type TeaContentBlockType = TeaContentBlock["type"];

export type TeaSubActivity = {
  id: string;
  title: string;
  blocks: TeaContentBlock[];
};

export type TeaActivity = {
  id: string;
  title: string;
  blocks: TeaContentBlock[];
  subActivities: TeaSubActivity[];
};

export type TeaDocument = {
  metadata: {
    serviceOrder: string;
    phase: string;
    ticket: string;
    subject: string;
    date: string;
    author: string;
  };
  overview: string;
  activityIntro: string;
  activityImages: EvidenceImage[];
  activities: TeaActivity[];
};
