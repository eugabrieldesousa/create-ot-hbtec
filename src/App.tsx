import {
  createContext,
  lazy,
  memo,
  Suspense,
  useCallback,
  useContext,
  useDeferredValue,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from "react";
import type { ChangeEvent, ClipboardEvent, ComponentProps, DragEvent, ReactNode } from "react";
import { flushSync } from "react-dom";
import {
  ActionIcon,
  Badge,
  Button,
  Card,
  Checkbox,
  Collapse,
  Container,
  Divider,
  FileButton,
  Group,
  Menu,
  Modal,
  Paper,
  SegmentedControl,
  Select,
  Stack,
  Tabs,
  Text,
  Textarea,
  TextInput,
  Title,
  Tooltip,
  useMantineColorScheme,
} from "@mantine/core";
import {
  ArrowDown,
  ArrowUp,
  ChevronDown,
  ChevronsDown,
  ChevronsUp,
  ClipboardPaste,
  AlertCircle,
  CheckCircle2,
  CircleHelp,
  ClipboardList,
  Copy,
  Download,
  FileSearch,
  FileText,
  FileUp,
  ImagePlus,
  ListChecks,
  Moon,
  MoreVertical,
  Plus,
  RotateCcw,
  Save,
  Sun,
  Trash2,
  Wrench,
} from "lucide-react";
import {
  checkLabels,
  checkOrder,
  createEmptyTestCorrection,
  createEmptyTestResult,
  createPermissionKey,
} from "./defaultDocument";
import { buildOtPreviewModel, buildTeaPreviewModel } from "./docxPreviewModel";
import type { DocxPreviewModel } from "./docxPreviewModel";
import type { DocxImportResult } from "./docxImport";
import { LoadingFeedback } from "./LoadingFeedback";
import {
  deleteEvidenceImageData,
  deleteEvidenceImageDataBatch,
  hydrateDocumentImages,
  hydrateTeaDocumentImages,
  persistEmbeddedEvidenceImages,
  persistEmbeddedTeaImages,
  saveEvidenceImageDataBatch,
} from "./imageStorage";
import { optimizeImageFile } from "./imageOptimizer";
import {
  clearDraft,
  clearTeaDraft,
  loadDraft,
  loadTeaDraft,
  saveDraft,
  saveTeaDraft,
} from "./storage";
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
  TeaContentBlockType,
  TeaDocument,
  TeaSubActivity,
  TestCorrection,
  TestResult,
} from "./types";

type PermissionBlockEntry = {
  key: string;
  macro: PermissionGroup;
  micro: PermissionItem;
};

type FilteredPermissionBlockEntry = PermissionBlockEntry & {
  block: PermissionBlock;
  sourceBlock: PermissionBlock;
};

type PermissionGroupEditorProps = {
  index: number;
  macro: PermissionGroup;
  onMacroChange: (macroId: string, updates: Partial<PermissionItem>) => void;
  onRemoveMacro: (macroId: string) => void;
  onAddMicro: (macroId: string) => void;
  onMicroChange: (
    macroId: string,
    microId: string,
    updates: Partial<PermissionItem>,
  ) => void;
  onRemoveMicro: (macroId: string, microId: string) => void;
};

type ActiveTab = "document" | "permissions" | "tests" | "corrections" | "review" | "preview";
type TeaTab = "document" | "activities" | "review" | "preview";
type DocumentKind = "ot" | "tea";
type MoveDirection = "up" | "down";
type TeaOutlineContext = "activities" | "preview";

type DraftStatus =
  | "Alterações pendentes"
  | "Salvando..."
  | "Rascunho salvo"
  | "Rascunho grande demais";

type LoadingTask = {
  label: string;
  detail?: string;
};

type ReviewSummary = {
  selectedPermissions: number;
  testCount: number;
  imageCount: number;
  issues: ReviewIssue[];
};

type ReviewSeverity = "warning" | "danger";

type ReviewIssue = {
  id: string;
  label: string;
  detail: string;
  tab: ActiveTab;
  severity: ReviewSeverity;
  targetId?: string;
  blockKey?: string;
  testId?: string;
};

type TeaReviewSummary = {
  activityCount: number;
  imageCount: number;
  issues: TeaReviewIssue[];
};

type TeaReviewSeverity = ReviewSeverity;
type TeaReviewCategory = "document" | "activity" | "image";

type TeaReviewIssue = {
  id: string;
  label: string;
  detail: string;
  tab: TeaTab;
  severity: TeaReviewSeverity;
  category: TeaReviewCategory;
  targetId?: string;
  activityId?: string;
  subActivityId?: string;
  blockId?: string;
};

type ReviewIssueIndex = {
  byBlockKey: Map<string, ReviewIssue[]>;
  byTestReferenceKey: Map<string, ReviewIssue[]>;
};

type TeaReviewIssueIndex = {
  byTargetId: Map<string, TeaReviewIssue[]>;
  byActivityId: Map<string, TeaReviewIssue[]>;
  byActivityRootId: Map<string, TeaReviewIssue[]>;
  bySubActivityId: Map<string, TeaReviewIssue[]>;
  byBlockId: Map<string, TeaReviewIssue[]>;
  activityIssues: TeaReviewIssue[];
  hasActivityIssues: boolean;
};

type TeaInlineReviewSummary = {
  total: number;
  danger: number;
  warning: number;
};

type InlineReviewSummary = TeaInlineReviewSummary;

const emptyInlineReviewSummary: InlineReviewSummary = { total: 0, danger: 0, warning: 0 };

type ConfirmationTone = "danger";

type ConfirmationOptions = {
  title: string;
  description: string;
  confirmLabel: string;
  tone?: ConfirmationTone;
};

type PendingConfirmation = ConfirmationOptions & {
  onConfirm: () => void | Promise<void>;
};

type TeaSubActivityCopyRequest = {
  sourceActivityId: string;
  subActivityId: string;
  targetActivityId: string | null;
};

type ConfirmAction = (
  options: ConfirmationOptions,
  onConfirm: () => void | Promise<void>,
) => void;

const ConfirmationContext = createContext<ConfirmAction | null>(null);
const BufferedCommitContext = createContext<((commit: () => void) => () => void) | null>(null);

type OutlineItemStatus = "pending" | "ok";

type DocumentOutlineItem = {
  id: string;
  title: string;
  meta?: string;
  tab: string;
  targetId?: string;
  status?: OutlineItemStatus;
  level?: 0 | 1 | 2;
  blockKey?: string;
  testId?: string;
  activityId?: string;
  subActivityId?: string;
  blockId?: string;
};

type DocumentOutlineGroup = {
  id: string;
  title: string;
  items: DocumentOutlineItem[];
};

type CorrectionOccurrence = PermissionBlockEntry & {
  test: PermissionBlockTest;
  testIndex: number;
  referenceKey: string;
};

type CorrectionGroup = {
  key: string;
  title: string;
  occurrences: CorrectionOccurrence[];
  correction: TestCorrection;
};

type FaqSection = {
  title: string;
  items: Array<{
    title: string;
    description: string;
    example?: string;
  }>;
};

type TestBlockFilter = "all" | "withoutTests" | "withoutImages" | "withPending" | "withProblem";

const standardTestTitles = ["Criação", "Edição", "Consulta", "Exclusão"];

type IdleDeadlineLike = {
  didTimeout: boolean;
  timeRemaining: () => number;
};

type IdleWindow = Window & {
  requestIdleCallback?: (
    callback: (deadline: IdleDeadlineLike) => void,
    options?: { timeout?: number },
  ) => number;
  cancelIdleCallback?: (handle: number) => void;
};

type BufferedTextState = {
  value: string;
  setValue: (value: string) => void;
  commit: () => void;
};

type PermissionBlockGroupProps = {
  macro: PermissionGroup;
  entries: FilteredPermissionBlockEntry[];
  expandedTests: Record<string, boolean>;
  reviewIssueIndex: ReviewIssueIndex;
  isCollapsed: boolean;
  collapsedBlocks: Record<string, boolean>;
  onMacroCollapseChange: (macroId: string, collapsed: boolean) => void;
  onBlockCollapseChange: (blockKey: string, collapsed: boolean) => void;
  onAddTest: (blockKey: string) => void;
  onAddStandardTests: (blockKey: string) => void;
  onDuplicateBlockStructure: (blockKey: string) => void;
  onDuplicateTest: (blockKey: string, testId: string) => void;
  onTestExpansionChange: (referenceKey: string, expanded: boolean) => void;
  onTestTitleChange: (blockKey: string, testId: string, title: string) => void;
  onTestRemove: (blockKey: string, testId: string) => void;
  onTestMove: (blockKey: string, index: number, direction: -1 | 1) => void;
  onResultChange: (
    blockKey: string,
    testId: string,
    updater: (result: TestResult) => TestResult,
  ) => void;
};

type PermissionBlockEditorProps = Omit<
  PermissionBlockGroupProps,
  "macro" | "entries" | "isCollapsed" | "collapsedBlocks" | "onMacroCollapseChange"
> & {
  blockKey: string;
  entry: PermissionBlockEntry;
  block: PermissionBlock;
  sourceBlock: PermissionBlock;
  isCollapsed: boolean;
};

type BlockTestEditorProps = {
  blockKey: string;
  index: number;
  test: PermissionBlockTest;
  selfReferenceKey: string;
  isExpanded: boolean;
  reviewIssueIndex: ReviewIssueIndex;
  canMoveUp: boolean;
  canMoveDown: boolean;
  onTestExpansionChange: (referenceKey: string, expanded: boolean) => void;
  onTestTitleChange: (blockKey: string, testId: string, title: string) => void;
  onTestMove: (blockKey: string, index: number, direction: -1 | 1) => void;
  onDuplicateTest: (blockKey: string, testId: string) => void;
  onTestRemove: (blockKey: string, testId: string) => void;
  onResultChange: (
    blockKey: string,
    testId: string,
    updater: (result: TestResult) => TestResult,
  ) => void;
};

type TeaWorkspaceProps = {
  documentData: TeaDocument;
  previewModel: DocxPreviewModel | null;
  isPreviewStale: boolean;
  activeTab: TeaTab;
  reviewSummary: TeaReviewSummary;
  reviewIssueIndex: TeaReviewIssueIndex;
  collapsedActivities: Record<string, boolean>;
  collapsedSubActivities: Record<string, boolean>;
  collapsedComposers: Record<string, boolean>;
  collapsedContentBlocks: Record<string, boolean>;
  onTabChange: (tab: TeaTab) => void;
  onMetadataChange: (field: keyof TeaDocument["metadata"], value: string) => void;
  onOverviewChange: (value: string) => void;
  onActivityIntroChange: (value: string) => void;
  onActivityImagesChange: (updater: (images: EvidenceImage[]) => EvidenceImage[]) => void;
  onAddActivity: () => void;
  onActivityChange: (
    activityId: string,
    updater: (activity: TeaActivity) => TeaActivity,
  ) => void;
  onActivityRemove: (activityId: string) => void;
  onActivityMove: (activityId: string, direction: MoveDirection) => void;
  onAddSubActivity: (activityId: string) => void;
  onSubActivityChange: (
    activityId: string,
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ) => void;
  onSubActivityRemove: (activityId: string, subActivityId: string) => void;
  onSubActivityMove: (
    activityId: string,
    subActivityId: string,
    direction: MoveDirection,
  ) => void;
  onSubActivityCopy: (activityId: string, subActivityId: string) => void;
  onActivityCollapseChange: (activityId: string, collapsed: boolean) => void;
  onSubActivityCollapseChange: (subActivityId: string, collapsed: boolean) => void;
  onComposerCollapseChange: (composerId: string, collapsed: boolean) => void;
  onContentBlockCollapseChange: (blockId: string, collapsed: boolean) => void;
  onReviewIssueClick: (issue: TeaReviewIssue) => void;
  onRefreshPreview: () => void;
};

type TeaActivityEditorProps = {
  index: number;
  totalActivities: number;
  activity: TeaActivity;
  reviewIssues: TeaReviewIssue[];
  isCollapsed: boolean;
  collapsedSubActivities: Record<string, boolean>;
  collapsedComposers: Record<string, boolean>;
  collapsedContentBlocks: Record<string, boolean>;
  onChange: (updater: (activity: TeaActivity) => TeaActivity) => void;
  onRemove: () => void;
  onMove: (direction: MoveDirection) => void;
  onAddSubActivity: () => void;
  onSubActivityChange: (
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ) => void;
  onSubActivityRemove: (subActivityId: string) => void;
  onSubActivityMove: (subActivityId: string, direction: MoveDirection) => void;
  onSubActivityCopy: (subActivityId: string) => void;
  onCollapseChange: (collapsed: boolean) => void;
  onSubActivityCollapseChange: (subActivityId: string, collapsed: boolean) => void;
  onComposerCollapseChange: (composerId: string, collapsed: boolean) => void;
  onContentBlockCollapseChange: (blockId: string, collapsed: boolean) => void;
};

type TeaSubActivityEditorProps = {
  activityIndex: number;
  index: number;
  totalSubActivities: number;
  subActivity: TeaSubActivity;
  reviewIssues: TeaReviewIssue[];
  isCollapsed: boolean;
  isComposerCollapsed: boolean;
  collapsedBlocks: Record<string, boolean>;
  onChange: (updater: (subActivity: TeaSubActivity) => TeaSubActivity) => void;
  onRemove: () => void;
  onMove: (direction: MoveDirection) => void;
  onCopy: () => void;
  onCollapseChange: (collapsed: boolean) => void;
  onComposerCollapseChange: (collapsed: boolean) => void;
  onContentBlockCollapseChange: (blockId: string, collapsed: boolean) => void;
};

type TeaContentComposerProps = {
  composerId: string;
  title: string;
  description: string;
  emptyMessage: string;
  emptyActionLabel: string;
  blocks: TeaContentBlock[];
  tone: "legacy" | "new";
  isCollapsed: boolean;
  collapsedBlocks: Record<string, boolean>;
  reviewIssues: TeaReviewIssue[];
  onCollapseChange: (collapsed: boolean) => void;
  onBlockCollapseChange: (blockId: string, collapsed: boolean) => void;
  onBlocksChange: (updater: (blocks: TeaContentBlock[]) => TeaContentBlock[]) => void;
};

type TeaContentBlockEditorProps = {
  block: TeaContentBlock;
  index: number;
  totalBlocks: number;
  tone: "legacy" | "new";
  isCollapsed: boolean;
  reviewIssues: TeaReviewIssue[];
  onChange: (updater: (block: TeaContentBlock) => TeaContentBlock) => void;
  onMove: (direction: MoveDirection) => void;
  onDuplicate: () => Promise<void>;
  onRemove: () => void;
  onCollapseChange: (collapsed: boolean) => void;
};

const problemCheckKeys: CheckKey[] = ["possibleIssue", "bothIssue", "newIssue", "errorReport"];
const emptyReviewIssues: ReviewIssue[] = [];
const emptyTeaReviewIssues: TeaReviewIssue[] = [];

const DocxPreview = lazy(async () => {
  const { DocxPreview } = await import("./docxPreview");

  return { default: DocxPreview };
});

function useConfirmAction(): ConfirmAction {
  const confirmAction = useContext(ConfirmationContext);

  if (!confirmAction) {
    throw new Error("useConfirmAction must be used inside ConfirmationContext.Provider");
  }

  return confirmAction;
}

const quickCheckLabels: Record<CheckKey, string> = {
  sameBehavior: "OK",
  possibleIssue: "Possível",
  bothIssue: "Ambos",
  newIssue: "Novo",
  errorReport: "Erros",
};

const quickCheckToneClassNames: Record<CheckKey, string> = {
  sameBehavior: "",
  possibleIssue: "",
  bothIssue: "quickCheck--warning",
  newIssue: "quickCheck--warning",
  errorReport: "quickCheck--danger",
};

const teaContentBlockLabels: Record<TeaContentBlockType, string> = {
  text: "Texto",
  list: "Lista",
  images: "Imagens",
};

const testBlockFilterLabels: Record<TestBlockFilter, string> = {
  all: "Todos",
  withoutTests: "Sem testes",
  withoutImages: "Sem imagens",
  withPending: "Com pendência",
  withProblem: "Com problema",
};

const testBlockFilterOrder: TestBlockFilter[] = [
  "all",
  "withoutTests",
  "withoutImages",
  "withPending",
  "withProblem",
];

const cloudStageOptions: Array<{ value: TestCorrection["cloudStage"]; label: string }> = [
  { value: "none", label: "Nao enviado" },
  { value: "dev", label: "Ate dev" },
  { value: "homolog", label: "Ate homolog" },
  { value: "production", label: "Ate producao" },
];

const emptyPermissionBlock: PermissionBlock = { tests: [] };

const faqSections: FaqSection[] = [
  {
    title: "Fluxo recomendado",
    items: [
      {
        title: "Comece pelo Documento",
        description:
          "Preencha tela, responsável, data, ambiente, elaborada por e objetivo. Esses campos montam o cabeçalho e ajudam a Revisão a avisar o que falta.",
      },
      {
        title: "Monte as permissões antes dos testes",
        description:
          "A aba Testes nasce das macros e micros marcadas como Usar. Se uma permissão não estiver selecionada, ela não aparece nos blocos nem no DOCX.",
      },
      {
        title: "Feche pela Revisão",
        description:
          "A Revisão mostra pendências, contadores de testes e imagens para conferir o documento antes da exportação.",
      },
    ],
  },
  {
    title: "Documento e passos",
    items: [
      {
        title: "Editar passos em lote",
        description:
          "Cada linha do campo vira um passo do documento. Enter cria uma nova linha; passos vazios podem ficar enquanto você edita e são ignorados na exportação.",
        example:
          "Acessar o menu Cadastros\nAbrir a tela de usuários\nSelecionar um registro",
      },
      {
        title: "Campos individuais de passo",
        description:
          "Servem para ajuste fino, remoção e revisão visual. Ao colar várias linhas em uma etapa, o sistema divide o texto em passos separados.",
      },
      {
        title: "Limpar",
        description:
          "Apaga o rascunho salvo no navegador e volta para o documento padrão. Use com cuidado quando quiser começar uma OT do zero.",
      },
    ],
  },
  {
    title: "Permissões",
    items: [
      {
        title: "Lista rápida",
        description:
          "Permite criar macros e micros em texto. Macro fica sem recuo; micro fica com espaços no começo ou com hífen.",
        example:
          "AO - Administrador Geral\n  AT - Atualização\n  SC - Somente Consulta\nUS - Usuário\n  CO - Consulta",
      },
      {
        title: "Carregar atual",
        description:
          "Copia as permissões já cadastradas para o campo de Lista rápida. É útil para reorganizar ou corrigir tudo em bloco.",
      },
      {
        title: "Aplicar lista",
        description:
          "Substitui a lista de permissões pelo texto informado. Quando os códigos de macro e micro batem com os antigos, os testes existentes são preservados.",
      },
    ],
  },
  {
    title: "Testes e evidências",
    items: [
      {
        title: "Adicionar pacote padrão",
        description:
          "Faz a mesma lista padrão, mas apenas no bloco de permissão onde o botão foi clicado.",
      },
      {
        title: "Copiar para vazios",
        description:
          "Copia os testes do bloco atual para outros blocos que ainda não têm testes. Copia títulos, checks e observações, mas não copia imagens.",
      },
      {
        title: "Checks e imagens",
        description:
          "Os botões OK, Possível, Ambos, Novo e Erros marcam rapidamente o resultado. Em Legado e Novo, você pode colar, arrastar ou selecionar imagens.",
      },
    ],
  },
  {
    title: "Saída e rascunho",
    items: [
      {
        title: "Importar DOCX",
        description:
          "Lê um documento existente, mostra uma prévia e só substitui o rascunho quando você confirma.",
      },
      {
        title: "Exportar DOCX",
        description:
          "Salva o rascunho atual e gera o arquivo final com cabeçalho, objetivo, passos, permissões, testes, observações e imagens.",
      },
      {
        title: "Salvamento automático",
        description:
          "As alterações ficam salvas no navegador. O selo do topo mostra se há alterações pendentes, salvando, salvo ou rascunho grande demais.",
      },
    ],
  },
];

export default function App() {
  const { colorScheme, toggleColorScheme } = useMantineColorScheme();
  const [documentKind, setDocumentKind] = useState<DocumentKind>("ot");
  const [documentData, setDocumentData] = useState<OtDocument>(() => loadDraft());
  const [teaData, setTeaData] = useState<TeaDocument>(() => loadTeaDraft());
  const [otPreviewDocumentData, setOtPreviewDocumentData] =
    useState<OtDocument>(() => documentData);
  const [teaPreviewDocumentData, setTeaPreviewDocumentData] =
    useState<TeaDocument>(() => teaData);
  const [expandedTests, setExpandedTests] = useState<Record<string, boolean>>({});
  const [collapsedMacros, setCollapsedMacros] = useState<Record<string, boolean>>({});
  const [collapsedPermissionBlocks, setCollapsedPermissionBlocks] = useState<
    Record<string, boolean>
  >({});
  const [collapsedTeaActivities, setCollapsedTeaActivities] = useState<Record<string, boolean>>({});
  const [collapsedTeaSubActivities, setCollapsedTeaSubActivities] = useState<
    Record<string, boolean>
  >({});
  const [collapsedTeaComposers, setCollapsedTeaComposers] = useState<Record<string, boolean>>({});
  const [collapsedTeaContentBlocks, setCollapsedTeaContentBlocks] = useState<
    Record<string, boolean>
  >({});
  const [activeTab, setActiveTab] = useState<ActiveTab>("document");
  const [teaActiveTab, setTeaActiveTab] = useState<TeaTab>("document");
  const [testBlockFilter, setTestBlockFilter] = useState<TestBlockFilter>("all");
  const [isFaqOpen, setIsFaqOpen] = useState(false);
  const [globalLoading, setGlobalLoading] = useState<LoadingTask | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [isConfirmingImport, setIsConfirmingImport] = useState(false);
  const [isConfirmingAction, setIsConfirmingAction] = useState(false);
  const [isCopyingTeaSubActivity, setIsCopyingTeaSubActivity] = useState(false);
  const [importPreview, setImportPreview] = useState<DocxImportResult | null>(null);
  const [pendingConfirmation, setPendingConfirmation] = useState<PendingConfirmation | null>(null);
  const [teaSubActivityCopyRequest, setTeaSubActivityCopyRequest] =
    useState<TeaSubActivityCopyRequest | null>(null);
  const [draftStatus, setDraftStatus] = useState<DraftStatus>("Rascunho salvo");
  const [permissionBulkText, setPermissionBulkText] = useState(() =>
    formatPermissionBulk(documentData.permissionGroups),
  );
  const documentDataRef = useRef(documentData);
  const teaDataRef = useRef(teaData);
  const documentKindRef = useRef<DocumentKind>(documentKind);
  const permissionBulkTextRef = useRef(permissionBulkText);
  documentDataRef.current = documentData;
  teaDataRef.current = teaData;
  documentKindRef.current = documentKind;
  const saveTimerRef = useRef<number | undefined>();
  const idleSaveRef = useRef<number | undefined>();
  const bufferedCommittersRef = useRef<Set<() => void>>(new Set());
  const isDarkMode = colorScheme === "dark";
  const deferredDocumentData = useDeferredValue(documentData);
  const deferredTeaData = useDeferredValue(teaData);
  const isGlobalLoading = globalLoading !== null;

  const runWithGlobalLoading = useCallback(
    async <T,>(task: LoadingTask, work: () => Promise<T>): Promise<T> => {
      setGlobalLoading(task);

      try {
        return await work();
      } finally {
        setGlobalLoading(null);
      }
    },
    [],
  );

  const selectedGroups = useMemo(
    () => selectedPermissionGroups(documentData.permissionGroups),
    [documentData.permissionGroups],
  );

  const permissionBlockEntries = useMemo(
    () =>
      selectedGroups.flatMap((macro) =>
        macro.microPermissions.map((micro) => ({
          key: createPermissionKey(macro.id, micro.id),
          macro,
          micro,
        })),
      ),
    [selectedGroups],
  );

  const testBlockFilterCounts = useMemo(
    () => buildTestBlockFilterCounts(permissionBlockEntries, documentData.permissionBlocks),
    [documentData.permissionBlocks, permissionBlockEntries],
  );

  const filteredPermissionBlockGroups = useMemo(
    () =>
      selectedGroups
        .map((macro) => {
          const entries = permissionBlockEntries
            .filter((entry) => entry.macro.id === macro.id)
            .reduce<FilteredPermissionBlockEntry[]>((filteredEntries, entry) => {
              const sourceBlock = documentData.permissionBlocks[entry.key] ?? emptyPermissionBlock;

              if (!testBlockMatchesFilter(sourceBlock, testBlockFilter)) {
                return filteredEntries;
              }

              filteredEntries.push({
                ...entry,
                block: filterPermissionBlockTests(sourceBlock, testBlockFilter),
                sourceBlock,
              });

              return filteredEntries;
            }, []);

          return { macro, entries };
        })
        .filter((group) => group.entries.length > 0),
    [documentData.permissionBlocks, permissionBlockEntries, selectedGroups, testBlockFilter],
  );

  const visibleTestCount = useMemo(
    () =>
      filteredPermissionBlockGroups.reduce(
        (total, group) =>
          total +
          group.entries.reduce(
            (entryTotal, entry) => entryTotal + entry.block.tests.length,
            0,
          ),
        0,
      ),
    [filteredPermissionBlockGroups],
  );
  const totalTestCount = testBlockFilterCounts.all;
  const correctionGroups = useMemo(
    () => buildCorrectionGroups(documentData, permissionBlockEntries),
    [documentData, permissionBlockEntries],
  );
  const correctionPendingCount = correctionGroups.filter(
    (group) => !group.correction.corrected,
  ).length;

  const registerBufferedCommit = useCallback((commit: () => void): (() => void) => {
    bufferedCommittersRef.current.add(commit);

    return () => {
      bufferedCommittersRef.current.delete(commit);
    };
  }, []);

  const flushBufferedCommits = useCallback((): void => {
    const committers = Array.from(bufferedCommittersRef.current);

    if (committers.length === 0) {
      return;
    }

    flushSync(() => {
      committers.forEach((commit) => commit());
    });
  }, []);

  const cancelScheduledDraftSave = useCallback(() => {
    if (saveTimerRef.current !== undefined) {
      window.clearTimeout(saveTimerRef.current);
      saveTimerRef.current = undefined;
    }

    if (idleSaveRef.current !== undefined) {
      const idleWindow = window as IdleWindow;

      if (idleWindow.cancelIdleCallback) {
        idleWindow.cancelIdleCallback(idleSaveRef.current);
      } else {
        window.clearTimeout(idleSaveRef.current);
      }

      idleSaveRef.current = undefined;
    }
  }, []);

  const scheduleIdleDraftFlush = useCallback((callback: () => void) => {
    const idleWindow = window as IdleWindow;

    if (idleWindow.requestIdleCallback) {
      idleSaveRef.current = idleWindow.requestIdleCallback(
        () => {
          idleSaveRef.current = undefined;
          callback();
        },
        { timeout: 1200 },
      );
      return;
    }

    idleSaveRef.current = window.setTimeout(() => {
      idleSaveRef.current = undefined;
      callback();
    }, 0);
  }, []);

  const flushDraft = useCallback(() => {
    cancelScheduledDraftSave();

    try {
      setDraftStatus("Salvando...");
      if (documentKindRef.current === "tea") {
        saveTeaDraft(teaDataRef.current);
      } else {
        saveDraft(documentDataRef.current);
      }
      setDraftStatus("Rascunho salvo");
    } catch {
      setDraftStatus("Rascunho grande demais");
    }
  }, [cancelScheduledDraftSave]);

  useEffect(() => {
    if (!draftStatus.startsWith("Altera")) {
      setDraftStatus("Alterações pendentes");
    }

    cancelScheduledDraftSave();

    saveTimerRef.current = window.setTimeout(() => {
      saveTimerRef.current = undefined;
      scheduleIdleDraftFlush(flushDraft);
    }, 700);

    return cancelScheduledDraftSave;
  }, [
    cancelScheduledDraftSave,
    documentData,
    documentKind,
    flushDraft,
    scheduleIdleDraftFlush,
    teaData,
  ]);

  useEffect(() => {
    let isMounted = true;

    async function prepareImages(): Promise<void> {
      try {
        const migratedDocument = await persistEmbeddedEvidenceImages(documentDataRef.current);
        const hydratedDocument = await hydrateDocumentImages(migratedDocument);
        const migratedTeaDocument = await persistEmbeddedTeaImages(teaDataRef.current);
        const hydratedTeaDocument = await hydrateTeaDocumentImages(migratedTeaDocument);

        if (isMounted) {
          setDocumentData(hydratedDocument);
          setTeaData(hydratedTeaDocument);
        }
      } catch {
        // A falta do IndexedDB nao deve impedir o preenchimento do documento atual.
      }
    }

    setGlobalLoading({
      label: "Preparando imagens...",
      detail: "Carregando imagens salvas no navegador.",
    });
    void prepareImages().finally(() => {
      if (isMounted) {
        setGlobalLoading(null);
      }
    });

    return () => {
      isMounted = false;
    };
  // intencional: executa apenas na montagem
  }, []);

  useEffect(() => {
    const handleBeforeUnload = () => {
      flushBufferedCommits();
      flushDraft();
    };

    window.addEventListener("beforeunload", handleBeforeUnload);

    return () => {
      window.removeEventListener("beforeunload", handleBeforeUnload);
    };
  }, [flushBufferedCommits, flushDraft]);

  useEffect(() => {
    permissionBulkTextRef.current = permissionBulkText;
  }, [permissionBulkText]);

  const reviewSummary = useMemo(
    () => buildReviewSummary(deferredDocumentData, permissionBlockEntries),
    [deferredDocumentData, permissionBlockEntries],
  );
  const reviewIssueIndex = useMemo(
    () => buildReviewIssueIndex(reviewSummary.issues),
    [reviewSummary.issues],
  );

  const teaReviewSummary = useMemo(
    () => buildTeaReviewSummary(deferredTeaData),
    [deferredTeaData],
  );
  const teaReviewIssueIndex = useMemo(
    () => buildTeaReviewIssueIndex(teaReviewSummary.issues),
    [teaReviewSummary.issues],
  );

  const refreshOtPreview = useCallback((): void => {
    flushBufferedCommits();
    setOtPreviewDocumentData(documentDataRef.current);
  }, [flushBufferedCommits]);

  const refreshTeaPreview = useCallback((): void => {
    flushBufferedCommits();
    setTeaPreviewDocumentData(teaDataRef.current);
  }, [flushBufferedCommits]);

  const isOtPreviewStale = otPreviewDocumentData !== documentData;
  const isTeaPreviewStale = teaPreviewDocumentData !== teaData;

  const otPreviewModel = useMemo(
    () =>
      documentKind === "ot" && activeTab === "preview"
        ? buildOtPreviewModel(otPreviewDocumentData)
        : null,
    [activeTab, documentKind, otPreviewDocumentData],
  );

  const teaPreviewModel = useMemo(
    () =>
      documentKind === "tea" && teaActiveTab === "preview"
        ? buildTeaPreviewModel(teaPreviewDocumentData)
        : null,
    [documentKind, teaActiveTab, teaPreviewDocumentData],
  );

  const otOutlineContext: ActiveTab | null =
    activeTab === "tests" || activeTab === "preview" ? activeTab : null;
  const otOutlineGroups = useMemo(
    () =>
      otOutlineContext
        ? buildOtOutlineItems(documentData, permissionBlockEntries, reviewSummary, otOutlineContext)
        : [],
    [documentData, otOutlineContext, permissionBlockEntries, reviewSummary],
  );

  const teaOutlineContext: TeaOutlineContext | null =
    teaActiveTab === "activities" || teaActiveTab === "preview" ? teaActiveTab : null;
  const teaOutlineGroups = useMemo(
    () =>
      teaOutlineContext
        ? buildTeaOutlineItems(teaData, teaReviewSummary, teaOutlineContext)
        : [],
    [teaData, teaOutlineContext, teaReviewSummary],
  );
  const showDocumentOutline =
    (documentKind === "ot" && otOutlineContext !== null) || teaOutlineContext !== null;
  const isTeaActivitiesOutlineNavigable =
    documentKind !== "tea" ||
    teaActiveTab !== "activities" ||
    areTeaActivityOutlineTargetsOpen(teaData, collapsedTeaActivities, collapsedTeaSubActivities);
  const isDocumentOutlineNavigationDisabled =
    documentKind === "tea" && teaActiveTab === "activities" && !isTeaActivitiesOutlineNavigable;
  const outlineGroups = documentKind === "tea" ? teaOutlineGroups : otOutlineGroups;
  const outlineTargetIds = useMemo(() => getOutlineTargetIds(outlineGroups), [outlineGroups]);
  const [activeOutlineTargetId, setActiveOutlineTargetId] = useActiveOutlineTargetId(
    outlineTargetIds,
    `${documentKind}:${documentKind === "tea" ? teaActiveTab : activeTab}`,
  );

  const updateDocument = useCallback((updater: (current: OtDocument) => OtDocument): void => {
    setDocumentData((current) => updater(current));
  }, []);

  const updateTeaDocument = useCallback((updater: (current: TeaDocument) => TeaDocument): void => {
    setTeaData((current) => updater(current));
  }, []);

  const requestConfirmation = useCallback<ConfirmAction>((options, onConfirm): void => {
    setPendingConfirmation({
      tone: "danger",
      ...options,
      onConfirm,
    });
  }, []);

  async function confirmPendingAction(): Promise<void> {
    const confirmation = pendingConfirmation;

    if (!confirmation || isConfirmingAction) {
      return;
    }

    setIsConfirmingAction(true);

    try {
      await confirmation.onConfirm();
      setPendingConfirmation(null);
    } finally {
      setIsConfirmingAction(false);
    }
  }

  const setTeaActivityCollapsed = useCallback((activityId: string, collapsed: boolean): void => {
    setCollapsedTeaActivities((current) => setCollapsedMapValue(current, activityId, collapsed));
  }, []);

  const setTeaSubActivityCollapsed = useCallback(
    (subActivityId: string, collapsed: boolean): void => {
      setCollapsedTeaSubActivities((current) =>
        setCollapsedMapValue(current, subActivityId, collapsed),
      );
    },
    [],
  );

  const setTeaComposerCollapsed = useCallback((composerId: string, collapsed: boolean): void => {
    setCollapsedTeaComposers((current) => setCollapsedMapValue(current, composerId, collapsed));
  }, []);

  const setTeaContentBlockCollapsed = useCallback(
    (blockId: string, collapsed: boolean): void => {
      setCollapsedTeaContentBlocks((current) =>
        setCollapsedMapValue(current, blockId, collapsed),
      );
    },
    [],
  );

  function selectDocumentKind(nextKind: DocumentKind): void {
    if (nextKind === documentKindRef.current) {
      return;
    }

    flushBufferedCommits();
    flushDraft();
    setDocumentKind(nextKind);
  }

  const updateTeaMetadata = useCallback((field: keyof TeaDocument["metadata"], value: string): void => {
    updateTeaDocument((current) => ({
      ...current,
      metadata: {
        ...current.metadata,
        [field]: value,
      },
    }));
  }, [updateTeaDocument]);

  const updateTeaActivity = useCallback((
    activityId: string,
    updater: (activity: TeaActivity) => TeaActivity,
  ): void => {
    updateTeaDocument((current) => ({
      ...current,
      activities: current.activities.map((activity) =>
        activity.id === activityId ? updater(activity) : activity,
      ),
    }));
  }, [updateTeaDocument]);

  const addTeaActivity = useCallback((): void => {
    const activityId = createId();

    setTeaActivityCollapsed(activityId, false);
    setTeaComposerCollapsed(activityId, false);
    updateTeaDocument((current) => ({
      ...current,
      activities: [
        ...current.activities,
        {
          id: activityId,
          title: "",
          blocks: [],
          subActivities: [],
        },
      ],
    }));
  }, [setTeaActivityCollapsed, setTeaComposerCollapsed, updateTeaDocument]);

  const removeTeaActivity = useCallback((activityId: string): void => {
    setTeaActivityCollapsed(activityId, false);
    setTeaComposerCollapsed(activityId, false);
    updateTeaDocument((current) => ({
      ...current,
      activities: current.activities.filter((activity) => activity.id !== activityId),
    }));
  }, [setTeaActivityCollapsed, setTeaComposerCollapsed, updateTeaDocument]);

  const moveTeaActivity = useCallback((activityId: string, direction: MoveDirection): void => {
    updateTeaDocument((current) => ({
      ...current,
      activities: moveItemById(current.activities, activityId, direction),
    }));
  }, [updateTeaDocument]);

  const addTeaSubActivity = useCallback((activityId: string): void => {
    const subActivityId = createId();

    setTeaActivityCollapsed(activityId, false);
    setTeaSubActivityCollapsed(subActivityId, false);
    setTeaComposerCollapsed(subActivityId, false);
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: [
        ...activity.subActivities,
        {
          id: subActivityId,
          title: "",
          blocks: [],
        },
      ],
    }));
  }, [
    setTeaActivityCollapsed,
    setTeaComposerCollapsed,
    setTeaSubActivityCollapsed,
    updateTeaActivity,
  ]);

  const updateTeaSubActivity = useCallback((
    activityId: string,
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ): void => {
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: activity.subActivities.map((subActivity) =>
        subActivity.id === subActivityId ? updater(subActivity) : subActivity,
      ),
    }));
  }, [updateTeaActivity]);

  const removeTeaSubActivity = useCallback((activityId: string, subActivityId: string): void => {
    setTeaSubActivityCollapsed(subActivityId, false);
    setTeaComposerCollapsed(subActivityId, false);
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: activity.subActivities.filter(
        (subActivity) => subActivity.id !== subActivityId,
      ),
    }));
  }, [setTeaComposerCollapsed, setTeaSubActivityCollapsed, updateTeaActivity]);

  const moveTeaSubActivity = useCallback((
    activityId: string,
    subActivityId: string,
    direction: MoveDirection,
  ): void => {
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: moveItemById(activity.subActivities, subActivityId, direction),
    }));
  }, [updateTeaActivity]);

  const openTeaSubActivityCopyModal = useCallback((
    sourceActivityId: string,
    subActivityId: string,
  ): void => {
    const targetActivityId =
      teaDataRef.current.activities.find((activity) => activity.id !== sourceActivityId)?.id ??
      null;

    setTeaSubActivityCopyRequest({
      sourceActivityId,
      subActivityId,
      targetActivityId,
    });
  }, []);

  const updateTeaSubActivityCopyTarget = useCallback((targetActivityId: string | null): void => {
    setTeaSubActivityCopyRequest((current) =>
      current ? { ...current, targetActivityId } : current,
    );
  }, []);

  async function confirmTeaSubActivityCopy(): Promise<void> {
    const request = teaSubActivityCopyRequest;

    if (!request?.targetActivityId || isCopyingTeaSubActivity) {
      return;
    }

    setIsCopyingTeaSubActivity(true);

    try {
      const sourceSubActivity = findTeaSubActivity(
        teaDataRef.current,
        request.sourceActivityId,
        request.subActivityId,
      );

      if (!sourceSubActivity) {
        setTeaSubActivityCopyRequest(null);
        return;
      }

      const duplicatedSubActivity = await duplicateTeaSubActivity(sourceSubActivity);

      setTeaSubActivityCopyRequest(null);
      setTeaActiveTab("activities");
      setTeaActivityCollapsed(request.targetActivityId, false);
      setTeaSubActivityCollapsed(duplicatedSubActivity.id, false);
      setTeaComposerCollapsed(duplicatedSubActivity.id, false);
      duplicatedSubActivity.blocks.forEach((block) => {
        setTeaContentBlockCollapsed(block.id, false);
      });

      updateTeaDocument((current) => ({
        ...current,
        activities: current.activities.map((activity) =>
          activity.id === request.targetActivityId
            ? {
                ...activity,
                subActivities: [...activity.subActivities, duplicatedSubActivity],
              }
            : activity,
        ),
      }));

      window.setTimeout(() => {
        scrollAndFocusReviewTarget(`tea-subactivity-${toDomId(duplicatedSubActivity.id)}`);
      }, 80);
    } finally {
      setIsCopyingTeaSubActivity(false);
    }
  }

  const updateTeaOverview = useCallback((overview: string): void => {
    updateTeaDocument((current) => ({ ...current, overview }));
  }, [updateTeaDocument]);

  const updateTeaActivityIntro = useCallback((activityIntro: string): void => {
    updateTeaDocument((current) => ({ ...current, activityIntro }));
  }, [updateTeaDocument]);

  const updateTeaActivityImages = useCallback((
    updater: (images: EvidenceImage[]) => EvidenceImage[],
  ): void => {
    updateTeaDocument((current) => ({
      ...current,
      activityImages: updater(current.activityImages),
    }));
  }, [updateTeaDocument]);

  function updateMetadata(field: keyof OtDocument["metadata"], value: string): void {
    updateDocument((current) => ({
      ...current,
      metadata: {
        ...current.metadata,
        [field]: value,
      },
    }));
  }

  function updateStep(stepId: string, value: string): void {
    updateDocument((current) => ({
      ...current,
      accessSteps: current.accessSteps.map((step) =>
        step.id === stepId ? { ...step, text: value } : step,
      ),
    }));
  }

  function addStep(): void {
    updateDocument((current) => ({
      ...current,
      accessSteps: [...current.accessSteps, { id: createId(), text: "" }],
    }));
  }

  function removeStep(stepId: string): void {
    const step = documentDataRef.current.accessSteps.find((candidate) => candidate.id === stepId);

    requestConfirmation(
      {
        title: "Remover passo?",
        description: `O passo ${formatConfirmationSubject(step?.text ?? "", "selecionado")} será removido do passo a passo.`,
        confirmLabel: "Remover passo",
      },
      () => {
        updateDocument((current) => ({
          ...current,
          accessSteps: current.accessSteps.filter((candidate) => candidate.id !== stepId),
        }));
      },
    );
  }

  function replaceAccessStepsFromBulk(value: string): void {
    const steps = splitBulkLines(value);

    updateDocument((current) => ({
      ...current,
      accessSteps: steps.map((text, index) => ({
        id: current.accessSteps[index]?.id ?? createId(),
        text,
      })),
    }));
  }

  const addMacroGroup = useCallback((): void => {
    updateDocument((current) => ({
      ...current,
      permissionGroups: [
        ...current.permissionGroups,
        {
          id: `macro-${createId()}`,
          code: "",
          label: "",
          selected: true,
          microPermissions: [],
        },
      ],
    }));
  }, [updateDocument]);

  const updateMacroGroup = useCallback((macroId: string, updates: Partial<PermissionItem>): void => {
    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.map((macro) =>
        macro.id === macroId ? { ...macro, ...updates } : macro,
      ),
    }));
  }, [updateDocument]);

  const removeMacroGroup = useCallback((macroId: string): void => {
    const macro = documentDataRef.current.permissionGroups.find(
      (candidate) => candidate.id === macroId,
    );

    requestConfirmation(
      {
        title: "Remover macro?",
        description: `A macro ${formatConfirmationSubject(macro ? formatPermission(macro) : "", "selecionada")} e seus testes vinculados serão removidos.`,
        confirmLabel: "Remover macro",
      },
      () => {
        updateDocument((current) => ({
          ...current,
          permissionGroups: current.permissionGroups.filter((candidate) => candidate.id !== macroId),
          permissionBlocks: removePermissionBlocks(current.permissionBlocks, (key) =>
            key.startsWith(`${macroId}:`),
          ),
        }));
      },
    );
  }, [requestConfirmation, updateDocument]);

  const addMicroPermission = useCallback((macroId: string): void => {
    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.map((macro) =>
        macro.id === macroId
          ? {
              ...macro,
              microPermissions: [
                ...macro.microPermissions,
                {
                  id: `micro-${createId()}`,
                  code: "",
                  label: "",
                  selected: true,
                },
              ],
            }
          : macro,
      ),
    }));
  }, [updateDocument]);

  const updateMicroPermission = useCallback((
    macroId: string,
    microId: string,
    updates: Partial<PermissionItem>,
  ): void => {
    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.map((macro) =>
        macro.id === macroId
          ? {
              ...macro,
              microPermissions: macro.microPermissions.map((micro) =>
                micro.id === microId ? { ...micro, ...updates } : micro,
              ),
            }
          : macro,
      ),
    }));
  }, [updateDocument]);

  const removeMicroPermission = useCallback((macroId: string, microId: string): void => {
    const blockKey = createPermissionKey(macroId, microId);
    const macro = documentDataRef.current.permissionGroups.find(
      (candidate) => candidate.id === macroId,
    );
    const micro = macro?.microPermissions.find((candidate) => candidate.id === microId);

    requestConfirmation(
      {
        title: "Remover micro-permissão?",
        description: `A micro-permissão ${formatConfirmationSubject(micro ? formatPermission(micro) : "", "selecionada")} e seus testes vinculados serão removidos.`,
        confirmLabel: "Remover micro-permissão",
      },
      () => {
        updateDocument((current) => ({
          ...current,
          permissionGroups: current.permissionGroups.map((candidate) =>
            candidate.id === macroId
              ? {
                  ...candidate,
                  microPermissions: candidate.microPermissions.filter(
                    (item) => item.id !== microId,
                  ),
                }
              : candidate,
          ),
          permissionBlocks: removePermissionBlocks(
            current.permissionBlocks,
            (key) => key === blockKey,
          ),
        }));
      },
    );
  }, [requestConfirmation, updateDocument]);

  const updatePermissionBulkDraft = useCallback((value: string): void => {
    permissionBulkTextRef.current = value;
  }, []);

  const commitPermissionBulkDraft = useCallback((): void => {
    setPermissionBulkText(permissionBulkTextRef.current);
  }, []);

  const loadCurrentPermissionBulk = useCallback((): void => {
    const value = formatPermissionBulk(documentDataRef.current.permissionGroups);

    permissionBulkTextRef.current = value;
    setPermissionBulkText(value);
  }, []);

  function applyPermissionBulk(): void {
    const parsedGroups = parsePermissionBulk(permissionBulkTextRef.current);

    if (parsedGroups.length === 0) {
      return;
    }

    commitPermissionBulkDraft();

    updateDocument((current) => {
      const existingBlocks = blocksByPermissionCode(current);
      const permissionGroups = parsedGroups.map((macro) => ({
        ...macro,
        id: `macro-${createId()}`,
        microPermissions: macro.microPermissions.map((micro) => ({
          ...micro,
          id: `micro-${createId()}`,
        })),
      }));
      const permissionBlocks: Record<string, PermissionBlock> = {};

      permissionGroups.forEach((macro) => {
        macro.microPermissions.forEach((micro) => {
          const codeKey = permissionCodeKey(macro.code, micro.code);
          const blockKey = createPermissionKey(macro.id, micro.id);
          permissionBlocks[blockKey] = existingBlocks.get(codeKey) ?? createEmptyBlock();
        });
      });

      return {
        ...current,
        permissionGroups,
        permissionBlocks,
      };
    });
  }

  const setTestExpansion = useCallback((referenceKey: string, expanded: boolean): void => {
    setExpandedTests((current) => ({
      ...current,
      [referenceKey]: expanded,
    }));
  }, []);

  const setMacroCollapsed = useCallback((macroId: string, collapsed: boolean): void => {
    setCollapsedMacros((current) => {
      if (collapsed) {
        return {
          ...current,
          [macroId]: true,
        };
      }

      const { [macroId]: _removed, ...rest } = current;
      return rest;
    });
  }, []);

  const setPermissionBlockCollapsed = useCallback((blockKey: string, collapsed: boolean): void => {
    setCollapsedPermissionBlocks((current) => {
      if (collapsed) {
        return {
          ...current,
          [blockKey]: true,
        };
      }

      const { [blockKey]: _removed, ...rest } = current;
      return rest;
    });
  }, []);

  const expandPermissionPath = useCallback((
    blockKey: string,
  ): void => {
    const macroId = getMacroIdFromBlockKey(blockKey);

    if (macroId) {
      setMacroCollapsed(macroId, false);
    }

    setPermissionBlockCollapsed(blockKey, false);
  }, [setMacroCollapsed, setPermissionBlockCollapsed]);

  const updateBlock = useCallback((
    blockKey: string,
    updater: (block: PermissionBlock) => PermissionBlock,
  ): void => {
    updateDocument((current) => ({
      ...current,
      permissionBlocks: {
        ...current.permissionBlocks,
        [blockKey]: updater(current.permissionBlocks[blockKey] ?? createEmptyBlock()),
      },
    }));
  }, [updateDocument]);

  const addBlockTest = useCallback((blockKey: string): void => {
    const testId = createId();

    expandPermissionPath(blockKey);
    setTestExpansion(createTestReferenceKey(blockKey, testId), true);

    updateBlock(blockKey, (block) => ({
      ...block,
      tests: [...block.tests, createBlockTest(testId)],
    }));
  }, [expandPermissionPath, setTestExpansion, updateBlock]);

  const addStandardTests = useCallback((blockKey: string): void => {
    const tests = standardTestTitles.map((title) => createBlockTest(createId(), title));
    const firstReferenceKey = createTestReferenceKey(blockKey, tests[0]?.id ?? "");

    expandPermissionPath(blockKey);

    if (tests[0]) {
      setTestExpansion(firstReferenceKey, true);
    }

    updateBlock(blockKey, (block) => ({
      ...block,
      tests: [...block.tests, ...tests],
    }));
  }, [expandPermissionPath, setTestExpansion, updateBlock]);

  const addStandardTestsToAll = useCallback((): void => {
    updateDocument((current) => ({
      ...current,
      permissionBlocks: {
        ...current.permissionBlocks,
        ...Object.fromEntries(
          permissionBlockEntries.map((entry) => {
            const block = current.permissionBlocks[entry.key] ?? createEmptyBlock();
            const existingTitles = new Set(
              block.tests.map((test) => normalizeTextKey(test.title)),
            );
            const missingTests = standardTestTitles
              .filter((title) => !existingTitles.has(normalizeTextKey(title)))
              .map((title) => createBlockTest(createId(), title));

            return [entry.key, { ...block, tests: [...block.tests, ...missingTests] }];
          }),
        ),
      },
    }));
  }, [permissionBlockEntries, updateDocument]);

  const duplicateBlockStructureToEmpty = useCallback((sourceBlockKey: string): void => {
    const sourceBlock = documentDataRef.current.permissionBlocks[sourceBlockKey];

    if (!sourceBlock || sourceBlock.tests.length === 0) {
      return;
    }

    updateDocument((current) => ({
      ...current,
      permissionBlocks: {
        ...current.permissionBlocks,
        ...Object.fromEntries(
          permissionBlockEntries
            .filter((entry) => entry.key !== sourceBlockKey)
            .filter((entry) => (current.permissionBlocks[entry.key]?.tests.length ?? 0) === 0)
            .map((entry) => [
              entry.key,
              {
                tests: sourceBlock.tests.map((test) => ({
                  id: createId(),
                  title: test.title,
                  result: {
                    checks: { ...test.result.checks },
                    observations: test.result.observations,
                    legacyImages: [],
                    newImages: [],
                  },
                  correction: createEmptyTestCorrection(),
                })),
              },
            ]),
        ),
      },
    }));
  }, [permissionBlockEntries, updateDocument]);

  const duplicateBlockTest = useCallback((blockKey: string, testId: string): void => {
    const duplicatedId = createId();

    expandPermissionPath(blockKey);
    setTestExpansion(createTestReferenceKey(blockKey, duplicatedId), true);

    updateBlock(blockKey, (block) => {
      const index = block.tests.findIndex((test) => test.id === testId);

      if (index === -1) {
        return block;
      }

      const source = block.tests[index];
      const duplicated: PermissionBlockTest = {
        id: duplicatedId,
        title: source.title,
        result: {
          checks: { ...source.result.checks },
          observations: source.result.observations,
          legacyImages: [],
          newImages: [],
        },
        correction: createEmptyTestCorrection(),
      };
      const tests = [...block.tests];
      tests.splice(index + 1, 0, duplicated);

      return { ...block, tests };
    });
  }, [expandPermissionPath, setTestExpansion, updateBlock]);

  const updateBlockTestTitle = useCallback((
    blockKey: string,
    testId: string,
    title: string,
  ): void => {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId ? { ...test, title } : test,
      ),
    }));
  }, [updateBlock]);

  const removeBlockTest = useCallback((blockKey: string, testId: string): void => {
    const test = documentDataRef.current.permissionBlocks[blockKey]?.tests.find(
      (candidate) => candidate.id === testId,
    );

    requestConfirmation(
      {
        title: "Remover teste?",
        description: `O teste ${formatConfirmationSubject(test?.title ?? "", "selecionado")} será removido deste bloco de permissão.`,
        confirmLabel: "Remover teste",
      },
      () => {
        const referenceKey = createTestReferenceKey(blockKey, testId);

        setExpandedTests((current) => {
          const { [referenceKey]: _removed, ...rest } = current;
          return rest;
        });

        updateBlock(blockKey, (block) => ({
          ...block,
          tests: block.tests.filter((candidate) => candidate.id !== testId),
        }));
      },
    );
  }, [requestConfirmation, updateBlock]);

  const moveBlockTest = useCallback((
    blockKey: string,
    index: number,
    direction: -1 | 1,
  ): void => {
    updateBlock(blockKey, (block) => {
      const nextIndex = index + direction;
      if (nextIndex < 0 || nextIndex >= block.tests.length) {
        return block;
      }

      const tests = [...block.tests];
      const [moved] = tests.splice(index, 1);
      tests.splice(nextIndex, 0, moved);

      return { ...block, tests };
    });
  }, [updateBlock]);

  const updateTestResult = useCallback((
    blockKey: string,
    testId: string,
    updater: (result: TestResult) => TestResult,
  ): void => {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId ? { ...test, result: updater(test.result) } : test,
      ),
    }));
  }, [updateBlock]);

  const updateCorrectionGroup = useCallback((
    groupKey: string,
    updater: (correction: TestCorrection) => TestCorrection,
  ): void => {
    updateDocument((current) => ({
      ...current,
      permissionBlocks: Object.fromEntries(
        Object.entries(current.permissionBlocks).map(([blockKey, block]) => [
          blockKey,
          {
            ...block,
            tests: block.tests.map((test) =>
              test.result.checks.newIssue && getCorrectionGroupKey(blockKey, test) === groupKey
                ? { ...test, correction: updater(getTestCorrection(test)) }
                : test,
            ),
          },
        ]),
      ),
    }));
  }, [updateDocument]);

  function handleStepPaste(stepId: string, event: ClipboardEvent<HTMLInputElement>): void {
    const lines = event.clipboardData
      .getData("text")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length <= 1) {
      return;
    }

    event.preventDefault();

    updateDocument((current) => {
      const stepIndex = current.accessSteps.findIndex((step) => step.id === stepId);

      if (stepIndex === -1) {
        return current;
      }

      const nextSteps = [...current.accessSteps];
      nextSteps.splice(
        stepIndex,
        1,
        { ...nextSteps[stepIndex], text: lines[0] },
        ...lines.slice(1).map((text) => ({ id: createId(), text })),
      );

      return {
        ...current,
        accessSteps: nextSteps,
      };
    });
  }

  const setVisibleTestsExpansion = useCallback((expanded: boolean): void => {
    const visibleMacroIds = filteredPermissionBlockGroups.map((group) => group.macro.id);
    const visibleBlockKeys = filteredPermissionBlockGroups.flatMap((group) =>
      group.entries.map((entry) => entry.key),
    );
    const visibleReferenceKeys = filteredPermissionBlockGroups.flatMap((group) =>
      group.entries.flatMap((entry) =>
        entry.block.tests.map((test) =>
          createTestReferenceKey(entry.key, test.id),
        ),
      ),
    );

    setExpandedTests((current) => {
      const next = { ...current };

      visibleReferenceKeys.forEach((referenceKey) => {
        if (expanded) {
          next[referenceKey] = true;
        } else {
          delete next[referenceKey];
        }
      });

      return next;
    });

    setCollapsedMacros((current) => {
      const next = { ...current };

      visibleMacroIds.forEach((macroId) => {
        if (expanded) {
          delete next[macroId];
        } else {
          next[macroId] = true;
        }
      });

      return next;
    });

    setCollapsedPermissionBlocks((current) => {
      const next = { ...current };

      visibleBlockKeys.forEach((blockKey) => {
        if (expanded) {
          delete next[blockKey];
        } else {
          next[blockKey] = true;
        }
      });

      return next;
    });
  }, [filteredPermissionBlockGroups]);

  function handleReviewIssueClick(issue: ReviewIssue): void {
    setActiveTab(issue.tab);

    if (issue.blockKey) {
      expandPermissionPath(issue.blockKey);
    }

    if (issue.blockKey && issue.testId) {
      setTestExpansion(createTestReferenceKey(issue.blockKey, issue.testId), true);
    }

    window.setTimeout(() => {
      if (issue.targetId) {
        window.document.getElementById(issue.targetId)?.scrollIntoView({
          behavior: "smooth",
          block: "center",
        });
      }
    }, 40);
  }

  function handleOtOutlineItemClick(item: DocumentOutlineItem): void {
    setActiveOutlineTargetId(item.targetId);
    if (item.tab === "preview") {
      refreshOtPreview();
    }
    setActiveTab(item.tab as ActiveTab);

    if (item.tab === "tests") {
      setTestBlockFilter("all");
    }

    if (item.blockKey) {
      expandPermissionPath(item.blockKey);
    }

    if (item.blockKey && item.testId) {
      setTestExpansion(createTestReferenceKey(item.blockKey, item.testId), true);
    }

    window.setTimeout(() => {
      scrollAndFocusReviewTarget(item.targetId);
    }, 60);
  }

  function goToNextIssue(): void {
    const [issue] = reviewSummary.issues;

    if (issue) {
      handleReviewIssueClick(issue);
    }
  }

  const handleTeaReviewIssueClick = useCallback((issue: TeaReviewIssue): void => {
    setTeaActiveTab(issue.tab);

    if (issue.activityId) {
      setTeaActivityCollapsed(issue.activityId, false);
      setTeaComposerCollapsed(issue.activityId, false);
    }

    if (issue.subActivityId) {
      setTeaSubActivityCollapsed(issue.subActivityId, false);
      setTeaComposerCollapsed(issue.subActivityId, false);
    }

    if (issue.blockId) {
      setTeaContentBlockCollapsed(issue.blockId, false);
    }

    window.setTimeout(() => {
      scrollAndFocusReviewTarget(issue.targetId);
    }, 40);
  }, [
    setTeaActivityCollapsed,
    setTeaComposerCollapsed,
    setTeaContentBlockCollapsed,
    setTeaSubActivityCollapsed,
  ]);

  function handleTeaOutlineItemClick(item: DocumentOutlineItem): void {
    setActiveOutlineTargetId(item.targetId);
    if (item.tab === "preview") {
      refreshTeaPreview();
    }
    setTeaActiveTab(item.tab as TeaTab);

    window.setTimeout(() => {
      scrollAndFocusReviewTarget(item.targetId, "start");
    }, 60);
  }

  async function handleImportFile(file: File | null): Promise<void> {
    if (!file || isGlobalLoading) {
      return;
    }

    flushBufferedCommits();
    flushDraft();
    setIsImporting(true);

    try {
      await runWithGlobalLoading(
        {
          label: "Importando DOCX...",
          detail: "Lendo o arquivo e preparando a previa.",
        },
        async () => {
          const { parseDocxFile } = await import("./docxImport");

          setImportPreview(await parseDocxFile(file, documentKindRef.current));
        },
      );
    } catch {
      window.alert("Nao foi possivel importar este DOCX.");
    } finally {
      setIsImporting(false);
    }
  }

  async function confirmImport(): Promise<void> {
    if (!importPreview || isConfirmingImport) {
      return;
    }

    const preview = importPreview;
    setIsConfirmingImport(true);

    try {
      if (preview.kind === "tea") {
        try {
          await persistEmbeddedTeaImages(preview.document);
        } catch {
          window.alert("O TEA foi importado, mas algumas imagens podem nao ficar salvas no rascunho.");
        }

        setTeaData(preview.document);
        setCollapsedTeaActivities({});
        setCollapsedTeaSubActivities({});
        setCollapsedTeaComposers({});
        setCollapsedTeaContentBlocks({});
        setTeaActiveTab("review");
        setImportPreview(null);
        return;
      }

      try {
        await persistEmbeddedEvidenceImages(preview.document);
      } catch {
        window.alert("A OT foi importada, mas algumas imagens podem nao ficar salvas no rascunho.");
      }

      setDocumentData(preview.document);
      setPermissionBulkText(formatPermissionBulk(preview.document.permissionGroups));
      setExpandedTests({});
      setActiveTab("review");
      setImportPreview(null);
    } finally {
      setIsConfirmingImport(false);
    }
  }

  async function handleExport(): Promise<void> {
    if (isGlobalLoading) {
      return;
    }

    flushBufferedCommits();
    flushDraft();
    setIsExporting(true);

    try {
      await runWithGlobalLoading(
        {
          label: "Exportando DOCX...",
          detail: "Otimizando imagens e gerando o arquivo.",
        },
        async () => {
          const { exportOtDocument, exportTeaDocument } = await import("./docxExport");

          if (documentKindRef.current === "tea") {
            await exportTeaDocument(teaDataRef.current);
          } else {
            await exportOtDocument(documentDataRef.current);
          }
        },
      );
    } finally {
      setIsExporting(false);
    }
  }

  async function clearCurrentDraft(): Promise<void> {
    flushBufferedCommits();

    if (documentKindRef.current === "tea") {
      await clearTeaDraft();
      setTeaData(loadTeaDraft());
      setTeaActiveTab("document");
      setCollapsedTeaActivities({});
      setCollapsedTeaSubActivities({});
      setCollapsedTeaComposers({});
      setCollapsedTeaContentBlocks({});
    } else {
      await clearDraft();
      const nextDocument = loadDraft();
      setDocumentData(nextDocument);
      setPermissionBulkText(formatPermissionBulk(nextDocument.permissionGroups));
      setExpandedTests({});
    }

    setDraftStatus("Rascunho salvo");
  }

  function handleClearDraft(): void {
    requestConfirmation(
      {
        title: "Limpar rascunho atual?",
        description:
          "O rascunho local será substituído pelo modelo inicial deste tipo de documento.",
        confirmLabel: "Limpar documento",
      },
      clearCurrentDraft,
    );
  }

  const topBarStatusText = globalLoading?.label ?? draftStatus;
  const topBarStatusColor = globalLoading ? "blue" : draftStatusColor(draftStatus);
  const topBarStatusIcon = isImporting ? (
    <FileUp size={14} />
  ) : isExporting ? (
    <Download size={14} />
  ) : globalLoading ? (
    <FileSearch size={14} />
  ) : (
    <Save size={14} />
  );

  return (
    <ConfirmationContext.Provider value={requestConfirmation}>
    <BufferedCommitContext.Provider value={registerBufferedCommit}>
    <a href="#main-content" className="skipLink">
      Pular para o conteúdo
    </a>
    <main id="main-content" tabIndex={-1}>
    <Container size="xl" py="xl">
      <Stack gap="lg">
        <Paper withBorder p="lg" className="topBar">
          <Group justify="space-between" align="flex-start" gap="md">
            <div className="topBarTitle">
              <Group gap="sm" align="center">
                <Title order={1} size="h2">
                  {documentKind === "tea" ? "Gerador de TEA" : "Gerador de OT"}
                </Title>
                <SegmentedControl
                  value={documentKind}
                  onChange={(value) => selectDocumentKind(value as DocumentKind)}
                  data={[
                    { value: "ot", label: "OT" },
                    { value: "tea", label: "TEA" },
                  ]}
                  size="xs"
                  className="documentKindSwitch"
                />
              </Group>
              <Text c="dimmed" mt={4}>
                {documentKind === "tea"
                  ? teaData.metadata.subject ||
                    teaData.metadata.serviceOrder ||
                    "TEA sem assunto definido"
                  : documentData.metadata.screen || "Documento sem tela definida"}
              </Text>
            </div>

            <Group gap="xs" className="topBarActions actionToolbar">
              <Badge
                variant="light"
                color={topBarStatusColor}
                leftSection={topBarStatusIcon}
                h={30}
                role="status"
                aria-live="polite"
                aria-atomic="true"
              >
                {topBarStatusText}
              </Badge>
              <div className="topBarDesktopOnly">
                <FileButton
                  onChange={(file) => {
                    void handleImportFile(file);
                  }}
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                >
                  {(props) => (
                    <Button
                      {...props}
                      variant="light"
                      leftSection={<FileUp size={17} />}
                      loading={isImporting}
                      disabled={isGlobalLoading && !isImporting}
                    >
                      Importar DOCX
                    </Button>
                  )}
                </FileButton>
              </div>
              {documentKind === "ot" ? (
                <Button
                  variant="light"
                  color="gray"
                  leftSection={<CircleHelp size={17} />}
                  className="topBarDesktopOnly"
                  onClick={() => setIsFaqOpen(true)}
                >
                  Ajuda
                </Button>
              ) : null}
              <Tooltip label={isDarkMode ? "Ativar modo claro" : "Ativar modo escuro"}>
                <ActionIcon
                  variant="light"
                  color="gray"
                  size="lg"
                  className="topBarDesktopOnly"
                  onClick={toggleColorScheme}
                  aria-label={isDarkMode ? "Ativar modo claro" : "Ativar modo escuro"}
                >
                  {isDarkMode ? <Sun size={18} /> : <Moon size={18} />}
                </ActionIcon>
              </Tooltip>
              <Button
                variant="light"
                color="gray"
                leftSection={<RotateCcw size={17} />}
                className="topBarDesktopOnly"
                onClick={handleClearDraft}
                disabled={isGlobalLoading}
              >
                Limpar documento
              </Button>
              <Button
                leftSection={<Download size={17} />}
                onClick={handleExport}
                loading={isExporting}
                disabled={isGlobalLoading && !isExporting}
              >
                Exportar DOCX
              </Button>
              <Menu position="bottom-end" withArrow>
                <Menu.Target>
                  <ActionIcon
                    variant="light"
                    color="gray"
                    size="lg"
                    className="topBarCompactMenu"
                    aria-label="Mais ações do documento"
                  >
                    <MoreVertical size={18} />
                  </ActionIcon>
                </Menu.Target>
                <Menu.Dropdown>
                  <FileButton
                    onChange={(file) => {
                      void handleImportFile(file);
                    }}
                    accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  >
                    {(props) => (
                      <Menu.Item
                        leftSection={<FileUp size={15} />}
                        disabled={isGlobalLoading}
                        onClick={props.onClick}
                      >
                        Importar DOCX
                      </Menu.Item>
                    )}
                  </FileButton>
                  {documentKind === "ot" ? (
                    <Menu.Item
                      leftSection={<CircleHelp size={15} />}
                      onClick={() => setIsFaqOpen(true)}
                    >
                      Ajuda
                    </Menu.Item>
                  ) : null}
                  <Menu.Item
                    leftSection={isDarkMode ? <Sun size={15} /> : <Moon size={15} />}
                    onClick={toggleColorScheme}
                  >
                    {isDarkMode ? "Ativar modo claro" : "Ativar modo escuro"}
                  </Menu.Item>
                  <Menu.Item
                    color="red"
                    leftSection={<RotateCcw size={15} />}
                    onClick={handleClearDraft}
                    disabled={isGlobalLoading}
                  >
                    Limpar documento
                  </Menu.Item>
                </Menu.Dropdown>
              </Menu>
            </Group>
          </Group>
        </Paper>

        {globalLoading ? (
          <LoadingFeedback
            variant="global"
            label={globalLoading.label}
            detail={globalLoading.detail}
          />
        ) : null}

        <div
          className={`workspaceLayout ${showDocumentOutline ? "" : "workspaceLayout--noOutline"}`}
        >
          {showDocumentOutline ? (
            <DocumentOutline
              title="Índice do documento"
              groups={outlineGroups}
              activeTargetId={
                isDocumentOutlineNavigationDisabled ? undefined : activeOutlineTargetId
              }
              isNavigationDisabled={isDocumentOutlineNavigationDisabled}
              onItemClick={
                documentKind === "tea" ? handleTeaOutlineItemClick : handleOtOutlineItemClick
              }
            />
          ) : null}
          <div className="workspaceContent">
        {documentKind === "ot" ? (
        <Tabs
          value={activeTab}
          onChange={(value) => {
            if (value) {
              const nextTab = value as ActiveTab;
              if (nextTab === "preview") {
                refreshOtPreview();
              }
              setActiveTab(nextTab);
            }
          }}
          keepMounted={false}
          className="workspaceTabs"
        >
          <Tabs.List aria-label="Navegação do documento OT">
            <Tabs.Tab value="document" leftSection={<ClipboardList size={16} />}>
              Documento
            </Tabs.Tab>
            <Tabs.Tab value="permissions" leftSection={<ListChecks size={16} />}>
              Permissões
            </Tabs.Tab>
            <Tabs.Tab value="tests" leftSection={<CheckCircle2 size={16} />}>
              Testes
            </Tabs.Tab>
            <Tabs.Tab value="corrections" leftSection={<Wrench size={16} />}>
              <CorrectionTabLabel count={correctionPendingCount} />
            </Tabs.Tab>
            <Tabs.Tab value="review" leftSection={<AlertCircle size={16} />}>
              <ReviewTabLabel count={reviewSummary.issues.length} />
            </Tabs.Tab>
            <Tabs.Tab value="preview" leftSection={<FileSearch size={16} />}>
              Prévia DOCX
            </Tabs.Tab>
          </Tabs.List>

          <Tabs.Panel value="document" pt="md">
            <Stack gap="md">
        <Section title="Documento" tone="document" sectionId="ot-section-document">
          <Stack gap="sm">
            <div className="documentFields">
            <BufferedTextInput
              label="Tela"
              value={documentData.metadata.screen}
              onCommit={(value) => updateMetadata("screen", value)}
            />
            <BufferedTextInput
              label="Responsável pelo teste"
              value={documentData.metadata.responsible}
              onCommit={(value) => updateMetadata("responsible", value)}
            />
            <BufferedTextInput
              label="Data"
              type="date"
              value={documentData.metadata.date}
              onCommit={(value) => updateMetadata("date", value)}
            />
            <BufferedTextInput
              label="Ambiente"
              value={documentData.metadata.environment}
              onCommit={(value) => updateMetadata("environment", value)}
            />
            <BufferedTextInput
              label="Elaborada por"
              value={documentData.metadata.author}
              onCommit={(value) => updateMetadata("author", value)}
            />
            </div>
            <BufferedTextarea
              label="Objetivo"
              minRows={4}
              styles={{ input: { resize: "vertical" } }}
              value={documentData.objective}
              onCommit={(value) =>
                updateDocument((current) => ({
                  ...current,
                  objective: value,
                }))
              }
            />
          </Stack>
        </Section>

        <Section
          sectionId="ot-section-steps"
          title="Passo a passo"
          tone="steps"
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={addStep}>
              Adicionar passo
            </Button>
          }
        >
          <Stack gap="xs">
            <BufferedTextarea
              label="Editar passos em lote"
              minRows={3}
              autosize
              value={documentData.accessSteps.map((step) => step.text).join("\n")}
              onCommit={replaceAccessStepsFromBulk}
              delay={180}
            />
            {documentData.accessSteps.map((step, index) => (
              <Group key={step.id} align="flex-end" wrap="nowrap">
                <Badge color="gray" variant="outline" w={34} h={34}>
                  {index + 1}
                </Badge>
                <BufferedTextInput
                  label={`Passo ${index + 1}`}
                  value={step.text}
                  onCommit={(value) => updateStep(step.id, value)}
                  onPaste={(event) => handleStepPaste(step.id, event)}
                  style={{ flex: 1 }}
                />
                <Tooltip label="Remover passo">
                  <ActionIcon
                    variant="subtle"
                    color="red"
                    onClick={() => removeStep(step.id)}
                    aria-label="Remover passo"
                    mb={index === 0 ? 1 : 0}
                  >
                    <Trash2 size={17} />
                  </ActionIcon>
                </Tooltip>
              </Group>
            ))}
            {documentData.accessSteps.length === 0 ? (
              <EmptyState
                message="Nenhum passo adicionado."
                actionLabel="Adicionar passo"
                onAction={addStep}
              />
            ) : null}
          </Stack>
        </Section>

            </Stack>
          </Tabs.Panel>

          <Tabs.Panel value="permissions" pt="md">
        <Section
          title="Permissões"
          sectionId="ot-section-permissions"
          tone="permissions"
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={addMacroGroup}>
              Adicionar macro
            </Button>
          }
        >
          <Stack gap="sm">
            <Paper withBorder p="sm" className="bulkEditor">
              <Stack gap="xs">
                <Group justify="space-between" align="center">
                  <Text fw={700} size="sm">
                    Lista rápida
                  </Text>
                  <Group gap="xs">
                    <Button
                      variant="subtle"
                      size="xs"
                      onClick={loadCurrentPermissionBulk}
                    >
                      Carregar atual
                    </Button>
                    <Button variant="light" size="xs" onClick={applyPermissionBulk}>
                      Aplicar lista
                    </Button>
                  </Group>
                </Group>
                <Textarea
                  key={permissionBulkText}
                  label="Permissões em lote"
                  description="Use uma macro por linha e micros indentadas abaixo dela."
                  minRows={4}
                  autosize
                  defaultValue={permissionBulkText}
                  placeholder={"AO - Administrador Geral\n  AT - Atualização\n  SC - Somente Consulta"}
                  onBlur={commitPermissionBulkDraft}
                  onChange={(event) => updatePermissionBulkDraft(event.currentTarget.value)}
                />
              </Stack>
            </Paper>
            {documentData.permissionGroups.map((macro, index) => (
              <PermissionGroupEditor
                key={macro.id}
                index={index}
                macro={macro}
                onMacroChange={updateMacroGroup}
                onRemoveMacro={removeMacroGroup}
                onAddMicro={addMicroPermission}
                onMicroChange={updateMicroPermission}
                onRemoveMicro={removeMicroPermission}
              />
            ))}
            {documentData.permissionGroups.length === 0 ? (
              <EmptyState
                message="Nenhuma macro adicionada."
                actionLabel="Adicionar macro"
                onAction={addMacroGroup}
              />
            ) : null}
          </Stack>
        </Section>

          </Tabs.Panel>

          <Tabs.Panel value="tests" pt="md">
        <Section
          title="Testes por permissão"
          sectionId="ot-section-tests"
          tone="blocks"
          action={
            <Badge variant="light" color={testBlockFilter === "all" ? "gray" : "blue"}>
              {visibleTestCount}/{totalTestCount} testes
            </Badge>
          }
        >
          <Stack gap="md">
            <TestBlockFilterBar
              value={testBlockFilter}
              counts={testBlockFilterCounts}
              onChange={setTestBlockFilter}
            />
            <Group gap="xs" justify="flex-end" className="actionToolbar testExpansionActions">
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsDown size={15} />}
                disabled={filteredPermissionBlockGroups.length === 0}
                onClick={() => setVisibleTestsExpansion(true)}
              >
                Expandir todos
              </Button>
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsUp size={15} />}
                disabled={filteredPermissionBlockGroups.length === 0}
                onClick={() => setVisibleTestsExpansion(false)}
              >
                Recolher todos
              </Button>
            </Group>
            {filteredPermissionBlockGroups.map(({ macro, entries }) => (
              <PermissionBlockGroup
                key={macro.id}
                macro={macro}
                entries={entries}
                expandedTests={expandedTests}
                reviewIssueIndex={reviewIssueIndex}
                isCollapsed={collapsedMacros[macro.id] ?? false}
                collapsedBlocks={collapsedPermissionBlocks}
                onMacroCollapseChange={setMacroCollapsed}
                onBlockCollapseChange={setPermissionBlockCollapsed}
                onAddTest={addBlockTest}
                onAddStandardTests={addStandardTests}
                onDuplicateBlockStructure={duplicateBlockStructureToEmpty}
                onDuplicateTest={duplicateBlockTest}
                onTestExpansionChange={setTestExpansion}
                onTestTitleChange={updateBlockTestTitle}
                onTestRemove={removeBlockTest}
                onTestMove={moveBlockTest}
                onResultChange={updateTestResult}
              />
            ))}
            {permissionBlockEntries.length === 0 ? (
              <Paper withBorder p="md" ta="center" className="softEmpty">
                <Text c="dimmed">Nenhuma macro com micro-permissão selecionada.</Text>
              </Paper>
            ) : null}
            {permissionBlockEntries.length > 0 && filteredPermissionBlockGroups.length === 0 ? (
              <Paper withBorder p="md" ta="center" className="softEmpty">
                <Stack gap="xs" align="center">
                  <Text c="dimmed">Nenhum item encontrado para este filtro.</Text>
                  <Button
                    variant="light"
                    size="xs"
                    onClick={() => setTestBlockFilter("all")}
                  >
                    Mostrar todos
                  </Button>
                </Stack>
              </Paper>
            ) : null}
          </Stack>
        </Section>

          </Tabs.Panel>

          <Tabs.Panel value="corrections" pt="md">
            <CorrectionPanel
              groups={correctionGroups}
              onChangeGroup={updateCorrectionGroup}
            />
          </Tabs.Panel>

          <Tabs.Panel value="review" pt="md" id="ot-section-review">
            <ReviewPanel summary={reviewSummary} onIssueClick={handleReviewIssueClick} />
          </Tabs.Panel>

          <Tabs.Panel value="preview" pt="md">
            {activeTab === "preview" && otPreviewModel ? (
              <PreviewPanel
                model={otPreviewModel}
                isStale={isOtPreviewStale}
                onRefresh={refreshOtPreview}
              />
            ) : null}
          </Tabs.Panel>
        </Tabs>
        ) : (
          <TeaWorkspace
            documentData={teaData}
            previewModel={teaPreviewModel}
            isPreviewStale={isTeaPreviewStale}
            activeTab={teaActiveTab}
            reviewSummary={teaReviewSummary}
            reviewIssueIndex={teaReviewIssueIndex}
            collapsedActivities={collapsedTeaActivities}
            collapsedSubActivities={collapsedTeaSubActivities}
            collapsedComposers={collapsedTeaComposers}
            collapsedContentBlocks={collapsedTeaContentBlocks}
            onTabChange={setTeaActiveTab}
            onRefreshPreview={refreshTeaPreview}
            onMetadataChange={updateTeaMetadata}
            onOverviewChange={updateTeaOverview}
            onActivityIntroChange={updateTeaActivityIntro}
            onActivityImagesChange={updateTeaActivityImages}
            onAddActivity={addTeaActivity}
            onActivityChange={updateTeaActivity}
            onActivityRemove={removeTeaActivity}
            onActivityMove={moveTeaActivity}
            onAddSubActivity={addTeaSubActivity}
            onSubActivityChange={updateTeaSubActivity}
            onSubActivityRemove={removeTeaSubActivity}
            onSubActivityMove={moveTeaSubActivity}
            onSubActivityCopy={openTeaSubActivityCopyModal}
            onActivityCollapseChange={setTeaActivityCollapsed}
            onSubActivityCollapseChange={setTeaSubActivityCollapsed}
            onComposerCollapseChange={setTeaComposerCollapsed}
            onContentBlockCollapseChange={setTeaContentBlockCollapsed}
            onReviewIssueClick={handleTeaReviewIssueClick}
          />
        )}
          </div>
        </div>
      </Stack>
    </Container>
    {documentKind === "tea" ? (
      <Tooltip label="Voltar ao topo" position="left">
        <ActionIcon
          aria-label="Voltar ao topo"
          className="teaBackToTopButton"
          size="xl"
          radius="xl"
          onClick={() => {
            window.scrollTo({ top: 0, behavior: "smooth" });
          }}
        >
          <ArrowUp size={20} />
        </ActionIcon>
      </Tooltip>
    ) : null}
    </main>
    <ImportPreviewModal
      result={importPreview}
      isConfirming={isConfirmingImport}
      onClose={() => {
        if (!isConfirmingImport) {
          setImportPreview(null);
        }
      }}
      onConfirm={() => {
        void confirmImport();
      }}
    />
    {isFaqOpen ? (
      <Modal
        opened
        onClose={() => setIsFaqOpen(false)}
        title="Ajuda do Gerador de OT"
        size="lg"
      >
        <FaqPanel />
      </Modal>
    ) : null}
    <ConfirmationModal
      confirmation={pendingConfirmation}
      isConfirming={isConfirmingAction}
      onCancel={() => {
        if (!isConfirmingAction) {
          setPendingConfirmation(null);
        }
      }}
      onConfirm={() => {
        void confirmPendingAction();
      }}
    />
    <TeaSubActivityCopyModal
      documentData={teaData}
      request={teaSubActivityCopyRequest}
      isCopying={isCopyingTeaSubActivity}
      onTargetActivityChange={updateTeaSubActivityCopyTarget}
      onClose={() => {
        if (!isCopyingTeaSubActivity) {
          setTeaSubActivityCopyRequest(null);
        }
      }}
      onConfirm={() => {
        void confirmTeaSubActivityCopy();
      }}
    />
    </BufferedCommitContext.Provider>
    </ConfirmationContext.Provider>
  );
}

const TeaWorkspace = memo(function TeaWorkspace({
  documentData,
  previewModel,
  isPreviewStale,
  activeTab,
  reviewSummary,
  reviewIssueIndex,
  collapsedActivities,
  collapsedSubActivities,
  collapsedComposers,
  collapsedContentBlocks,
  onTabChange,
  onMetadataChange,
  onOverviewChange,
  onActivityIntroChange,
  onActivityImagesChange,
  onAddActivity,
  onActivityChange,
  onActivityRemove,
  onActivityMove,
  onAddSubActivity,
  onSubActivityChange,
  onSubActivityRemove,
  onSubActivityMove,
  onSubActivityCopy,
  onActivityCollapseChange,
  onSubActivityCollapseChange,
  onComposerCollapseChange,
  onContentBlockCollapseChange,
  onReviewIssueClick,
  onRefreshPreview,
}: TeaWorkspaceProps) {
  const overview = useBufferedText(documentData.overview, onOverviewChange, 180);
  const activityIntro = useBufferedText(documentData.activityIntro, onActivityIntroChange, 180);
  const activityCount = documentData.activities.length;
  const hasActivityReviewIssues = reviewIssueIndex.hasActivityIssues;
  const inlineError = useCallback(
    (targetId: string): string | undefined =>
      getTeaInlineReviewError(
        getTeaReviewIssuesByKey(reviewIssueIndex.byTargetId, targetId),
        targetId,
      ),
    [reviewIssueIndex],
  );

  function setAllActivityPanelsCollapsed(collapsed: boolean): void {
    documentData.activities.forEach((activity) => {
      onActivityCollapseChange(activity.id, collapsed);
      onComposerCollapseChange(activity.id, collapsed);
      activity.blocks.forEach((block) => onContentBlockCollapseChange(block.id, collapsed));

      activity.subActivities.forEach((subActivity) => {
        onSubActivityCollapseChange(subActivity.id, collapsed);
        onComposerCollapseChange(subActivity.id, collapsed);
        subActivity.blocks.forEach((block) => onContentBlockCollapseChange(block.id, collapsed));
      });
    });
  }

  function expandTeaPendingPanels(): void {
    setAllActivityPanelsCollapsed(true);

    reviewIssueIndex.activityIssues
      .forEach((issue) => {
        if (!issue.activityId) {
          return;
        }

        onActivityCollapseChange(issue.activityId, false);

        if (issue.subActivityId) {
          onSubActivityCollapseChange(issue.subActivityId, false);
          onComposerCollapseChange(issue.subActivityId, false);
        } else {
          onComposerCollapseChange(issue.activityId, false);
        }

        if (issue.blockId) {
          onContentBlockCollapseChange(issue.blockId, false);
        }
      });
  }

  return (
    <Tabs
      value={activeTab}
      onChange={(value) => {
        if (value) {
          const nextTab = value as TeaTab;
          if (nextTab === "preview") {
            onRefreshPreview();
          }
          onTabChange(nextTab);
        }
      }}
      keepMounted={false}
      className="workspaceTabs"
    >
      <Tabs.List aria-label="Navegação do documento TEA">
        <Tabs.Tab value="document" leftSection={<FileText size={16} />}>
          Documento
        </Tabs.Tab>
        <Tabs.Tab value="activities" leftSection={<ListChecks size={16} />}>
          Atividades
        </Tabs.Tab>
        <Tabs.Tab value="review" leftSection={<AlertCircle size={16} />}>
          <ReviewTabLabel count={reviewSummary.issues.length} />
        </Tabs.Tab>
        <Tabs.Tab value="preview" leftSection={<FileSearch size={16} />}>
          Prévia DOCX
        </Tabs.Tab>
      </Tabs.List>

      <Tabs.Panel value="document" pt="md">
        <Stack gap="md">
          <Section title="Documento TEA" tone="document" sectionId="tea-section-document">
            <Stack gap="sm">
              <div className="documentFields">
                <BufferedTextInput
                  label="Ordem de Serviço"
                  id="tea-metadata-service-order"
                  value={documentData.metadata.serviceOrder}
                  error={inlineError("tea-metadata-service-order")}
                  placeholder="OS2171 - Login/Menu/Prestador/Documentos Vencidos"
                  onCommit={(value) => onMetadataChange("serviceOrder", value)}
                />
                <BufferedTextInput
                  label="Fase/Etapa"
                  id="tea-metadata-phase"
                  value={documentData.metadata.phase}
                  error={inlineError("tea-metadata-phase")}
                  placeholder="Etapa 5"
                  onCommit={(value) => onMetadataChange("phase", value)}
                />
                <BufferedTextInput
                  label="Chamado"
                  id="tea-metadata-ticket"
                  value={documentData.metadata.ticket}
                  error={inlineError("tea-metadata-ticket")}
                  placeholder="Chamado 202504000396"
                  onCommit={(value) => onMetadataChange("ticket", value)}
                />
                <BufferedTextInput
                  label="Assunto"
                  id="tea-metadata-subject"
                  value={documentData.metadata.subject}
                  error={inlineError("tea-metadata-subject")}
                  placeholder="Telas - Novo Layout"
                  onCommit={(value) => onMetadataChange("subject", value)}
                />
                <BufferedTextInput
                  label="Data"
                  id="tea-metadata-date"
                  type="date"
                  value={documentData.metadata.date}
                  error={inlineError("tea-metadata-date")}
                  onCommit={(value) => onMetadataChange("date", value)}
                />
                <BufferedTextInput
                  label="Elaborado por"
                  id="tea-metadata-author"
                  value={documentData.metadata.author}
                  error={inlineError("tea-metadata-author")}
                  onCommit={(value) => onMetadataChange("author", value)}
                />
              </div>
              <Textarea
                label="1. Visão geral"
                minRows={5}
                autosize
                styles={{ input: { resize: "vertical" } }}
                id="tea-overview"
                value={overview.value}
                error={inlineError("tea-overview")}
                onChange={(event) => overview.setValue(event.currentTarget.value)}
                onBlur={overview.commit}
              />
              <Textarea
                label="Texto inicial das atividades"
                minRows={2}
                autosize
                styles={{ input: { resize: "vertical" } }}
                id="tea-activity-intro"
                value={activityIntro.value}
                error={inlineError("tea-activity-intro")}
                onChange={(event) => activityIntro.setValue(event.currentTarget.value)}
                onBlur={activityIntro.commit}
              />
            </Stack>
          </Section>

          <Section title="Imagem geral" tone="steps" sectionId="tea-section-images">
            <EvidenceUploader
              title="Atividades realizadas"
              tone="new"
              images={documentData.activityImages}
              onChange={onActivityImagesChange}
            />
          </Section>
        </Stack>
      </Tabs.Panel>

      <Tabs.Panel value="activities" pt="md">
        <Section
          sectionId="tea-section-activities"
          title="Atividades"
          tone="blocks"
          action={
            <Group gap="xs" className="actionToolbar sectionActionToolbar">
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsDown size={15} />}
                disabled={activityCount === 0}
                onClick={() => setAllActivityPanelsCollapsed(false)}
              >
                Expandir todos
              </Button>
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsUp size={15} />}
                disabled={activityCount === 0}
                onClick={() => setAllActivityPanelsCollapsed(true)}
              >
                Recolher todos
              </Button>
              <Button
                variant="light"
                size="xs"
                leftSection={<AlertCircle size={15} />}
                disabled={!hasActivityReviewIssues}
                onClick={expandTeaPendingPanels}
              >
                Expandir pendências
              </Button>
              <Button variant="light" leftSection={<Plus size={17} />} onClick={onAddActivity}>
                Adicionar atividade
              </Button>
            </Group>
          }
        >
          <Stack gap="md">
            {documentData.activities.map((activity, index) => (
              <TeaActivityEditor
                key={activity.id}
                index={index}
                totalActivities={documentData.activities.length}
                activity={activity}
                reviewIssues={getTeaReviewIssuesByKey(
                  reviewIssueIndex.byActivityId,
                  activity.id,
                )}
                isCollapsed={collapsedActivities[activity.id] ?? false}
                collapsedSubActivities={collapsedSubActivities}
                collapsedComposers={collapsedComposers}
                collapsedContentBlocks={collapsedContentBlocks}
                onChange={(updater) => onActivityChange(activity.id, updater)}
                onRemove={() => onActivityRemove(activity.id)}
                onMove={(direction) => onActivityMove(activity.id, direction)}
                onAddSubActivity={() => onAddSubActivity(activity.id)}
                onSubActivityChange={(subActivityId, updater) =>
                  onSubActivityChange(activity.id, subActivityId, updater)
                }
                onSubActivityRemove={(subActivityId) =>
                  onSubActivityRemove(activity.id, subActivityId)
                }
                onSubActivityMove={(subActivityId, direction) =>
                  onSubActivityMove(activity.id, subActivityId, direction)
                }
                onSubActivityCopy={(subActivityId) =>
                  onSubActivityCopy(activity.id, subActivityId)
                }
                onCollapseChange={(collapsed) =>
                  onActivityCollapseChange(activity.id, collapsed)
                }
                onSubActivityCollapseChange={onSubActivityCollapseChange}
                onComposerCollapseChange={onComposerCollapseChange}
                onContentBlockCollapseChange={onContentBlockCollapseChange}
              />
            ))}

            {documentData.activities.length === 0 ? (
              <EmptyState
                message="Nenhuma atividade adicionada."
                actionLabel="Adicionar atividade"
                onAction={onAddActivity}
              />
            ) : null}
          </Stack>
        </Section>
      </Tabs.Panel>

      <Tabs.Panel value="review" pt="md" id="tea-section-review">
        <TeaReviewPanel summary={reviewSummary} onIssueClick={onReviewIssueClick} />
      </Tabs.Panel>

      <Tabs.Panel value="preview" pt="md">
        {activeTab === "preview" && previewModel ? (
          <PreviewPanel
            model={previewModel}
            isStale={isPreviewStale}
            onRefresh={onRefreshPreview}
          />
        ) : null}
      </Tabs.Panel>
    </Tabs>
  );
});

const TeaActivityEditor = memo(function TeaActivityEditor({
  index,
  totalActivities,
  activity,
  reviewIssues,
  isCollapsed,
  collapsedSubActivities,
  collapsedComposers,
  collapsedContentBlocks,
  onChange,
  onRemove,
  onMove,
  onAddSubActivity,
  onSubActivityChange,
  onSubActivityRemove,
  onSubActivityMove,
  onSubActivityCopy,
  onCollapseChange,
  onSubActivityCollapseChange,
  onComposerCollapseChange,
  onContentBlockCollapseChange,
}: TeaActivityEditorProps) {
  const activityId = `tea-activity-${toDomId(activity.id)}`;
  const titleInputId = `tea-activity-title-${toDomId(activity.id)}`;
  const panelId = `${activityId}-panel`;
  const isExpanded = !isCollapsed;
  const review = summarizeTeaReviewIssues(reviewIssues);
  const summaryItems = buildTeaActivitySummaryItems(activity);
  const titleError = getTeaInlineReviewError(reviewIssues, titleInputId);
  const activityNumber = `2.${index + 1}`;
  const hasTitle = Boolean(activity.title.trim());
  const headerTitle = formatTeaEditorTitle(
    activityNumber,
    activity.title,
    "Atividade sem título",
  );
  const composerTitle = formatTeaComposerTitle("activity", activityNumber, activity.title);
  const confirmAction = useConfirmAction();
  const commitTitle = useCallback(
    (title: string) => onChange((current) => ({ ...current, title })),
    [onChange],
  );
  const title = useBufferedText(activity.title, commitTitle);

  const updateBlocks = useCallback((updater: (blocks: TeaContentBlock[]) => TeaContentBlock[]): void => {
    onChange((current) => ({ ...current, blocks: updater(current.blocks) }));
  }, [onChange]);
  const {
    rootIssues: activityRootReviewIssues,
    bySubActivityId: subActivityReviewIssuesById,
  } = useMemo(
    () => splitTeaActivityReviewIssues(reviewIssues),
    [reviewIssues],
  );

  function confirmAndRemove(): void {
    if (!hasTeaActivityRemovalContent(activity)) {
      onRemove();
      return;
    }

    confirmAction(
      {
        title: "Remover atividade?",
        description: `A atividade ${formatConfirmationSubject(activity.title, "selecionada")} e todo o seu conteúdo serão removidos do documento.`,
        confirmLabel: "Remover atividade",
      },
      onRemove,
    );
  }

  return (
    <Paper
      id={activityId}
      tabIndex={-1}
      withBorder
      p="md"
      className={`teaActivityCard ${isExpanded ? "teaActivityCard--expanded" : "teaActivityCard--collapsed"}`}
    >
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="teaActivityHeader">
          <Group gap="xs" wrap="nowrap" className="teaActivityTitle">
            <Tooltip label={isExpanded ? "Recolher atividade" : "Abrir atividade"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onCollapseChange(isExpanded)}
                aria-label={isExpanded ? "Recolher atividade" : "Abrir atividade"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${isExpanded ? "testToggleIcon--open" : ""}`}
                />
              </ActionIcon>
            </Tooltip>
            <div className="teaHeaderCopy">
              <Text
                fw={800}
                className={`teaPrimaryTitle ${hasTitle ? "" : "teaPrimaryTitle--missing"}`}
              >
                {headerTitle}
              </Text>
              <TeaHeaderMeta items={["Atividade", ...summaryItems]} />
              <TeaSummaryChips items={[]} review={review} />
            </div>
          </Group>
          <TeaActivityActionsMenu
            canMoveUp={index > 0}
            canMoveDown={index < totalActivities - 1}
            onMove={onMove}
            onRemove={confirmAndRemove}
          />
        </Group>

        <LazyCollapse in={isExpanded}>
          <div id={panelId} className="collapseBody">
            <Stack gap="sm">
        <TextInput
          label="Título"
          id={titleInputId}
          value={title.value}
          error={titleError}
          placeholder="Seletor Situação do Prestador"
          onChange={(event) => {
            title.setValue(event.currentTarget.value);
          }}
          onBlur={title.commit}
        />
        <TeaContentComposer
          composerId={activity.id}
          title={composerTitle}
          description={`${formatTeaCount(activity.blocks.length, "bloco", "blocos")} nesta atividade`}
          emptyMessage="Nenhum bloco adicionado nesta atividade."
          emptyActionLabel="Adicionar primeiro bloco"
          blocks={activity.blocks}
          tone="new"
          isCollapsed={collapsedComposers[activity.id] ?? false}
          collapsedBlocks={collapsedContentBlocks}
          reviewIssues={activityRootReviewIssues}
          onCollapseChange={(collapsed) => onComposerCollapseChange(activity.id, collapsed)}
          onBlockCollapseChange={onContentBlockCollapseChange}
          onBlocksChange={updateBlocks}
        />

        <Divider className="teaSubActivityDivider" />

        <Group justify="space-between" align="center" className="teaSubActivityGroupHeader">
          <Text fw={700} size="sm">
            Subtópicos
          </Text>
          <Button
            variant="subtle"
            size="xs"
            leftSection={<Plus size={15} />}
            onClick={onAddSubActivity}
          >
            Adicionar subtópico
          </Button>
        </Group>

        <Stack gap="sm" className="teaSubActivityList">
          {activity.subActivities.map((subActivity, subIndex) => (
            <TeaSubActivityEditor
              key={subActivity.id}
              activityIndex={index}
              index={subIndex}
              totalSubActivities={activity.subActivities.length}
              subActivity={subActivity}
              reviewIssues={getTeaReviewIssuesByKey(
                subActivityReviewIssuesById,
                subActivity.id,
              )}
              isCollapsed={collapsedSubActivities[subActivity.id] ?? false}
              isComposerCollapsed={collapsedComposers[subActivity.id] ?? false}
              collapsedBlocks={collapsedContentBlocks}
              onChange={(updater) => onSubActivityChange(subActivity.id, updater)}
              onRemove={() => onSubActivityRemove(subActivity.id)}
              onMove={(direction) => onSubActivityMove(subActivity.id, direction)}
              onCopy={() => onSubActivityCopy(subActivity.id)}
              onCollapseChange={(collapsed) =>
                onSubActivityCollapseChange(subActivity.id, collapsed)
              }
              onComposerCollapseChange={(collapsed) =>
                onComposerCollapseChange(subActivity.id, collapsed)
              }
              onContentBlockCollapseChange={onContentBlockCollapseChange}
            />
          ))}
        </Stack>
            </Stack>
          </div>
        </LazyCollapse>
      </Stack>
    </Paper>
  );
}, areTeaActivityEditorPropsEqual);

const TeaSubActivityEditor = memo(function TeaSubActivityEditor({
  activityIndex,
  index,
  totalSubActivities,
  subActivity,
  reviewIssues,
  isCollapsed,
  isComposerCollapsed,
  collapsedBlocks,
  onChange,
  onRemove,
  onMove,
  onCopy,
  onCollapseChange,
  onComposerCollapseChange,
  onContentBlockCollapseChange,
}: TeaSubActivityEditorProps) {
  const subActivityId = `tea-subactivity-${toDomId(subActivity.id)}`;
  const titleInputId = `tea-subactivity-title-${toDomId(subActivity.id)}`;
  const panelId = `${subActivityId}-panel`;
  const isExpanded = !isCollapsed;
  const review = summarizeTeaReviewIssues(reviewIssues);
  const summaryItems = buildTeaSubActivitySummaryItems(subActivity);
  const titleError = getTeaInlineReviewError(reviewIssues, titleInputId);
  const subActivityNumber = `2.${activityIndex + 1}.${index + 1}`;
  const hasTitle = Boolean(subActivity.title.trim());
  const headerTitle = formatTeaEditorTitle(
    subActivityNumber,
    subActivity.title,
    "Subtópico sem título",
  );
  const composerTitle = formatTeaComposerTitle("subActivity", subActivityNumber, subActivity.title);
  const confirmAction = useConfirmAction();
  const commitTitle = useCallback(
    (title: string) => onChange((current) => ({ ...current, title })),
    [onChange],
  );
  const title = useBufferedText(subActivity.title, commitTitle);

  const updateBlocks = useCallback((updater: (blocks: TeaContentBlock[]) => TeaContentBlock[]): void => {
    onChange((current) => ({ ...current, blocks: updater(current.blocks) }));
  }, [onChange]);

  function confirmAndRemove(): void {
    if (!hasTeaSubActivityRemovalContent(subActivity)) {
      onRemove();
      return;
    }

    confirmAction(
      {
        title: "Remover subtópico?",
        description: `O subtópico ${formatConfirmationSubject(subActivity.title, "selecionado")} e todo o seu conteúdo serão removidos da atividade.`,
        confirmLabel: "Remover subtópico",
      },
      onRemove,
    );
  }

  return (
    <Paper id={subActivityId} p="sm" className="teaSubActivityCard" tabIndex={-1}>
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="teaSubActivityHeader">
          <Group gap="xs" wrap="nowrap" className="teaSubActivityTitle">
            <Tooltip label={isExpanded ? "Recolher subtópico" : "Abrir subtópico"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onCollapseChange(isExpanded)}
                aria-label={isExpanded ? "Recolher subtópico" : "Abrir subtópico"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${isExpanded ? "testToggleIcon--open" : ""}`}
                />
              </ActionIcon>
            </Tooltip>
            <div className="teaHeaderCopy">
              <Text
                fw={800}
                size="sm"
                className={`teaPrimaryTitle ${hasTitle ? "" : "teaPrimaryTitle--missing"}`}
              >
                {headerTitle}
              </Text>
              <TeaHeaderMeta items={["Subtópico", ...summaryItems]} />
              <TeaSummaryChips items={[]} review={review} />
            </div>
          </Group>
          <TeaSubActivityActionsMenu
            canMoveUp={index > 0}
            canMoveDown={index < totalSubActivities - 1}
            onMove={onMove}
            onCopy={onCopy}
            onRemove={confirmAndRemove}
          />
          <Group gap={4} wrap="nowrap" className="teaLegacyActionsHidden">
            <Tooltip label="Mover subtópico para cima">
              <ActionIcon
                variant="subtle"
                onClick={() => onMove("up")}
                disabled={index === 0}
                aria-label="Mover subtópico para cima"
              >
                <ArrowUp size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Mover subtópico para baixo">
              <ActionIcon
                variant="subtle"
                onClick={() => onMove("down")}
                disabled={index >= totalSubActivities - 1}
                aria-label="Mover subtópico para baixo"
              >
                <ArrowDown size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Remover subtópico">
              <ActionIcon
                variant="subtle"
                color="red"
                onClick={confirmAndRemove}
                aria-label="Remover subtópico"
              >
                <Trash2 size={17} />
              </ActionIcon>
            </Tooltip>
          </Group>
        </Group>

        <LazyCollapse in={isExpanded}>
          <div id={panelId} className="collapseBody">
            <Stack gap="sm">
        <TextInput
          label="Título"
          id={titleInputId}
          value={title.value}
          error={titleError}
          placeholder="Botão de Anexo"
          onChange={(event) => {
            title.setValue(event.currentTarget.value);
          }}
          onBlur={title.commit}
        />
        <TeaContentComposer
          title={composerTitle}
          description={`${formatTeaCount(subActivity.blocks.length, "bloco", "blocos")} neste subtópico`}
          emptyMessage="Nenhum bloco adicionado neste subtópico."
          emptyActionLabel="Adicionar primeiro bloco"
          composerId={subActivity.id}
          blocks={subActivity.blocks}
          tone="legacy"
          isCollapsed={isComposerCollapsed}
          collapsedBlocks={collapsedBlocks}
          reviewIssues={reviewIssues}
          onCollapseChange={onComposerCollapseChange}
          onBlockCollapseChange={onContentBlockCollapseChange}
          onBlocksChange={updateBlocks}
        />
            </Stack>
          </div>
        </LazyCollapse>
      </Stack>
    </Paper>
  );
}, areTeaSubActivityEditorPropsEqual);

const TeaContentComposer = memo(function TeaContentComposer({
  composerId,
  title,
  description,
  emptyMessage,
  emptyActionLabel,
  blocks,
  tone,
  isCollapsed,
  collapsedBlocks,
  reviewIssues,
  onCollapseChange,
  onBlockCollapseChange,
  onBlocksChange,
}: TeaContentComposerProps) {
  const panelId = `tea-composer-${toDomId(composerId)}-panel`;
  const isExpanded = !isCollapsed;
  const confirmAction = useConfirmAction();
  const blockReviewIssuesById = useMemo(
    () => groupTeaReviewIssuesByBlock(reviewIssues),
    [reviewIssues],
  );

  function addBlock(type: TeaContentBlockType): void {
    onCollapseChange(false);
    onBlocksChange((current) => [...current, createTeaContentBlock(type)]);
  }

  function updateBlock(
    blockId: string,
    updater: (block: TeaContentBlock) => TeaContentBlock,
  ): void {
    onBlocksChange((current) =>
      current.map((block) => (block.id === blockId ? updater(block) : block)),
    );
  }

  function moveBlock(blockId: string, direction: MoveDirection): void {
    onBlocksChange((current) => moveItemById(current, blockId, direction));
  }

  function removeBlock(block: TeaContentBlock): void {
    function removeCurrentBlock(): void {
      if (block.type === "images") {
        void deleteEvidenceImageDataBatch(block.images.map((image) => image.id));
      }

      onBlocksChange((current) => current.filter((candidate) => candidate.id !== block.id));
    }

    if (!hasMeaningfulTeaBlockContent(block)) {
      removeCurrentBlock();
      return;
    }

    confirmAction(
      {
        title: "Remover bloco?",
        description: `${getTeaBlockDisplayLabel(block)} será removido do documento.`,
        confirmLabel: "Remover bloco",
      },
      removeCurrentBlock,
    );
  }

  async function duplicateBlock(block: TeaContentBlock): Promise<void> {
    const duplicate = await duplicateTeaContentBlock(block);

    onBlocksChange((current) => {
      const sourceIndex = current.findIndex((candidate) => candidate.id === block.id);
      const insertIndex = sourceIndex >= 0 ? sourceIndex + 1 : current.length;

      return [
        ...current.slice(0, insertIndex),
        duplicate,
        ...current.slice(insertIndex),
      ];
    });
  }

  return (
    <Paper p="sm" className="teaComposer">
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="teaComposerHeader">
          <Group gap="xs" wrap="nowrap" className="teaComposerTitle">
            <Tooltip label={isExpanded ? "Recolher conteudo" : "Abrir conteudo"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onCollapseChange(isExpanded)}
                aria-label={isExpanded ? "Recolher conteudo" : "Abrir conteudo"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${isExpanded ? "testToggleIcon--open" : ""}`}
                />
              </ActionIcon>
            </Tooltip>
          <div>
            <Text fw={800} size="sm">
              {title}
            </Text>
            <Text c="dimmed" size="xs">
              {description}
            </Text>
          </div>
          </Group>
          <TeaAddBlockMenu onAddBlock={addBlock} />
        </Group>

        <LazyCollapse in={isExpanded}>
          <div id={panelId} className="collapseBody">
        {blocks.length > 0 ? (
          <Stack gap="xs">
            {blocks.map((block, blockIndex) => (
              <TeaContentBlockEditor
                key={block.id}
                block={block}
                index={blockIndex}
                totalBlocks={blocks.length}
                tone={tone}
                isCollapsed={collapsedBlocks[block.id] ?? false}
                reviewIssues={getTeaReviewIssuesByKey(blockReviewIssuesById, block.id)}
                onChange={(updater) => updateBlock(block.id, updater)}
                onMove={(direction) => moveBlock(block.id, direction)}
                onDuplicate={() => duplicateBlock(block)}
                onRemove={() => removeBlock(block)}
                onCollapseChange={(collapsed) => onBlockCollapseChange(block.id, collapsed)}
              />
            ))}
          </Stack>
        ) : (
          <Paper p="md" ta="center" className="teaComposerEmpty">
            <Stack gap="xs" align="center">
              <Text c="dimmed">{emptyMessage}</Text>
              <TeaAddBlockMenu onAddBlock={addBlock} label={emptyActionLabel} />
            </Stack>
          </Paper>
        )}
          </div>
        </LazyCollapse>
      </Stack>
    </Paper>
  );
}, areTeaContentComposerPropsEqual);

const TeaContentBlockEditor = memo(function TeaContentBlockEditor({
  block,
  index,
  totalBlocks,
  tone,
  isCollapsed,
  reviewIssues,
  onChange,
  onMove,
  onDuplicate,
  onRemove,
  onCollapseChange,
}: TeaContentBlockEditorProps) {
  const blockId = `tea-content-block-${toDomId(block.id)}`;
  const blockInputId = getTeaContentBlockInputId(block);
  const panelId = `${blockId}-panel`;
  const label = teaContentBlockLabels[block.type];
  const review = summarizeTeaReviewIssues(reviewIssues);
  const summaryItems = [buildTeaContentBlockSummaryItem(block)];
  const blockInputError = blockInputId
    ? getTeaInlineReviewError(reviewIssues, blockInputId)
    : undefined;
  const [isDuplicating, setIsDuplicating] = useState(false);
  const commitText = useCallback(
    (text: string) =>
      onChange((current) => (current.type === "text" ? { ...current, text } : current)),
    [onChange],
  );
  const text = useBufferedText(block.type === "text" ? block.text : "", commitText);
  const listValue = block.type === "list" ? block.items.map((item) => item.text).join("\n") : "";
  const commitList = useCallback(
    (value: string) =>
      onChange((current) =>
        current.type === "list"
          ? { ...current, items: updateTeaItemsFromBulk(current.items, value) }
          : current,
      ),
    [onChange],
  );
  const list = useBufferedText(listValue, commitList);

  async function duplicateCurrentBlock(): Promise<void> {
    if (isDuplicating) {
      return;
    }

    setIsDuplicating(true);

    try {
      await onDuplicate();
    } finally {
      setIsDuplicating(false);
    }
  }

  return (
    <Paper id={blockId} p="sm" className="teaContentBlock" tabIndex={-1}>
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="teaContentBlockHeader">
          <Group gap="xs" wrap="nowrap" className="teaContentBlockTitle">
            <Tooltip label={isCollapsed ? "Abrir bloco" : "Recolher bloco"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onCollapseChange(!isCollapsed)}
                disabled={isDuplicating}
                aria-label={isCollapsed ? "Abrir bloco" : "Recolher bloco"}
                aria-expanded={!isCollapsed}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${!isCollapsed ? "testToggleIcon--open" : ""}`}
                />
              </ActionIcon>
            </Tooltip>
            <span className={`teaBlockIcon teaBlockIcon--${block.type}`}>
              {block.type === "text" ? <FileText size={16} /> : null}
              {block.type === "list" ? <ListChecks size={16} /> : null}
              {block.type === "images" ? <ImagePlus size={16} /> : null}
            </span>
            <div className="teaHeaderCopy">
              <Text fw={800} size="sm">
                Bloco {index + 1} - {label}
              </Text>
              <TeaHeaderMeta items={summaryItems} />
              <TeaSummaryChips items={[]} review={review} />
            </div>
          </Group>

          <TeaBlockActionsMenu
            canMoveUp={index > 0}
            canMoveDown={index < totalBlocks - 1}
            isBusy={isDuplicating}
            onDuplicate={() => {
              void duplicateCurrentBlock();
            }}
            onMove={onMove}
            onRemove={onRemove}
          />
        </Group>

        {isDuplicating ? (
          <LoadingFeedback
            variant="compact"
            label="Duplicando bloco..."
            detail={block.type === "images" ? "Copiando imagens no rascunho." : undefined}
          />
        ) : null}

        <LazyCollapse in={!isCollapsed}>
          <div id={panelId} className="collapseBody">
            {block.type === "text" ? (
              <Textarea
                label="Texto"
                id={blockInputId}
                minRows={4}
                autosize
                error={blockInputError}
                styles={{ input: { resize: "vertical" } }}
                value={text.value}
                onChange={(event) => {
                  text.setValue(event.currentTarget.value);
                }}
                onBlur={text.commit}
              />
            ) : null}

            {block.type === "list" ? (
              <Textarea
                label="Itens em lista"
                id={blockInputId}
                minRows={3}
                autosize
                error={blockInputError}
                placeholder="Um item por linha"
                styles={{ input: { resize: "vertical" } }}
                value={list.value}
                onChange={(event) => {
                  list.setValue(event.currentTarget.value);
                }}
                onBlur={list.commit}
              />
            ) : null}

            {block.type === "images" ? (
              <EvidenceUploader
                title="Imagens"
                tone={tone}
                images={block.images}
                onChange={(updater) =>
                  onChange((current) =>
                    current.type === "images"
                      ? { ...current, images: updater(current.images) }
                      : current,
                  )
                }
              />
            ) : null}
          </div>
        </LazyCollapse>
      </Stack>
    </Paper>
  );
}, areTeaContentBlockEditorPropsEqual);

function SummaryChips({
  items,
  review,
}: {
  items: string[];
  review?: InlineReviewSummary;
}) {
  const summary = review ?? emptyInlineReviewSummary;
  const issueTone = getInlineReviewTone(summary);

  return (
    <div className="summaryChips">
      {items.map((item) => (
        <Badge key={item} size="xs" variant="outline" color="gray" className="summaryChip">
          {item}
        </Badge>
      ))}
      <Badge
        size="xs"
        variant={summary.total > 0 ? "light" : "outline"}
        color={issueTone}
        className={`summaryChip summaryChip--${issueTone}`}
      >
        {formatOtCount(summary.total, "pendencia", "pendencias")}
      </Badge>
    </div>
  );
}

function CheckStatusChips({ result }: { result: TestResult }) {
  const selectedKeys = getSelectedCheckKeys(result);

  return (
    <div className="statusChips">
      {selectedKeys.length > 0 ? (
        selectedKeys.map((key) => (
          <Badge
            key={key}
            size="xs"
            variant="light"
            color={getCheckTone(key)}
            leftSection={renderCheckIcon(key, 13)}
            className="statusChip"
          >
            {quickCheckLabels[key]}
          </Badge>
        ))
      ) : (
        <Badge
          size="xs"
          variant="outline"
          color="gray"
          leftSection={<AlertCircle size={13} />}
          className="statusChip"
        >
          Sem status
        </Badge>
      )}
    </div>
  );
}

function TeaHeaderMeta({ items }: { items: string[] }) {
  return (
    <Text c="dimmed" size="xs" className="teaHeaderMeta">
      {items.join(" • ")}
    </Text>
  );
}

function TeaSummaryChips({
  items,
  review,
}: {
  items: string[];
  review: TeaInlineReviewSummary;
}) {
  const issueTone = getTeaReviewTone(review);
  const showReviewChip = review.total > 0;

  if (items.length === 0 && !showReviewChip) {
    return null;
  }

  return (
    <div className="teaSummaryChips">
      {items.map((item) => (
        <Badge key={item} size="xs" variant="outline" color="gray" className="teaSummaryChip">
          {item}
        </Badge>
      ))}
      {showReviewChip ? (
        <Badge
          size="xs"
          variant="light"
          color={issueTone}
          className={`teaSummaryChip teaSummaryChip--${issueTone}`}
        >
          {formatTeaCount(review.total, "pendência", "pendências")}
        </Badge>
      ) : null}
    </div>
  );
}

function TeaAddBlockMenu({
  onAddBlock,
  label = "Adicionar bloco",
}: {
  onAddBlock: (type: TeaContentBlockType) => void;
  label?: string;
}) {
  const [opened, setOpened] = useState(false);

  function addBlock(type: TeaContentBlockType): void {
    onAddBlock(type);
    setOpened(false);
  }

  return (
    <Menu opened={opened} onChange={setOpened} position="bottom-end" withArrow>
      <Menu.Target>
        <span className="teaMenuTarget">
          <Button
            variant="light"
            size="xs"
            leftSection={<Plus size={15} />}
            className="teaAddBlockButton"
            onClick={(event) => {
              event.stopPropagation();
              setOpened(true);
            }}
          >
            {label}
          </Button>
        </span>
      </Menu.Target>
      <Menu.Dropdown>
        <Menu.Item leftSection={<FileText size={15} />} onClick={() => addBlock("text")}>
          Texto
        </Menu.Item>
        <Menu.Item leftSection={<ListChecks size={15} />} onClick={() => addBlock("list")}>
          Lista
        </Menu.Item>
        <Menu.Item leftSection={<ImagePlus size={15} />} onClick={() => addBlock("images")}>
          Imagens
        </Menu.Item>
      </Menu.Dropdown>
    </Menu>
  );
}

function TeaActivityActionsMenu({
  canMoveUp,
  canMoveDown,
  onMove,
  onRemove,
}: {
  canMoveUp: boolean;
  canMoveDown: boolean;
  onMove: (direction: MoveDirection) => void;
  onRemove: () => void;
}) {
  return (
    <TeaMoveRemoveMenu
      ariaLabel="Mais ações da atividade"
      upLabel="Mover atividade para cima"
      downLabel="Mover atividade para baixo"
      removeLabel="Remover atividade"
      canMoveUp={canMoveUp}
      canMoveDown={canMoveDown}
      onMove={onMove}
      onRemove={onRemove}
    />
  );
}

function TeaSubActivityActionsMenu({
  canMoveUp,
  canMoveDown,
  onMove,
  onCopy,
  onRemove,
}: {
  canMoveUp: boolean;
  canMoveDown: boolean;
  onMove: (direction: MoveDirection) => void;
  onCopy: () => void;
  onRemove: () => void;
}) {
  return (
    <TeaMoveRemoveMenu
      ariaLabel="Mais ações do subtópico"
      upLabel="Mover subtópico para cima"
      downLabel="Mover subtópico para baixo"
      removeLabel="Remover subtópico"
      canMoveUp={canMoveUp}
      canMoveDown={canMoveDown}
      onMove={onMove}
      copyLabel="Copiar subtópico"
      onCopy={onCopy}
      onRemove={onRemove}
    />
  );
}

function TeaMoveRemoveMenu({
  ariaLabel,
  upLabel,
  downLabel,
  copyLabel,
  removeLabel,
  canMoveUp,
  canMoveDown,
  onMove,
  onCopy,
  onRemove,
}: {
  ariaLabel: string;
  upLabel: string;
  downLabel: string;
  copyLabel?: string;
  removeLabel: string;
  canMoveUp: boolean;
  canMoveDown: boolean;
  onMove: (direction: MoveDirection) => void;
  onCopy?: () => void;
  onRemove: () => void;
}) {
  const [opened, setOpened] = useState(false);

  function copy(): void {
    onCopy?.();
    setOpened(false);
  }

  function move(direction: MoveDirection): void {
    onMove(direction);
    setOpened(false);
  }

  function remove(): void {
    onRemove();
    setOpened(false);
  }

  return (
    <Menu opened={opened} onChange={setOpened} position="bottom-end" withArrow>
      <Menu.Target>
        <span className="teaMenuTarget">
          <ActionIcon
            variant="subtle"
            aria-label={ariaLabel}
            className="teaActionsMenu"
            onClick={(event) => {
              event.stopPropagation();
              setOpened(true);
            }}
          >
            <MoreVertical size={17} />
          </ActionIcon>
        </span>
      </Menu.Target>
      <Menu.Dropdown>
        {onCopy ? (
          <>
            <Menu.Item leftSection={<Copy size={15} />} onClick={copy}>
              {copyLabel ?? "Copiar"}
            </Menu.Item>
            <Menu.Divider />
          </>
        ) : null}
        <Menu.Item
          leftSection={<ArrowUp size={15} />}
          disabled={!canMoveUp}
          onClick={() => move("up")}
        >
          {upLabel}
        </Menu.Item>
        <Menu.Item
          leftSection={<ArrowDown size={15} />}
          disabled={!canMoveDown}
          onClick={() => move("down")}
        >
          {downLabel}
        </Menu.Item>
        <Menu.Divider />
        <Menu.Item color="red" leftSection={<Trash2 size={15} />} onClick={remove}>
          {removeLabel}
        </Menu.Item>
      </Menu.Dropdown>
    </Menu>
  );
}

function TeaBlockActionsMenu({
  canMoveUp,
  canMoveDown,
  isBusy,
  onDuplicate,
  onMove,
  onRemove,
}: {
  canMoveUp: boolean;
  canMoveDown: boolean;
  isBusy: boolean;
  onDuplicate: () => void;
  onMove: (direction: MoveDirection) => void;
  onRemove: () => void;
}) {
  const [opened, setOpened] = useState(false);

  function duplicate(): void {
    if (isBusy) {
      return;
    }

    onDuplicate();
    setOpened(false);
  }

  function move(direction: MoveDirection): void {
    if (isBusy) {
      return;
    }

    onMove(direction);
    setOpened(false);
  }

  function remove(): void {
    if (isBusy) {
      return;
    }

    onRemove();
    setOpened(false);
  }

  return (
    <Menu opened={opened} onChange={setOpened} position="bottom-end" withArrow>
      <Menu.Target>
        <span className="teaMenuTarget">
          <ActionIcon
            variant="subtle"
            aria-label="Mais ações do bloco"
            className="teaActionsMenu"
            disabled={isBusy}
            onClick={(event) => {
              event.stopPropagation();
              if (isBusy) {
                return;
              }
              setOpened(true);
            }}
          >
            <MoreVertical size={17} />
          </ActionIcon>
        </span>
      </Menu.Target>
      <Menu.Dropdown>
        <Menu.Item leftSection={<Copy size={15} />} disabled={isBusy} onClick={duplicate}>
          Duplicar bloco
        </Menu.Item>
        <Menu.Item
          leftSection={<ArrowUp size={15} />}
          disabled={isBusy || !canMoveUp}
          onClick={() => move("up")}
        >
          Mover bloco para cima
        </Menu.Item>
        <Menu.Item
          leftSection={<ArrowDown size={15} />}
          disabled={isBusy || !canMoveDown}
          onClick={() => move("down")}
        >
          Mover bloco para baixo
        </Menu.Item>
        <Menu.Divider />
        <Menu.Item
          color="red"
          leftSection={<Trash2 size={15} />}
          disabled={isBusy}
          onClick={remove}
        >
          Remover bloco
        </Menu.Item>
      </Menu.Dropdown>
    </Menu>
  );
}

function getTeaInlineReviewError(issues: TeaReviewIssue[], targetId: string): string | undefined {
  const issue =
    issues.find((candidate) => candidate.targetId === targetId && candidate.severity === "danger") ??
    issues.find((candidate) => candidate.targetId === targetId);

  return issue ? `${issue.label}: ${issue.detail}` : undefined;
}

function buildTeaReviewCategoryGroups(issues: TeaReviewIssue[]): Array<{
  category: TeaReviewCategory;
  title: string;
  emptyMessage: string;
  issues: TeaReviewIssue[];
}> {
  const groups: Array<{
    category: TeaReviewCategory;
    title: string;
    emptyMessage: string;
  }> = [
    {
      category: "document",
      title: "Documento",
      emptyMessage: "Os dados principais do documento estão completos.",
    },
    {
      category: "activity",
      title: "Atividades",
      emptyMessage: "As atividades e subtópicos não têm pendências.",
    },
    {
      category: "image",
      title: "Imagens",
      emptyMessage: "As imagens necessárias estão disponíveis.",
    },
  ];

  return groups.map((group) => ({
    ...group,
    issues: issues.filter((issue) => issue.category === group.category),
  }));
}

function getTeaContentBlockInputId(block: TeaContentBlock): string | undefined {
  if (block.type !== "text" && block.type !== "list") {
    return undefined;
  }

  return `tea-content-block-input-${toDomId(block.id)}`;
}

function getTeaContentBlockReviewTargetId(block: TeaContentBlock): string {
  return getTeaContentBlockInputId(block) ?? `tea-content-block-${toDomId(block.id)}`;
}

function ReviewTabLabel({ count }: { count: number }) {
  return (
    <span className="reviewTabLabel">
      <span>Revisão</span>
      {count > 0 ? (
        <span className="reviewTabBadge" aria-label={`${count} pendências de revisão`}>
          {count > 99 ? "99+" : count}
        </span>
      ) : null}
    </span>
  );
}

function CorrectionTabLabel({ count }: { count: number }) {
  return (
    <span className="reviewTabLabel">
      <span>Para corrigir</span>
      {count > 0 ? (
        <span className="reviewTabBadge" aria-label={`${count} correcoes pendentes`}>
          {count > 99 ? "99+" : count}
        </span>
      ) : null}
    </span>
  );
}

function PreviewPanel({
  model,
  isStale,
  onRefresh,
}: {
  model: DocxPreviewModel;
  isStale: boolean;
  onRefresh: () => void;
}) {
  return (
    <Stack gap="sm">
      <Paper withBorder p="sm" className="previewToolbar">
        <Group justify="space-between" align="center" gap="sm">
          <Group gap="xs">
            <Text fw={750} size="sm">
              Prévia DOCX
            </Text>
            <Badge color={isStale ? "yellow" : "green"} variant="light" aria-live="polite">
              {isStale ? "Desatualizada" : "Atualizada"}
            </Badge>
          </Group>
          <Button
            variant="light"
            size="xs"
            leftSection={<RotateCcw size={15} />}
            onClick={onRefresh}
          >
            Atualizar prévia
          </Button>
        </Group>
      </Paper>
      <Suspense fallback={<PreviewSkeleton />}>
        <DocxPreview model={model} />
      </Suspense>
    </Stack>
  );
}

function PreviewSkeleton() {
  return (
    <section className="docxPreviewSection docxPreviewSkeleton" aria-hidden="true">
      <div className="docxPreviewShell">
        <div className="docxPreviewSkeletonPage" />
      </div>
    </section>
  );
}

function TeaReviewPanel({
  summary,
  onIssueClick,
}: {
  summary: TeaReviewSummary;
  onIssueClick: (issue: TeaReviewIssue) => void;
}) {
  const totalIssues = summary.issues.length;
  const dangerIssues = summary.issues.filter((issue) => issue.severity === "danger");
  const categoryGroups = buildTeaReviewCategoryGroups(summary.issues);

  return (
    <Section
      title="Revisão"
      tone="document"
      action={
        <Badge
          color={dangerIssues.length > 0 ? "red" : totalIssues > 0 ? "yellow" : "green"}
          variant="light"
        >
          {totalIssues > 0 ? "Pendências" : "Pronto"}
        </Badge>
      }
    >
      <Stack gap="md">
        <div className="reviewMetrics">
          <ReviewMetric label="Atividades" value={summary.activityCount} />
          <ReviewMetric label="Imagens" value={summary.imageCount} />
          <ReviewMetric label="Pendências" value={totalIssues} />
        </div>

        {totalIssues > 0 ? (
          <div className="reviewIssues">
            {categoryGroups.map((group) => (
              <TeaReviewCategorySection
                key={group.category}
                title={group.title}
                emptyMessage={group.emptyMessage}
                issues={group.issues}
                onIssueClick={onIssueClick}
              />
            ))}
          </div>
        ) : (
          <Group gap="xs" className="reviewOk">
            <CheckCircle2 size={18} />
            <Text fw={700}>Nenhuma pendência encontrada.</Text>
          </Group>
        )}
      </Stack>
    </Section>
  );
}

function DocumentOutline({
  title,
  groups,
  activeTargetId,
  isNavigationDisabled = false,
  onItemClick,
}: {
  title: string;
  groups: DocumentOutlineGroup[];
  activeTargetId?: string;
  isNavigationDisabled?: boolean;
  onItemClick: (item: DocumentOutlineItem) => void;
}) {
  const [isOpen, setIsOpen] = useState(true);
  const itemRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const pendingCount = groups.reduce(
    (total, group) =>
      total + group.items.filter((item) => item.status === "pending").length,
    0,
  );

  useEffect(() => {
    if (!activeTargetId) {
      return;
    }

    itemRefs.current[activeTargetId]?.scrollIntoView?.({
      block: "nearest",
    });
  }, [activeTargetId]);

  return (
    <aside className="documentOutline" aria-label={title}>
      <Paper withBorder p="sm" className="documentOutlinePanel">
        <Group justify="space-between" align="center" gap="xs" className="documentOutlineHeader">
          <div>
            <Text fw={800} size="sm">
              Índice
            </Text>
            <Text c="dimmed" size="xs">
              Documento inteiro
            </Text>
          </div>
          <Group gap="xs" wrap="nowrap">
            {pendingCount > 0 ? (
              <Badge size="xs" color="red" variant="light">
                {pendingCount}
              </Badge>
            ) : null}
            <Button
              variant="subtle"
              size="xs"
              className="documentOutlineToggle"
              onClick={() => setIsOpen((current) => !current)}
              aria-expanded={isOpen}
            >
              {isOpen ? "Ocultar" : "Mostrar"}
            </Button>
          </Group>
        </Group>

        <Collapse in={isOpen}>
          <nav className="documentOutlineNav" aria-label={title}>
            {groups.map((group) => (
              <div key={group.id} className="documentOutlineGroup">
                <Text
                  size="xs"
                  fw={800}
                  c="dimmed"
                  tt="uppercase"
                  className="documentOutlineGroupTitle"
                >
                  {group.title}
                </Text>
                <div className="documentOutlineItems">
                  {group.items.map((item) => {
                    const isActive = Boolean(item.targetId && item.targetId === activeTargetId);

                    return (
                      <button
                        key={item.id}
                        type="button"
                        ref={(element) => {
                          if (item.targetId) {
                            itemRefs.current[item.targetId] = element;
                          }
                        }}
                        className={`documentOutlineItem documentOutlineItem--level${item.level ?? 0}`}
                        data-active={isActive ? "true" : undefined}
                        aria-current={isActive ? "location" : undefined}
                        disabled={isNavigationDisabled}
                        onClick={() => onItemClick(item)}
                      >
                        <span className="documentOutlineItemCopy">
                          <span className="documentOutlineItemTitle">{item.title}</span>
                        </span>
                        {item.status === "pending" ? (
                          <span className="documentOutlinePending">Pendente</span>
                        ) : null}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </nav>
        </Collapse>
      </Paper>
    </aside>
  );
}

type LazyCollapseProps = Omit<ComponentProps<typeof Collapse>, "children"> & {
  children: ReactNode;
};

function LazyCollapse({
  in: isOpen,
  children,
  transitionDuration = 200,
  animateOpacity = false,
  onTransitionEnd,
  ...props
}: LazyCollapseProps) {
  const [shouldRender, setShouldRender] = useState(isOpen);

  useEffect(() => {
    if (isOpen) {
      setShouldRender(true);
      return undefined;
    }

    const timeoutId = window.setTimeout(() => {
      setShouldRender(false);
    }, transitionDuration + 40);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [isOpen, transitionDuration]);

  return (
    <Collapse
      {...props}
      in={isOpen}
      animateOpacity={animateOpacity}
      transitionDuration={transitionDuration}
      onTransitionEnd={() => {
        onTransitionEnd?.();
        if (!isOpen) {
          setShouldRender(false);
        }
      }}
    >
      {shouldRender ? children : null}
    </Collapse>
  );
}

function Section({
  title,
  tone,
  action,
  sectionId,
  children,
}: {
  title: string;
  tone: "document" | "steps" | "permissions" | "blocks";
  action?: ReactNode;
  sectionId?: string;
  children: ReactNode;
}) {
  const titleId = useId();

  return (
    <Card
      withBorder
      component="section"
      id={sectionId}
      aria-labelledby={titleId}
      padding="lg"
      radius="md"
      className={`sectionCard sectionCard--${tone}`}
    >
      <Group justify="space-between" mb="md" align="center" className="sectionHeader">
        <Title id={titleId} order={2} size="h4">
          {title}
        </Title>
        {action ? <div className="sectionActions">{action}</div> : null}
      </Group>
      {children}
    </Card>
  );
}

type BufferedTextInputProps = Omit<
  ComponentProps<typeof TextInput>,
  "value" | "onChange"
> & {
  value: string;
  onCommit: (value: string) => void;
  delay?: number;
};

function BufferedTextInput({
  value,
  onCommit,
  delay,
  onBlur,
  ...props
}: BufferedTextInputProps) {
  const buffered = useBufferedText(value, onCommit, delay);

  return (
    <TextInput
      {...props}
      value={buffered.value}
      onChange={(event) => buffered.setValue(event.currentTarget.value)}
      onBlur={(event) => {
        buffered.commit();
        onBlur?.(event);
      }}
    />
  );
}

type BufferedTextareaProps = Omit<
  ComponentProps<typeof Textarea>,
  "value" | "onChange"
> & {
  value: string;
  onCommit: (value: string) => void;
  delay?: number;
};

function BufferedTextarea({
  value,
  onCommit,
  delay,
  onBlur,
  ...props
}: BufferedTextareaProps) {
  const buffered = useBufferedText(value, onCommit, delay);

  return (
    <Textarea
      {...props}
      value={buffered.value}
      onChange={(event) => buffered.setValue(event.currentTarget.value)}
      onBlur={(event) => {
        buffered.commit();
        onBlur?.(event);
      }}
    />
  );
}

function TestBlockFilterBar({
  value,
  counts,
  onChange,
}: {
  value: TestBlockFilter;
  counts: Record<TestBlockFilter, number>;
  onChange: (value: TestBlockFilter) => void;
}) {
  const selectedCountLabel =
    value === "withoutTests"
      ? `${counts[value]} bloco${counts[value] === 1 ? "" : "s"}`
      : `${counts[value]} teste${counts[value] === 1 ? "" : "s"}`;

  return (
    <Paper withBorder p="sm" className="testBlockFilterBar">
      <Stack gap="xs">
        <Group justify="space-between" align="center" gap="sm">
          <div>
            <Text fw={750} size="sm">
              Filtrar testes por situação
            </Text>
            <Text c="dimmed" size="xs">
              Reduza a lista para revisar só os testes que precisam de ação.
            </Text>
          </div>
          <Badge variant="outline" color={value === "all" ? "gray" : "blue"}>
            {selectedCountLabel}
          </Badge>
        </Group>

        <div className="testBlockFilters" role="tablist" aria-label="Filtros de testes">
          {testBlockFilterOrder.map((option) => {
            const isActive = value === option;
            const isEmpty = option !== "all" && counts[option] === 0;

            return (
              <button
                key={option}
                type="button"
                role="tab"
                aria-selected={isActive}
                className={`testBlockFilter ${isActive ? "testBlockFilter--active" : ""}`}
                disabled={!isActive && isEmpty}
                onClick={() => onChange(option)}
              >
                <span>{testBlockFilterLabels[option]}</span>
                <strong>{counts[option]}</strong>
              </button>
            );
          })}
        </div>
      </Stack>
    </Paper>
  );
}

const PermissionGroupEditor = memo(function PermissionGroupEditor({
  index,
  macro,
  onMacroChange,
  onRemoveMacro,
  onAddMicro,
  onMicroChange,
  onRemoveMicro,
}: PermissionGroupEditorProps) {
  const macroId = macro.id;
  const handleMacroSelectedChange = useCallback((event: ChangeEvent<HTMLInputElement>): void => {
    onMacroChange(macroId, { selected: event.currentTarget.checked });
  }, [macroId, onMacroChange]);
  const handleMacroCodeCommit = useCallback((value: string): void => {
    onMacroChange(macroId, { code: value });
  }, [macroId, onMacroChange]);
  const handleMacroLabelCommit = useCallback((value: string): void => {
    onMacroChange(macroId, { label: value });
  }, [macroId, onMacroChange]);
  const handleRemoveMacro = useCallback((): void => {
    onRemoveMacro(macroId);
  }, [macroId, onRemoveMacro]);
  const handleAddMicro = useCallback((): void => {
    onAddMicro(macroId);
  }, [macroId, onAddMicro]);

  return (
    <Paper id={`permission-macro-${toDomId(macro.id)}`} withBorder p="md" className="permissionGroup">
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="permissionGroupHeader">
          <Group gap="xs">
            <Badge variant="outline" color="gray">
              Macro {index + 1}
            </Badge>
            <Checkbox
              label="Usar"
              checked={macro.selected}
              onChange={handleMacroSelectedChange}
            />
          </Group>
          <Tooltip label="Remover macro">
            <ActionIcon
              variant="subtle"
              color="red"
              onClick={handleRemoveMacro}
              aria-label="Remover macro"
            >
              <Trash2 size={17} />
            </ActionIcon>
          </Tooltip>
        </Group>

        <div className="permissionFields">
          <BufferedTextInput
            label="Código"
            value={macro.code}
            placeholder="AO"
            onCommit={handleMacroCodeCommit}
          />
          <BufferedTextInput
            label="Descrição"
            value={macro.label}
            placeholder="Administrador Geral"
            onCommit={handleMacroLabelCommit}
          />
        </div>

        <Divider />

        <Group justify="space-between" align="center" className="permissionMicroHeader">
          <Text fw={700} size="sm">
            Micro-permissões
          </Text>
          <Button variant="subtle" size="xs" leftSection={<Plus size={15} />} onClick={handleAddMicro}>
            Adicionar micro-permissão
          </Button>
        </Group>

        <Stack gap="xs">
          {macro.microPermissions.map((micro) => (
            <div className="microPermissionRow" key={micro.id}>
              <Checkbox
                label="Usar"
                checked={micro.selected}
                onChange={(event) =>
                  onMicroChange(macroId, micro.id, { selected: event.currentTarget.checked })
                }
              />
              <BufferedTextInput
                label="Código"
                value={micro.code}
                placeholder="AT"
                onCommit={(value) => onMicroChange(macroId, micro.id, { code: value })}
              />
              <BufferedTextInput
                label="Descrição"
                value={micro.label}
                placeholder="Atualização"
                onCommit={(value) => onMicroChange(macroId, micro.id, { label: value })}
              />
              <Tooltip label="Remover micro">
                <ActionIcon
                  variant="subtle"
                  color="red"
                  onClick={() => onRemoveMicro(macroId, micro.id)}
                  aria-label="Remover micro"
                  mt={24}
                >
                  <Trash2 size={17} />
                </ActionIcon>
              </Tooltip>
            </div>
          ))}
        </Stack>

        {macro.microPermissions.length === 0 ? (
          <Paper withBorder p="sm" ta="center" className="softEmpty">
            <Text c="dimmed" size="sm">
              Nenhuma micro-permissão nesta macro.
            </Text>
          </Paper>
        ) : null}
      </Stack>
    </Paper>
  );
}, arePermissionGroupEditorPropsEqual);

const PermissionBlockGroup = memo(function PermissionBlockGroup({
  macro,
  entries,
  expandedTests,
  reviewIssueIndex,
  isCollapsed,
  collapsedBlocks,
  onMacroCollapseChange,
  onBlockCollapseChange,
  onAddTest,
  onAddStandardTests,
  onDuplicateBlockStructure,
  onDuplicateTest,
  onTestExpansionChange,
  onTestTitleChange,
  onTestRemove,
  onTestMove,
  onResultChange,
}: PermissionBlockGroupProps) {
  const isExpanded = !isCollapsed;
  const panelId = `macro-tests-${toDomId(macro.id)}`;

  return (
    <Paper withBorder p="md" className="blockGroup">
      <Stack gap="sm">
        <Group justify="space-between" align="center" className="blockGroupHeader">
          <Group gap="xs" align="center" wrap="nowrap" className="blockGroupTitle">
            <Tooltip label={isExpanded ? "Recolher macro" : "Abrir macro"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onMacroCollapseChange(macro.id, isExpanded)}
                aria-label={isExpanded ? "Recolher macro" : "Abrir macro"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${
                    isExpanded ? "testToggleIcon--open" : ""
                  }`}
                />
              </ActionIcon>
            </Tooltip>
            <div>
              <Text size="xs" c="dimmed" fw={700} tt="uppercase">
                Macro
              </Text>
              <Title order={3} size="h5">
                {formatPermission(macro)}
              </Title>
            </div>
          </Group>
          <Badge variant="outline" color="gray">
            {entries.length} micro{entries.length === 1 ? "" : "s"}
          </Badge>
        </Group>

        <Collapse in={isExpanded}>
          {isExpanded ? (
          <Stack gap="sm" id={panelId}>
            {entries.map((entry) => (
              <PermissionBlockEditor
                key={entry.key}
                blockKey={entry.key}
                entry={entry}
                block={entry.block}
                sourceBlock={entry.sourceBlock}
                expandedTests={expandedTests}
                reviewIssueIndex={reviewIssueIndex}
                isCollapsed={collapsedBlocks[entry.key] ?? false}
                onBlockCollapseChange={onBlockCollapseChange}
                onAddTest={onAddTest}
                onAddStandardTests={onAddStandardTests}
                onDuplicateBlockStructure={onDuplicateBlockStructure}
                onTestExpansionChange={onTestExpansionChange}
                onTestTitleChange={onTestTitleChange}
                onDuplicateTest={onDuplicateTest}
                onTestRemove={onTestRemove}
                onTestMove={onTestMove}
                onResultChange={onResultChange}
              />
            ))}
          </Stack>
          ) : null}
        </Collapse>
      </Stack>
    </Paper>
  );
}, arePermissionBlockGroupPropsEqual);

const PermissionBlockEditor = memo(function PermissionBlockEditor({
  blockKey,
  entry,
  block,
  sourceBlock,
  expandedTests,
  reviewIssueIndex,
  isCollapsed,
  onBlockCollapseChange,
  onAddTest,
  onAddStandardTests,
  onDuplicateBlockStructure,
  onTestExpansionChange,
  onTestTitleChange,
  onDuplicateTest,
  onTestRemove,
  onTestMove,
  onResultChange,
}: PermissionBlockEditorProps) {
  const isExpanded = !isCollapsed;
  const panelId = `micro-tests-${toDomId(blockKey)}`;
  const blockReviewIssues = getReviewIssuesByKey(reviewIssueIndex.byBlockKey, blockKey);
  const blockReview = summarizeReviewIssues(blockReviewIssues);
  const summaryItems = buildPermissionBlockSummaryItems(sourceBlock);

  return (
    <Paper
      id={`permission-block-${toDomId(blockKey)}`}
      withBorder
      p="md"
      className="permissionBlock"
    >
      <Stack gap="sm">
        <Group justify="space-between" align="center" gap="md" className="permissionBlockHeader">
          <Group gap="xs" align="center" wrap="nowrap" className="permissionBlockTitle">
            <Tooltip label={isExpanded ? "Recolher micro" : "Abrir micro"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onBlockCollapseChange(blockKey, isExpanded)}
                aria-label={isExpanded ? "Recolher micro" : "Abrir micro"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${
                    isExpanded ? "testToggleIcon--open" : ""
                  }`}
                />
              </ActionIcon>
            </Tooltip>
            <div className="summaryHeaderCopy">
              <Group gap="xs" wrap="wrap">
                <Badge variant="outline" color="gray">
                  Micro
                </Badge>
                <Text fw={800}>{formatPermission(entry.micro)}</Text>
              </Group>
              <Text c="dimmed" size="xs">
                {formatPermission(entry.macro)}
              </Text>
              <SummaryChips items={summaryItems} review={blockReview} />
            </div>
          </Group>

          <PermissionBlockActionsMenu
            canCopy={sourceBlock.tests.length > 0}
            onAddStandardTests={() => onAddStandardTests(blockKey)}
            onDuplicateBlockStructure={() => onDuplicateBlockStructure(blockKey)}
            onAddTest={() => onAddTest(blockKey)}
          />
        </Group>

        <Collapse in={isExpanded}>
          {isExpanded ? (
          <Stack gap="sm" id={panelId}>
            {block.tests.map((test, visibleIndex) => {
              const selfReferenceKey = createTestReferenceKey(entry.key, test.id);
              const sourceIndex = sourceBlock.tests.findIndex(
                (sourceTest) => sourceTest.id === test.id,
              );
              const testIndex = sourceIndex >= 0 ? sourceIndex : visibleIndex;

              return (
                <BlockTestEditor
                  key={test.id}
                  blockKey={blockKey}
                  index={testIndex}
                  test={test}
                  selfReferenceKey={selfReferenceKey}
                  isExpanded={expandedTests[selfReferenceKey] ?? false}
                  reviewIssueIndex={reviewIssueIndex}
                  canMoveUp={testIndex > 0}
                  canMoveDown={testIndex < sourceBlock.tests.length - 1}
                  onTestExpansionChange={onTestExpansionChange}
                  onTestTitleChange={onTestTitleChange}
                  onTestMove={onTestMove}
                  onDuplicateTest={onDuplicateTest}
                  onTestRemove={onTestRemove}
                  onResultChange={onResultChange}
                />
              );
            })}

            {block.tests.length === 0 ? (
              <EmptyState
                message="Nenhum teste adicionado nesta micro."
                actionLabel="Adicionar teste"
                onAction={() => onAddTest(blockKey)}
              />
            ) : null}
          </Stack>
          ) : null}
        </Collapse>
      </Stack>
    </Paper>
  );
}, arePermissionBlockEditorPropsEqual);

function PermissionBlockActionsMenu({
  canCopy,
  onAddStandardTests,
  onDuplicateBlockStructure,
  onAddTest,
}: {
  canCopy: boolean;
  onAddStandardTests: () => void;
  onDuplicateBlockStructure: () => void;
  onAddTest: () => void;
}) {
  return (
    <Menu position="bottom-end" withArrow>
      <Menu.Target>
        <span className="actionsMenuTarget">
          <ActionIcon variant="subtle" aria-label="Mais ações da micro" className="actionsMenu">
            <MoreVertical size={17} />
          </ActionIcon>
        </span>
      </Menu.Target>
      <Menu.Dropdown>
        <Menu.Item leftSection={<ListChecks size={15} />} onClick={onAddStandardTests}>
          Adicionar pacote padrão
        </Menu.Item>
        <Menu.Item
          leftSection={<Copy size={15} />}
          disabled={!canCopy}
          onClick={onDuplicateBlockStructure}
        >
          Copiar para vazios
        </Menu.Item>
        <Menu.Item leftSection={<Plus size={15} />} onClick={onAddTest}>
          Adicionar teste
        </Menu.Item>
      </Menu.Dropdown>
    </Menu>
  );
}

function TeaReviewCategorySection({
  title,
  emptyMessage,
  issues,
  onIssueClick,
}: {
  title: string;
  emptyMessage: string;
  issues: TeaReviewIssue[];
  onIssueClick: (issue: TeaReviewIssue) => void;
}) {
  const dangerCount = issues.filter((issue) => issue.severity === "danger").length;
  const tone = dangerCount > 0 ? "danger" : "warning";
  const countLabel = formatTeaCount(issues.length, "pendência", "pendências");

  if (issues.length === 0) {
    return (
      <div className="reviewIssueGroup reviewIssueGroup--ok">
        <Group justify="space-between" align="center" gap="sm" mb={6}>
          <Group gap="xs">
            <CheckCircle2 size={17} />
            <Text fw={700}>{title}</Text>
          </Group>
          <Badge variant="outline" color="green">
            0 pendências
          </Badge>
        </Group>
        <Text c="dimmed" size="sm">
          {emptyMessage}
        </Text>
      </div>
    );
  }

  return (
    <ReviewIssueGroup
      title={title}
      tone={tone}
      issues={issues}
      countLabel={countLabel}
      actionLabel="Corrigir agora"
      showSeverity
      onIssueClick={onIssueClick}
    />
  );
}

function BlockTestActionsMenu({
  canMoveUp,
  canMoveDown,
  onDuplicate,
  onMove,
  onRemove,
}: {
  canMoveUp: boolean;
  canMoveDown: boolean;
  onDuplicate: () => void;
  onMove: (direction: MoveDirection) => void;
  onRemove: () => void;
}) {
  return (
    <Menu position="bottom-end" withArrow>
      <Menu.Target>
        <span className="actionsMenuTarget">
          <ActionIcon variant="subtle" aria-label="Mais ações do teste" className="actionsMenu">
            <MoreVertical size={17} />
          </ActionIcon>
        </span>
      </Menu.Target>
      <Menu.Dropdown>
        <Menu.Item leftSection={<Copy size={15} />} onClick={onDuplicate}>
          Duplicar teste
        </Menu.Item>
        <Menu.Item
          leftSection={<ArrowUp size={15} />}
          disabled={!canMoveUp}
          onClick={() => onMove("up")}
        >
          Mover teste para cima
        </Menu.Item>
        <Menu.Item
          leftSection={<ArrowDown size={15} />}
          disabled={!canMoveDown}
          onClick={() => onMove("down")}
        >
          Mover teste para baixo
        </Menu.Item>
        <Menu.Divider />
        <Menu.Item color="red" leftSection={<Trash2 size={15} />} onClick={onRemove}>
          Remover teste
        </Menu.Item>
      </Menu.Dropdown>
    </Menu>
  );
}

const BlockTestEditor = memo(function BlockTestEditor({
  blockKey,
  index,
  test,
  selfReferenceKey,
  isExpanded,
  reviewIssueIndex,
  canMoveUp,
  canMoveDown,
  onTestExpansionChange,
  onTestTitleChange,
  onTestMove,
  onDuplicateTest,
  onTestRemove,
  onResultChange,
}: BlockTestEditorProps) {
  const testPanelId = `test-details-${toDomId(selfReferenceKey)}`;
  const testReviewIssues = getReviewIssuesByKey(
    reviewIssueIndex.byTestReferenceKey,
    selfReferenceKey,
  );
  const testReview = summarizeReviewIssues(testReviewIssues);
  const summaryItems = buildTestSummaryItems(test);
  const correction = getTestCorrection(test);
  const displayTitle = test.title.trim() || `Teste ${index + 1} sem nome`;
  const commitTitle = useCallback(
    (value: string) => onTestTitleChange(blockKey, test.id, value),
    [blockKey, onTestTitleChange, test.id],
  );
  const title = useBufferedText(test.title, commitTitle);

  function toggleCheck(key: CheckKey): void {
    onResultChange(
      blockKey,
      test.id,
      (current) => ({
        ...current,
        checks: {
          ...current.checks,
          [key]: !current.checks[key],
        },
      }),
    );
  }

  return (
    <Paper
      id={`test-card-${toDomId(selfReferenceKey)}`}
      withBorder
      p="md"
      className={`testCard ${isExpanded ? "testCard--expanded" : "testCard--collapsed"}`}
    >
      <Stack gap="sm">
        <Group justify="space-between" align="center" gap="md" className="testHeader">
          <Group gap="xs" wrap="nowrap" className="testTitle">
            <Tooltip label={isExpanded ? "Recolher teste" : "Abrir teste"}>
              <ActionIcon
                variant="subtle"
                onClick={() => onTestExpansionChange(selfReferenceKey, !isExpanded)}
                aria-label={isExpanded ? "Recolher teste" : "Abrir teste"}
                aria-expanded={isExpanded}
                aria-controls={testPanelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${
                    isExpanded ? "testToggleIcon--open" : ""
                  }`}
                />
              </ActionIcon>
            </Tooltip>
            <Badge color="gray" variant="outline" w={34} h={34}>
              {index + 1}
            </Badge>
            <div className="summaryHeaderCopy">
              <Text fw={800} size="sm">
                {displayTitle}
              </Text>
              <CheckStatusChips result={test.result} />
              {correction.corrected ? (
                <Badge size="xs" color="green" variant="light" className="correctionStatusBadge">
                  Corrigido
                </Badge>
              ) : null}
              <SummaryChips items={summaryItems} review={testReview} />
            </div>
          </Group>
          <BlockTestActionsMenu
            canMoveUp={canMoveUp}
            canMoveDown={canMoveDown}
            onDuplicate={() => onDuplicateTest(blockKey, test.id)}
            onMove={(direction) => onTestMove(blockKey, index, direction === "up" ? -1 : 1)}
            onRemove={() => onTestRemove(blockKey, test.id)}
          />
        </Group>

        {isExpanded ? (
        <Collapse in transitionDuration={0}>
          <div id={testPanelId} className="testBody">
            <Stack gap="md">
              <TextInput
                label="Nome do teste"
                value={title.value}
                placeholder="Criação, edição, consulta..."
                onBlur={title.commit}
                onChange={(event) => {
                  const value = event.currentTarget.value;
                  title.setValue(value);
                }}
              />
              <Paper withBorder p="sm" className="quickChecksPanel">
                <Stack gap="xs">
                  <Text fw={750} size="sm">
                    Status rapido
                  </Text>
                  <Group gap={6} className="quickChecks">
                    {checkOrder.map((key) => (
                      <Tooltip key={key} label={checkLabels[key]}>
                        <button
                          type="button"
                          aria-pressed={test.result.checks[key]}
                          aria-label={`Alternar ${checkLabels[key]}`}
                          className={[
                            "quickCheck",
                            quickCheckToneClassNames[key],
                            test.result.checks[key] ? "quickCheck--active" : "",
                          ]
                            .filter(Boolean)
                            .join(" ")}
                          onClick={() => toggleCheck(key)}
                        >
                          {renderCheckIcon(key, 14)}
                          <span>{quickCheckLabels[key]}</span>
                        </button>
                      </Tooltip>
                    ))}
                  </Group>
                </Stack>
              </Paper>
              <TestResultEditor
                result={test.result}
                onChange={(updater) => onResultChange(blockKey, test.id, updater)}
              />
              <CorrectionReadonlyPanel correction={correction} />
            </Stack>
          </div>
        </Collapse>
        ) : null}
      </Stack>
    </Paper>
  );
});

function CorrectionPanel({
  groups,
  onChangeGroup,
}: {
  groups: CorrectionGroup[];
  onChangeGroup: (
    groupKey: string,
    updater: (correction: TestCorrection) => TestCorrection,
  ) => void;
}) {
  const correctedCount = groups.filter((group) => group.correction.corrected).length;

  return (
    <Section
      title="Para corrigir"
      sectionId="ot-section-corrections"
      tone="blocks"
      action={
        <Badge variant="light" color={groups.length === correctedCount ? "green" : "yellow"}>
          {correctedCount}/{groups.length} corrigidos
        </Badge>
      }
    >
      {groups.length > 0 ? (
        <Stack gap="md">
          {groups.map((group) => (
            <CorrectionGroupCard
              key={group.key}
              group={group}
              onChange={(updater) => onChangeGroup(group.key, updater)}
            />
          ))}
        </Stack>
      ) : (
        <Paper withBorder p="md" ta="center" className="softEmpty">
          <Text c="dimmed">Nenhum teste marcado como Novo.</Text>
        </Paper>
      )}
    </Section>
  );
}

function CorrectionGroupCard({
  group,
  onChange,
}: {
  group: CorrectionGroup;
  onChange: (updater: (correction: TestCorrection) => TestCorrection) => void;
}) {
  const [isExpanded, setIsExpanded] = useState(false);
  const panelId = `correction-details-${toDomId(group.key)}`;
  const legacyImages = group.occurrences.flatMap((occurrence) => occurrence.test.result.legacyImages);
  const newImages = group.occurrences.flatMap((occurrence) => occurrence.test.result.newImages);
  const observations = group.occurrences
    .map((occurrence) => occurrence.test.result.observations.trim())
    .filter(Boolean);
  const hasBeforeImage = group.correction.beforeImages.length > 0;
  const hasAfterImage = group.correction.afterImages.length > 0;
  const canToggleCorrected = hasBeforeImage && hasAfterImage;
  const correctionSummaryItems = [
    `Antes ${group.correction.beforeImages.length}`,
    `Depois ${group.correction.afterImages.length}`,
    group.correction.hotfixTag.trim() ? group.correction.hotfixTag.trim() : "Sem hotfix",
    group.correction.correctedBy.trim()
      ? `Por ${group.correction.correctedBy.trim()}`
      : "Sem responsavel",
  ];

  function updateImages(
    field: "beforeImages" | "afterImages",
    updater: (images: EvidenceImage[]) => EvidenceImage[],
  ): void {
    onChange((current) => ({
      ...current,
      [field]: updater(current[field]),
    }));
  }

  return (
    <Paper
      withBorder
      p="sm"
      className={`correctionCard ${
        group.correction.corrected ? "correctionCard--done" : ""
      } ${isExpanded ? "correctionCard--expanded" : "correctionCard--collapsed"}`}
    >
      <Stack gap="sm">
        <Group justify="space-between" align="flex-start" gap="md" className="correctionHeader">
          <Group gap="xs" wrap="nowrap" className="correctionTitle">
            <Tooltip label={isExpanded ? "Recolher item" : "Abrir item"}>
              <ActionIcon
                variant="subtle"
                onClick={() => setIsExpanded((current) => !current)}
                aria-label={isExpanded ? "Recolher item para corrigir" : "Abrir item para corrigir"}
                aria-expanded={isExpanded}
                aria-controls={panelId}
              >
                <ChevronDown
                  size={18}
                  className={`testToggleIcon ${
                    isExpanded ? "testToggleIcon--open" : ""
                  }`}
                />
              </ActionIcon>
            </Tooltip>
            <div className="summaryHeaderCopy">
              <Group gap="xs" align="center">
                <Text fw={850} size="sm">
                  {group.title}
                </Text>
                {group.correction.corrected ? (
                  <Badge color="green" variant="light">
                    Corrigido
                  </Badge>
                ) : (
                  <Badge color="yellow" variant="light">
                    Pendente
                  </Badge>
                )}
              </Group>
              <Text c="dimmed" size="xs">
                {formatOtCount(group.occurrences.length, "ocorrencia", "ocorrencias")}{" "}
                {group.occurrences.length === 1 ? "agrupada" : "agrupadas"}
              </Text>
              <SummaryChips items={correctionSummaryItems} />
            </div>
          </Group>
        </Group>

        <Collapse in={isExpanded} transitionDuration={0}>
          <div id={panelId} className="correctionBody">
            <Stack gap="sm">
              <div className="correctionSourceGrid">
                <CorrectionReadonlyField title="Observacoes">
                  {observations.length > 0 ? (
                    <Stack gap={4}>
                      {observations.map((observation, index) => (
                        <Text key={`${group.key}-observation-${index}`} size="sm">
                          {observations.length > 1 ? `${index + 1}. ` : ""}
                          {observation}
                        </Text>
                      ))}
                    </Stack>
                  ) : (
                    <Text c="dimmed" size="sm">
                      Sem observacoes.
                    </Text>
                  )}
                </CorrectionReadonlyField>
                <CorrectionReadonlyField title="Legado">
                  <ReadonlyImageStrip images={legacyImages} emptyLabel="Sem imagens do legado." />
                </CorrectionReadonlyField>
                <CorrectionReadonlyField title="Novo">
                  <ReadonlyImageStrip images={newImages} emptyLabel="Sem imagens do novo." />
                </CorrectionReadonlyField>
              </div>

              <Paper withBorder p="sm" className="correctionTodoPanel">
                <Stack gap="sm">
                  <div className="correctionTodoFields">
                    <BufferedTextInput
                      label="Tag da hotfix"
                      value={group.correction.hotfixTag}
                      placeholder="hotfix 1.2.2"
                      onCommit={(value) =>
                        onChange((current) => ({ ...current, hotfixTag: value }))
                      }
                    />
                    <BufferedTextInput
                      label="Corrigido por"
                      value={group.correction.correctedBy}
                      placeholder="Nome de quem corrigiu"
                      onCommit={(value) =>
                        onChange((current) => ({ ...current, correctedBy: value }))
                      }
                    />
                    <Select
                      label="Nuvem"
                      data={cloudStageOptions}
                      value={group.correction.cloudStage}
                      allowDeselect={false}
                      onChange={(value) =>
                        onChange((current) => ({
                          ...current,
                          cloudStage: parseCloudStage(value),
                        }))
                      }
                    />
                  </div>

                  <div className="evidenceGrid">
                    <EvidenceUploader
                      title="Antes (com erro)"
                      tone="legacy"
                      images={group.correction.beforeImages}
                      onChange={(updater) => updateImages("beforeImages", updater)}
                    />
                    <EvidenceUploader
                      title="Depois (corrigido)"
                      tone="new"
                      images={group.correction.afterImages}
                      onChange={(updater) => updateImages("afterImages", updater)}
                    />
                  </div>

                  <Tooltip
                    label={
                      canToggleCorrected
                        ? group.correction.corrected
                          ? "Voltar item para pendente"
                          : "Marcar item como corrigido"
                        : "Adicione prints de antes e depois para marcar como corrigido"
                    }
                  >
                    <Button
                      size="md"
                      fullWidth
                      color={group.correction.corrected ? "green" : "blue"}
                      variant={group.correction.corrected ? "light" : "filled"}
                      leftSection={<CheckCircle2 size={18} />}
                      className="correctionDoneButton"
                      disabled={!canToggleCorrected}
                      onClick={() =>
                        onChange((current) => ({
                          ...current,
                          corrected: !group.correction.corrected,
                        }))
                      }
                    >
                      {group.correction.corrected ? "Corrigido" : "Marcar como corrigido"}
                    </Button>
                  </Tooltip>
                </Stack>
              </Paper>
            </Stack>
          </div>
        </Collapse>
      </Stack>
    </Paper>
  );
}

function CorrectionReadonlyPanel({ correction }: { correction: TestCorrection }) {
  if (!hasCorrectionDetails(correction)) {
    return null;
  }

  return (
    <Paper withBorder p="sm" className="correctionReadonlyPanel">
      <Stack gap="sm">
        <Group justify="space-between" gap="sm">
          <Text fw={750} size="sm">
            Correcao
          </Text>
          {correction.corrected ? (
            <Badge color="green" variant="light">
              Corrigido
            </Badge>
          ) : null}
        </Group>
        <div className="correctionReadonlyMeta">
          <CorrectionReadonlyValue label="Hotfix" value={correction.hotfixTag || "Nao informado"} />
          <CorrectionReadonlyValue
            label="Corrigido por"
            value={correction.correctedBy || "Nao informado"}
          />
          <CorrectionReadonlyValue label="Nuvem" value={formatCloudStage(correction.cloudStage)} />
        </div>
        <div className="correctionReadonlyImages">
          <CorrectionReadonlyField title="Antes (com erro)">
            <ReadonlyImageStrip images={correction.beforeImages} emptyLabel="Sem print antes." />
          </CorrectionReadonlyField>
          <CorrectionReadonlyField title="Depois (corrigido)">
            <ReadonlyImageStrip images={correction.afterImages} emptyLabel="Sem print depois." />
          </CorrectionReadonlyField>
        </div>
      </Stack>
    </Paper>
  );
}

function CorrectionReadonlyField({
  title,
  children,
}: {
  title: string;
  children: ReactNode;
}) {
  return (
    <Paper withBorder p="sm" className="correctionReadonlyField">
      <Stack gap="xs">
        <Text fw={750} size="xs" tt="uppercase" c="dimmed">
          {title}
        </Text>
        {children}
      </Stack>
    </Paper>
  );
}

function CorrectionReadonlyValue({ label, value }: { label: string; value: string }) {
  return (
    <div className="correctionReadonlyValue">
      <Text c="dimmed" size="xs">
        {label}
      </Text>
      <Text fw={750} size="sm">
        {value}
      </Text>
    </div>
  );
}

function ReadonlyImageStrip({
  images,
  emptyLabel,
}: {
  images: EvidenceImage[];
  emptyLabel: string;
}) {
  if (images.length === 0) {
    return (
      <Text c="dimmed" size="sm">
        {emptyLabel}
      </Text>
    );
  }

  return (
    <div className="readonlyImageStrip">
      {images.map((image) =>
        image.dataUrl ? (
          <img
            key={image.id}
            className="imagePreview"
            src={image.dataUrl}
            alt={image.label || image.name || "Imagem anexada"}
            title={image.label || image.name}
          />
        ) : (
          <div key={image.id} className="imagePreview imagePreview--missing">
            Sem preview
          </div>
        ),
      )}
    </div>
  );
}

const TestResultEditor = memo(function TestResultEditor({
  result,
  onChange,
}: {
  result: TestResult;
  onChange: (updater: (result: TestResult) => TestResult) => void;
}) {
  const commitObservations = useCallback(
    (value: string) => onChange((current) => ({ ...current, observations: value })),
    [onChange],
  );
  const observations = useBufferedText(result.observations, commitObservations);
  const needsProblemObservation = hasProblemStatus(result) && !observations.value.trim();

  function updateImages(
    field: "legacyImages" | "newImages",
    updater: (images: EvidenceImage[]) => EvidenceImage[],
  ): void {
    onChange((current) => ({
      ...current,
      [field]: updater(current[field]),
    }));
  }

  return (
    <Stack gap="md">
      <Textarea
        label="Observações"
        error={
          needsProblemObservation
            ? "Obrigatoria quando o status indica problema ou erro."
            : undefined
        }
        minRows={4}
        styles={{ input: { resize: "vertical" } }}
        value={observations.value}
        onBlur={observations.commit}
        onChange={(event) => {
          const value = event.currentTarget.value;
          observations.setValue(value);
        }}
      />

      <div className="evidenceGrid">
        <EvidenceUploader
          title="Legado"
          tone="legacy"
          images={result.legacyImages}
          onChange={(updater) => updateImages("legacyImages", updater)}
        />
        <EvidenceUploader
          title="Novo"
          tone="new"
          images={result.newImages}
          onChange={(updater) => updateImages("newImages", updater)}
        />
      </div>
    </Stack>
  );
});

const EvidenceUploader = memo(function EvidenceUploader({
  title,
  tone,
  images,
  onChange,
}: {
  title: string;
  tone: "legacy" | "new";
  images: EvidenceImage[];
  onChange: (updater: (images: EvidenceImage[]) => EvidenceImage[]) => void;
}) {
  const confirmAction = useConfirmAction();
  const descriptionId = useId();
  const [isProcessingFiles, setIsProcessingFiles] = useState(false);

  async function handlePaste(event: ClipboardEvent<HTMLDivElement>): Promise<void> {
    const files = getPastedImageFiles(event.clipboardData);

    if (!files.length) {
      return;
    }

    event.preventDefault();
    await addFiles(files);
  }

  async function handleDrop(event: DragEvent<HTMLDivElement>): Promise<void> {
    event.preventDefault();
    await addFiles(Array.from(event.dataTransfer.files));
  }

  async function addFiles(files: File[] | File | null): Promise<void> {
    const fileList = Array.isArray(files) ? files : files ? [files] : [];
    const imageFiles = fileList.filter(isImageFile);

    if (!imageFiles.length || isProcessingFiles) {
      return;
    }

    setIsProcessingFiles(true);

    try {
      const evidence = await Promise.all(
        imageFiles.map(async (file) => {
          const optimized = await optimizeImageFile(file);
          const id = createId();

          return {
            id,
            label: "",
            name: file.name,
            dataUrl: optimized.dataUrl,
            width: optimized.width,
            height: optimized.height,
            originalBytes: optimized.originalBytes,
            savedBytes: optimized.savedBytes,
            optimized: optimized.optimized,
          };
        }),
      );

      try {
        await saveEvidenceImageDataBatch(
          evidence.map((image) => ({ id: image.id, dataUrl: image.dataUrl })),
        );
      } catch {
        window.alert("Nao foi possivel salvar a imagem no rascunho do navegador.");
      }

      onChange((current) => [...current, ...evidence]);
    } finally {
      setIsProcessingFiles(false);
    }
  }

  function updateImageLabel(imageId: string, label: string): void {
    onChange((current) =>
      current.map((image) => (image.id === imageId ? { ...image, label } : image)),
    );
  }

  function removeImage(image: EvidenceImage): void {
    function removeCurrentImage(): void {
      void deleteEvidenceImageData(image.id);
      onChange((current) => current.filter((candidate) => candidate.id !== image.id));
    }

    if (!hasMeaningfulEvidenceImageContent(image)) {
      removeCurrentImage();
      return;
    }

    confirmAction(
      {
        title: "Remover imagem?",
        description: `A imagem ${formatConfirmationSubject(image.label || image.name, "selecionada")} será removida deste rascunho.`,
        confirmLabel: "Remover imagem",
      },
      removeCurrentImage,
    );
  }

  return (
    <Paper
      withBorder
      p="sm"
      className={`evidencePanel evidencePanel--${tone}`}
      onPaste={(event) => {
        void handlePaste(event);
      }}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        void handleDrop(event);
      }}
      tabIndex={0}
      role="group"
      aria-busy={isProcessingFiles}
      aria-label={`${title}: cole ou arraste uma imagem`}
      aria-describedby={descriptionId}
    >
      <Stack gap="sm">
        <Group justify="space-between" className="evidenceHeader">
          <div className="evidenceTitle">
            <Group gap="xs" wrap="nowrap">
              <Text fw={700}>{title}</Text>
              <Tooltip label="Aceita imagem colada ou arrastada">
                <ClipboardPaste size={16} className="pasteIndicator" aria-hidden="true" />
              </Tooltip>
            </Group>
            <Text id={descriptionId} c="dimmed" size="xs">
              Cole, arraste ou selecione arquivos.
            </Text>
          </div>
          <FileButton
            onChange={(files) => {
              void addFiles(files);
            }}
            accept="image/*"
            multiple
          >
            {(props) => (
              <Button
                {...props}
                variant="light"
                leftSection={<ImagePlus size={17} />}
                loading={isProcessingFiles}
                disabled={isProcessingFiles}
              >
                Adicionar imagem
              </Button>
            )}
          </FileButton>
        </Group>

        {isProcessingFiles ? (
          <LoadingFeedback
            variant="inline"
            label="Processando imagens..."
            detail="Otimizando e salvando no rascunho."
          />
        ) : null}

        {images.length > 0 ? (
          <Stack gap="xs">
            {images.map((image) => (
              <Paper withBorder p="xs" key={image.id} className="evidenceItem">
                <Group align="center" wrap="nowrap" className="evidenceImageRow">
                  {image.dataUrl ? (
                    <img
                      className="imagePreview"
                      src={image.dataUrl}
                      alt={image.label || image.name || "Imagem anexada"}
                    />
                  ) : (
                    <div className="imagePreview imagePreview--missing">Sem preview</div>
                  )}
                  <Stack gap={4} className="evidenceImageFields">
                    <BufferedTextInput
                      label="Legenda da imagem"
                      value={image.label}
                      onCommit={(value) => updateImageLabel(image.id, value)}
                    />
                    <Text size="xs" c="dimmed" truncate>
                      {image.name}
                      {image.savedBytes ? ` · ${formatBytes(image.savedBytes)}` : ""}
                      {image.optimized ? " · otimizada" : ""}
                    </Text>
                  </Stack>
                  <Tooltip label="Remover imagem">
                    <ActionIcon
                      variant="subtle"
                      color="red"
                      onClick={() => removeImage(image)}
                      aria-label="Remover imagem"
                    >
                      <Trash2 size={17} />
                    </ActionIcon>
                  </Tooltip>
                </Group>
              </Paper>
            ))}
          </Stack>
        ) : (
          <Text c="dimmed" ta="center" py="md" className="evidenceEmpty">
            Sem imagens
          </Text>
        )}
      </Stack>
    </Paper>
  );
});

function EmptyState({
  actionLabel,
  onAction,
  message = "Nenhum item adicionado.",
}: {
  actionLabel: string;
  onAction: () => void;
  message?: string;
}) {
  return (
    <Paper withBorder p="md" ta="center" className="softEmpty">
      <Stack gap="xs" align="center">
        <Text c="dimmed">{message}</Text>
        <Button variant="light" size="xs" leftSection={<Plus size={15} />} onClick={onAction}>
          {actionLabel}
        </Button>
      </Stack>
    </Paper>
  );
}

function ConfirmationModal({
  confirmation,
  isConfirming,
  onCancel,
  onConfirm,
}: {
  confirmation: PendingConfirmation | null;
  isConfirming: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <Modal
      opened={confirmation !== null}
      onClose={isConfirming ? () => undefined : onCancel}
      title={confirmation?.title ?? "Confirmar ação"}
      centered
      closeOnClickOutside={!isConfirming}
      closeOnEscape={!isConfirming}
    >
      {confirmation ? (
        <Stack gap="md">
          <Group gap="xs" align="flex-start" className="confirmationMessage">
            <AlertCircle size={18} aria-hidden="true" />
            <Text size="sm">{confirmation.description}</Text>
          </Group>
          <Group justify="flex-end" gap="xs">
            <Button variant="light" color="gray" onClick={onCancel} disabled={isConfirming}>
              Cancelar
            </Button>
            <Button
              color={confirmation.tone === "danger" ? "red" : undefined}
              leftSection={<Trash2 size={17} />}
              onClick={onConfirm}
              loading={isConfirming}
            >
              {confirmation.confirmLabel}
            </Button>
          </Group>
        </Stack>
      ) : null}
    </Modal>
  );
}

function TeaSubActivityCopyModal({
  documentData,
  request,
  isCopying,
  onTargetActivityChange,
  onClose,
  onConfirm,
}: {
  documentData: TeaDocument;
  request: TeaSubActivityCopyRequest | null;
  isCopying: boolean;
  onTargetActivityChange: (targetActivityId: string | null) => void;
  onClose: () => void;
  onConfirm: () => void;
}) {
  const sourceSubActivity = request
    ? findTeaSubActivity(documentData, request.sourceActivityId, request.subActivityId)
    : undefined;
  const targetOptions = request
    ? buildTeaActivitySelectOptions(documentData, request.sourceActivityId)
    : [];
  const canConfirm = Boolean(
    request?.targetActivityId &&
      sourceSubActivity &&
      targetOptions.some((option) => option.value === request.targetActivityId),
  );

  return (
    <Modal
      opened={request !== null}
      onClose={isCopying ? () => undefined : onClose}
      title="Copiar subtópico"
      centered
      closeOnClickOutside={!isCopying}
      closeOnEscape={!isCopying}
    >
      {request ? (
        <Stack gap="md">
          <div>
            <Text fw={800}>
              {sourceSubActivity?.title.trim() || "Subtópico sem título"}
            </Text>
            <Text size="sm" c="dimmed">
              Selecione a atividade que receberá uma cópia completa deste subtópico.
            </Text>
          </div>

          <Select
            label="Atividade destino"
            aria-label="Atividade destino"
            placeholder={
              targetOptions.length > 0
                ? "Selecione uma atividade"
                : "Nenhuma outra atividade disponível"
            }
            data={targetOptions}
            value={request.targetActivityId}
            onChange={onTargetActivityChange}
            disabled={isCopying || targetOptions.length === 0}
            allowDeselect={false}
            searchable
          />

          {targetOptions.length === 0 ? (
            <Paper withBorder p="sm" className="softEmpty">
              <Text size="sm" c="dimmed">
                Crie outra atividade antes de copiar este subtópico.
              </Text>
            </Paper>
          ) : null}

          <Group justify="flex-end" gap="xs">
            <Button variant="light" color="gray" onClick={onClose} disabled={isCopying}>
              Cancelar
            </Button>
            <Button
              leftSection={<Copy size={17} />}
              disabled={!canConfirm}
              loading={isCopying}
              onClick={onConfirm}
            >
              Copiar subtópico
            </Button>
          </Group>
        </Stack>
      ) : null}
    </Modal>
  );
}

function ImportPreviewModal({
  result,
  isConfirming,
  onClose,
  onConfirm,
}: {
  result: DocxImportResult | null;
  isConfirming: boolean;
  onClose: () => void;
  onConfirm: () => void;
}) {
  return (
    <Modal
      opened={result !== null}
      onClose={isConfirming ? () => undefined : onClose}
      title="Prévia da importação"
      size="lg"
      centered
      closeOnClickOutside={!isConfirming}
      closeOnEscape={!isConfirming}
    >
      {result ? (
        <Stack gap="md">
          <div>
            <Text fw={800}>
              {result.kind === "tea"
                ? result.summary.subject || result.sourceName
                : result.summary.screen || result.sourceName}
            </Text>
            <Text size="sm" c="dimmed">
              {result.sourceName}
            </Text>
          </div>

          <div className="importMetrics">
            {result.kind === "tea" ? (
              <>
                <ReviewMetric label="Atividades" value={result.summary.activities} />
                <ReviewMetric label="Subtópicos" value={result.summary.subActivities} />
                <ReviewMetric label="Blocos" value={result.summary.blocks} />
              </>
            ) : (
              <>
                <ReviewMetric label="Passos" value={result.summary.accessSteps} />
                <ReviewMetric label="Permissões" value={result.summary.selectedPermissions} />
                <ReviewMetric label="Testes" value={result.summary.tests} />
              </>
            )}
            <ReviewMetric label="Imagens" value={result.summary.images} />
          </div>

          {result.warnings.length > 0 ? (
            <div className="importWarnings" role="alert" aria-live="polite">
              <Group gap="xs" mb={6}>
                <AlertCircle size={17} />
                <Text fw={700}>Avisos</Text>
              </Group>
              <Stack gap={4}>
                {result.warnings.map((warning) => (
                  <Text key={warning} size="sm" c="dimmed">
                    {warning}
                  </Text>
                ))}
              </Stack>
            </div>
          ) : null}

          <Group justify="flex-end" gap="xs">
            <Button variant="light" color="gray" onClick={onClose} disabled={isConfirming}>
              Cancelar
            </Button>
            <Button
              leftSection={<FileUp size={17} />}
              loading={isConfirming}
              onClick={onConfirm}
            >
              Substituir rascunho por importação
            </Button>
          </Group>
        </Stack>
      ) : null}
    </Modal>
  );
}

function FaqPanel() {
  return (
    <Stack gap="md">
      <div className="helpIntro">
        <div className="helpIntroIcon" aria-hidden="true">
          <CircleHelp size={22} />
        </div>
        <div>
          <Title order={2} size="h4">
            FAQ do Gerador de OT
          </Title>
          <Text c="dimmed" size="sm">
            Guia rápido para lembrar o que cada ação faz durante o preenchimento da OT.
          </Text>
        </div>
      </div>

      <Section title="Como usar" tone="document">
        <div className="faqSectionList">
          {faqSections.map((section) => (
            <div className="faqSection" key={section.title}>
              <Text fw={800} className="faqSectionTitle">
                {section.title}
              </Text>
              <div className="faqItemList">
                {section.items.map((item) => (
                  <div className="faqItem" key={item.title}>
                    <div>
                      <Text fw={750}>{item.title}</Text>
                      <Text size="sm" c="dimmed">
                        {item.description}
                      </Text>
                    </div>
                    {item.example ? <pre className="faqExample">{item.example}</pre> : null}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </Section>

    </Stack>
  );
}

function ReviewPanel({
  summary,
  onIssueClick,
}: {
  summary: ReviewSummary;
  onIssueClick: (issue: ReviewIssue) => void;
}) {
  const totalIssues = summary.issues.length;
  const dangerIssues = summary.issues.filter((issue) => issue.severity === "danger");
  const warningIssues = summary.issues.filter((issue) => issue.severity === "warning");

  return (
    <Section
      title="Revisão"
      tone="document"
      action={
        <Badge
          color={dangerIssues.length > 0 ? "red" : totalIssues > 0 ? "yellow" : "green"}
          variant="light"
        >
          {totalIssues > 0 ? "Pendências" : "Pronto"}
        </Badge>
      }
    >
      <Stack gap="md">
        <div className="reviewMetrics">
          <ReviewMetric label="Permissões" value={summary.selectedPermissions} />
          <ReviewMetric label="Testes" value={summary.testCount} />
          <ReviewMetric label="Imagens" value={summary.imageCount} />
          <ReviewMetric label="Pendências" value={totalIssues} />
        </div>

        <div className="reviewIssues">
          {dangerIssues.length > 0 ? (
            <ReviewIssueGroup
              title="Pendencias criticas"
              tone="danger"
              issues={dangerIssues}
              onIssueClick={onIssueClick}
            />
          ) : null}
          {warningIssues.length > 0 ? (
            <ReviewIssueGroup
              title="Avisos para revisar"
              tone="warning"
              issues={warningIssues}
              onIssueClick={onIssueClick}
            />
          ) : null}
          {totalIssues === 0 ? (
            <Group gap="xs" className="reviewOk">
              <CheckCircle2 size={18} />
              <Text fw={700}>Nenhuma pendência encontrada.</Text>
            </Group>
          ) : null}
        </div>
      </Stack>
    </Section>
  );
}

function ReviewMetric({ label, value }: { label: string; value: number }) {
  return (
    <div className="reviewMetric">
      <Text size="xs" c="dimmed" fw={700} tt="uppercase">
        {label}
      </Text>
      <Text size="xl" fw={800}>
        {value}
      </Text>
    </div>
  );
}

function ReviewIssueGroup<TIssue extends {
  id: string;
  label: string;
  detail: string;
  severity?: ReviewSeverity;
}>({
  title,
  tone,
  issues,
  countLabel,
  actionLabel,
  showSeverity = false,
  onIssueClick,
}: {
  title: string;
  tone: "warning" | "danger";
  issues: TIssue[];
  countLabel?: string;
  actionLabel?: string;
  showSeverity?: boolean;
  onIssueClick: (issue: TIssue) => void;
}) {
  return (
    <div className={`reviewIssueGroup reviewIssueGroup--${tone}`}>
      <Group justify="space-between" align="center" gap="sm" mb={6}>
        <Group gap="xs">
          <AlertCircle size={17} />
          <Text fw={700}>{title}</Text>
        </Group>
        {countLabel ? (
          <Badge variant="light" color={tone === "danger" ? "red" : "yellow"}>
            {countLabel}
          </Badge>
        ) : null}
      </Group>
      <Stack gap={4}>
        {issues.map((issue) => (
          <button
            key={issue.id}
            type="button"
            className="reviewIssueButton"
            aria-label={
              actionLabel
                ? `${actionLabel}: ${issue.label}. ${issue.detail}`
                : `${issue.label}. ${issue.detail}`
            }
            onClick={() => onIssueClick(issue)}
          >
            <div className="reviewIssueButtonHeader">
              <Text size="sm" fw={700}>
                {issue.label}
              </Text>
              {showSeverity && issue.severity ? (
                <Badge
                  size="xs"
                  variant="light"
                  color={issue.severity === "danger" ? "red" : "yellow"}
                >
                  {getReviewSeverityLabel(issue.severity)}
                </Badge>
              ) : null}
            </div>
            <Text size="xs" c="dimmed">
              {issue.detail}
            </Text>
            {actionLabel ? (
              <span className="reviewIssueCta">{actionLabel}</span>
            ) : null}
          </button>
        ))}
      </Stack>
    </div>
  );
}

function getReviewSeverityLabel(severity: ReviewSeverity): string {
  return severity === "danger" ? "Crítica" : "Aviso";
}

function getOutlineTargetIds(groups: DocumentOutlineGroup[]): string[] {
  return groups.flatMap((group) =>
    group.items.flatMap((item) => (item.targetId ? [item.targetId] : [])),
  );
}

function useActiveOutlineTargetId(
  targetIds: string[],
  scopeKey: string,
): [string | undefined, (targetId?: string) => void] {
  const [activeTargetId, setActiveTargetId] = useState<string | undefined>();
  const targetSignature = targetIds.join("|");

  useEffect(() => {
    let animationFrame: number | undefined;
    const scopedTargetIds = targetSignature ? targetSignature.split("|") : [];

    function getMountedTargets(): HTMLElement[] {
      return scopedTargetIds
        .map((targetId) => window.document.getElementById(targetId))
        .filter((element): element is HTMLElement => Boolean(element));
    }

    function updateActiveTarget(): void {
      const mountedTargets = getMountedTargets();

      if (mountedTargets.length === 0) {
        setActiveTargetId(undefined);
        return;
      }

      const marker = Math.min(180, window.innerHeight * 0.3);
      let activeTarget: HTMLElement | undefined;
      let activeTargetTop = Number.NEGATIVE_INFINITY;
      let closestBelow: HTMLElement | undefined;
      let closestBelowDistance = Number.POSITIVE_INFINITY;
      let firstTargetTop: number | undefined;
      let hasDistinctTargetPositions = false;

      mountedTargets.forEach((target) => {
        const rect = target.getBoundingClientRect();

        if (firstTargetTop === undefined) {
          firstTargetTop = rect.top;
        } else if (rect.top !== firstTargetTop) {
          hasDistinctTargetPositions = true;
        }

        if (rect.top <= marker && rect.top > activeTargetTop) {
          activeTarget = target;
          activeTargetTop = rect.top;
          return;
        }

        const distance = rect.top - marker;

        if (distance < closestBelowDistance) {
          closestBelow = target;
          closestBelowDistance = distance;
        }
      });

      const nextTargetId = (activeTarget ?? closestBelow ?? mountedTargets[0]).id;

      setActiveTargetId((currentTargetId) =>
        currentTargetId &&
        !hasDistinctTargetPositions &&
        mountedTargets.some((target) => target.id === currentTargetId)
          ? currentTargetId
          : nextTargetId,
      );
    }

    function scheduleActiveTargetUpdate(): void {
      if (animationFrame !== undefined) {
        return;
      }

      animationFrame = window.requestAnimationFrame(() => {
        animationFrame = undefined;
        updateActiveTarget();
      });
    }

    updateActiveTarget();
    const timeoutId = window.setTimeout(updateActiveTarget, 80);
    window.addEventListener("scroll", scheduleActiveTargetUpdate, { passive: true });
    window.addEventListener("resize", scheduleActiveTargetUpdate);

    return () => {
      window.clearTimeout(timeoutId);
      window.removeEventListener("scroll", scheduleActiveTargetUpdate);
      window.removeEventListener("resize", scheduleActiveTargetUpdate);

      if (animationFrame !== undefined) {
        window.cancelAnimationFrame(animationFrame);
      }
    };
  }, [scopeKey, targetSignature]);

  return [activeTargetId, setActiveTargetId];
}

function scrollAndFocusReviewTarget(
  targetId: string | undefined,
  block: ScrollLogicalPosition = "center",
): void {
  if (!targetId) {
    return;
  }

  const target = window.document.getElementById(targetId);

  if (!target) {
    return;
  }

  target.scrollIntoView?.({
    behavior: "smooth",
    block,
  });

  if (!target.matches("input, textarea, button, select, a[href], [tabindex]")) {
    target.setAttribute("tabindex", "-1");
  }

  target.focus({ preventScroll: true });
}

function useBufferedText(
  externalValue: string,
  onCommit: (value: string) => void,
  delay = 120,
): BufferedTextState {
  const registerBufferedCommit = useContext(BufferedCommitContext);
  const [value, setValueState] = useState(externalValue);
  const valueRef = useRef(externalValue);
  const lastExternalValueRef = useRef(externalValue);
  const lastCommittedValueRef = useRef(externalValue);
  const commitTimerRef = useRef<number | undefined>();
  const onCommitRef = useRef(onCommit);

  useEffect(() => {
    onCommitRef.current = onCommit;
  }, [onCommit]);

  const clearScheduledCommit = useCallback(() => {
    if (commitTimerRef.current !== undefined) {
      window.clearTimeout(commitTimerRef.current);
      commitTimerRef.current = undefined;
    }
  }, []);

  const commit = useCallback(() => {
    clearScheduledCommit();

    const nextValue = valueRef.current;
    if (nextValue === lastCommittedValueRef.current) {
      return;
    }

    lastCommittedValueRef.current = nextValue;
    onCommitRef.current(nextValue);
  }, [clearScheduledCommit]);

  const setValue = useCallback(
    (nextValue: string) => {
      valueRef.current = nextValue;
      setValueState(nextValue);
      clearScheduledCommit();

      if (nextValue === lastCommittedValueRef.current) {
        return;
      }

      commitTimerRef.current = window.setTimeout(commit, delay);
    },
    [clearScheduledCommit, commit, delay],
  );

  useEffect(() => {
    if (externalValue === lastExternalValueRef.current) {
      return;
    }

    lastExternalValueRef.current = externalValue;

    if (externalValue === lastCommittedValueRef.current && valueRef.current !== externalValue) {
      return;
    }

    lastCommittedValueRef.current = externalValue;
    valueRef.current = externalValue;
    setValueState(externalValue);
    clearScheduledCommit();
  }, [clearScheduledCommit, externalValue]);

  useEffect(() => commit, [commit]);
  useEffect(
    () => registerBufferedCommit?.(commit),
    [commit, registerBufferedCommit],
  );

  return { value, setValue, commit };
}

function arePermissionGroupEditorPropsEqual(
  previous: PermissionGroupEditorProps,
  next: PermissionGroupEditorProps,
): boolean {
  return (
    previous.index === next.index &&
    arePermissionGroupsEqual(previous.macro, next.macro) &&
    previous.onMacroChange === next.onMacroChange &&
    previous.onRemoveMacro === next.onRemoveMacro &&
    previous.onAddMicro === next.onAddMicro &&
    previous.onMicroChange === next.onMicroChange &&
    previous.onRemoveMicro === next.onRemoveMicro
  );
}

function arePermissionGroupsEqual(
  previous: PermissionGroup,
  next: PermissionGroup,
): boolean {
  if (
    !arePermissionItemsEqual(previous, next) ||
    previous.microPermissions.length !== next.microPermissions.length
  ) {
    return false;
  }

  return previous.microPermissions.every((micro, index) =>
    arePermissionItemsEqual(micro, next.microPermissions[index]),
  );
}

function arePermissionItemsEqual(previous: PermissionItem, next: PermissionItem): boolean {
  return (
    previous.id === next.id &&
    previous.code === next.code &&
    previous.label === next.label &&
    previous.selected === next.selected
  );
}

function arePermissionBlockGroupPropsEqual(
  previous: PermissionBlockGroupProps,
  next: PermissionBlockGroupProps,
): boolean {
  if (
    previous.macro !== next.macro ||
    previous.entries !== next.entries ||
    previous.reviewIssueIndex !== next.reviewIssueIndex ||
    previous.isCollapsed !== next.isCollapsed ||
    previous.onMacroCollapseChange !== next.onMacroCollapseChange ||
    previous.onBlockCollapseChange !== next.onBlockCollapseChange ||
    previous.onAddTest !== next.onAddTest ||
    previous.onAddStandardTests !== next.onAddStandardTests ||
    previous.onDuplicateBlockStructure !== next.onDuplicateBlockStructure ||
    previous.onDuplicateTest !== next.onDuplicateTest ||
    previous.onTestExpansionChange !== next.onTestExpansionChange ||
    previous.onTestTitleChange !== next.onTestTitleChange ||
    previous.onTestRemove !== next.onTestRemove ||
    previous.onTestMove !== next.onTestMove ||
    previous.onResultChange !== next.onResultChange
  ) {
    return false;
  }

  for (const entry of next.entries) {
    const previousEntry = previous.entries.find((candidate) => candidate.key === entry.key);

    if (!previousEntry || previousEntry.block !== entry.block) {
      return false;
    }

    if (
      (previous.collapsedBlocks[entry.key] ?? false) !==
      (next.collapsedBlocks[entry.key] ?? false)
    ) {
      return false;
    }

    if (
      previous.expandedTests !== next.expandedTests &&
      !areExpandedStatesEqual(entry.key, entry.block, previous.expandedTests, next.expandedTests)
    ) {
      return false;
    }
  }

  return true;
}

function arePermissionBlockEditorPropsEqual(
  previous: PermissionBlockEditorProps,
  next: PermissionBlockEditorProps,
): boolean {
  if (
    previous.blockKey !== next.blockKey ||
    previous.entry !== next.entry ||
    previous.block !== next.block ||
    previous.sourceBlock !== next.sourceBlock ||
    previous.reviewIssueIndex !== next.reviewIssueIndex ||
    previous.isCollapsed !== next.isCollapsed ||
    previous.onBlockCollapseChange !== next.onBlockCollapseChange ||
    previous.onAddTest !== next.onAddTest ||
    previous.onAddStandardTests !== next.onAddStandardTests ||
    previous.onDuplicateBlockStructure !== next.onDuplicateBlockStructure ||
    previous.onDuplicateTest !== next.onDuplicateTest ||
    previous.onTestExpansionChange !== next.onTestExpansionChange ||
    previous.onTestTitleChange !== next.onTestTitleChange ||
    previous.onTestRemove !== next.onTestRemove ||
    previous.onTestMove !== next.onTestMove ||
    previous.onResultChange !== next.onResultChange
  ) {
    return false;
  }

  return (
    previous.expandedTests === next.expandedTests ||
    areExpandedStatesEqual(next.blockKey, next.block, previous.expandedTests, next.expandedTests)
  );
}

function areExpandedStatesEqual(
  blockKey: string,
  block: PermissionBlock,
  previousExpandedTests: Record<string, boolean>,
  nextExpandedTests: Record<string, boolean>,
): boolean {
  return block.tests.every((test) => {
    const referenceKey = createTestReferenceKey(blockKey, test.id);
    return (previousExpandedTests[referenceKey] ?? false) === (nextExpandedTests[referenceKey] ?? false);
  });
}

function areTeaActivityEditorPropsEqual(
  previous: TeaActivityEditorProps,
  next: TeaActivityEditorProps,
): boolean {
  if (
    previous.index !== next.index ||
    previous.totalActivities !== next.totalActivities ||
    previous.activity !== next.activity ||
    previous.reviewIssues !== next.reviewIssues ||
    previous.isCollapsed !== next.isCollapsed
  ) {
    return false;
  }

  if (next.isCollapsed) {
    return true;
  }

  if (
    (previous.collapsedComposers[next.activity.id] ?? false) !==
    (next.collapsedComposers[next.activity.id] ?? false)
  ) {
    return false;
  }

  if (
    !areTeaBlockCollapseStatesEqual(
      next.activity.blocks,
      previous.collapsedContentBlocks,
      next.collapsedContentBlocks,
    )
  ) {
    return false;
  }

  return next.activity.subActivities.every((subActivity) => {
    if (
      (previous.collapsedSubActivities[subActivity.id] ?? false) !==
      (next.collapsedSubActivities[subActivity.id] ?? false)
    ) {
      return false;
    }

    if (
      (previous.collapsedComposers[subActivity.id] ?? false) !==
      (next.collapsedComposers[subActivity.id] ?? false)
    ) {
      return false;
    }

    return areTeaBlockCollapseStatesEqual(
      subActivity.blocks,
      previous.collapsedContentBlocks,
      next.collapsedContentBlocks,
    );
  });
}

function areTeaSubActivityEditorPropsEqual(
  previous: TeaSubActivityEditorProps,
  next: TeaSubActivityEditorProps,
): boolean {
  if (
    previous.activityIndex !== next.activityIndex ||
    previous.index !== next.index ||
    previous.totalSubActivities !== next.totalSubActivities ||
    previous.subActivity !== next.subActivity ||
    previous.reviewIssues !== next.reviewIssues ||
    previous.isCollapsed !== next.isCollapsed ||
    previous.isComposerCollapsed !== next.isComposerCollapsed
  ) {
    return false;
  }

  return next.isCollapsed ||
    areTeaBlockCollapseStatesEqual(next.subActivity.blocks, previous.collapsedBlocks, next.collapsedBlocks);
}

function areTeaContentComposerPropsEqual(
  previous: TeaContentComposerProps,
  next: TeaContentComposerProps,
): boolean {
  if (
    previous.composerId !== next.composerId ||
    previous.title !== next.title ||
    previous.description !== next.description ||
    previous.emptyMessage !== next.emptyMessage ||
    previous.emptyActionLabel !== next.emptyActionLabel ||
    previous.blocks !== next.blocks ||
    previous.tone !== next.tone ||
    previous.isCollapsed !== next.isCollapsed ||
    previous.reviewIssues !== next.reviewIssues
  ) {
    return false;
  }

  return next.isCollapsed ||
    areTeaBlockCollapseStatesEqual(next.blocks, previous.collapsedBlocks, next.collapsedBlocks);
}

function areTeaContentBlockEditorPropsEqual(
  previous: TeaContentBlockEditorProps,
  next: TeaContentBlockEditorProps,
): boolean {
  return (
    previous.block === next.block &&
    previous.index === next.index &&
    previous.totalBlocks === next.totalBlocks &&
    previous.tone === next.tone &&
    previous.isCollapsed === next.isCollapsed &&
    previous.reviewIssues === next.reviewIssues
  );
}

function areTeaBlockCollapseStatesEqual(
  blocks: TeaContentBlock[],
  previousCollapsedBlocks: Record<string, boolean>,
  nextCollapsedBlocks: Record<string, boolean>,
): boolean {
  return blocks.every(
    (block) =>
      (previousCollapsedBlocks[block.id] ?? false) ===
      (nextCollapsedBlocks[block.id] ?? false),
  );
}

function buildTestBlockFilterCounts(
  entries: PermissionBlockEntry[],
  blocks: Record<string, PermissionBlock>,
): Record<TestBlockFilter, number> {
  const counts: Record<TestBlockFilter, number> = {
    all: 0,
    withoutTests: 0,
    withoutImages: 0,
    withPending: 0,
    withProblem: 0,
  };

  entries.forEach((entry) => {
    const block = blocks[entry.key] ?? emptyPermissionBlock;
    const testCount = block.tests.length;

    counts.all += testCount;

    if (testCount === 0) {
      counts.withoutTests += 1;
    }

    counts.withoutImages += block.tests.filter((test) =>
      testMatchesFilter(test, "withoutImages"),
    ).length;
    counts.withPending += block.tests.filter((test) =>
      testMatchesFilter(test, "withPending"),
    ).length;
    counts.withProblem += block.tests.filter((test) =>
      testMatchesFilter(test, "withProblem"),
    ).length;
  });

  return counts;
}

function testBlockMatchesFilter(block: PermissionBlock, filter: TestBlockFilter): boolean {
  if (filter === "all") {
    return true;
  }

  if (filter === "withoutTests") {
    return block.tests.length === 0;
  }

  return block.tests.some((test) => testMatchesFilter(test, filter));
}

function filterPermissionBlockTests(
  block: PermissionBlock,
  filter: TestBlockFilter,
): PermissionBlock {
  if (filter === "all" || filter === "withoutTests") {
    return block;
  }

  return {
    ...block,
    tests: block.tests.filter((test) => testMatchesFilter(test, filter)),
  };
}

function testMatchesFilter(
  test: PermissionBlockTest,
  filter: Exclude<TestBlockFilter, "all" | "withoutTests">,
): boolean {
  if (filter === "withoutImages") {
    return getTestImageCount(test) === 0;
  }

  if (filter === "withPending") {
    return testHasPendingReview(test);
  }

  return problemCheckKeys.some((key) => test.result.checks[key]);
}

function getTestImageCount(test: PermissionBlockTest): number {
  return test.result.legacyImages.length + test.result.newImages.length;
}

function getSelectedCheckKeys(result: TestResult): CheckKey[] {
  return checkOrder.filter((key) => result.checks[key]);
}

function hasProblemStatus(result: TestResult): boolean {
  return problemCheckKeys.some((key) => result.checks[key]);
}

function testHasPendingReview(test: PermissionBlockTest): boolean {
  const selectedCheckCount = getSelectedCheckKeys(test.result).length;

  return (
    !test.title.trim() ||
    selectedCheckCount === 0 ||
    test.result.legacyImages.length === 0 ||
    test.result.newImages.length === 0 ||
    (hasProblemStatus(test.result) && !test.result.observations.trim())
  );
}

function summarizeReviewIssues(issues: Array<{ severity: ReviewSeverity }>): InlineReviewSummary {
  return issues.reduce<InlineReviewSummary>(
    (summary, issue) => ({
      total: summary.total + 1,
      danger: summary.danger + (issue.severity === "danger" ? 1 : 0),
      warning: summary.warning + (issue.severity === "warning" ? 1 : 0),
    }),
    { total: 0, danger: 0, warning: 0 },
  );
}

function getInlineReviewTone(review: InlineReviewSummary): "gray" | "red" | "yellow" {
  if (review.danger > 0) {
    return "red";
  }

  if (review.warning > 0) {
    return "yellow";
  }

  return "gray";
}

function buildReviewIssueIndex(issues: ReviewIssue[]): ReviewIssueIndex {
  const byBlockKey = new Map<string, ReviewIssue[]>();
  const byTestReferenceKey = new Map<string, ReviewIssue[]>();

  issues.forEach((issue) => {
    if (issue.blockKey) {
      addIssueToMap(byBlockKey, issue.blockKey, issue);
    }

    if (issue.blockKey && issue.testId) {
      addIssueToMap(
        byTestReferenceKey,
        createTestReferenceKey(issue.blockKey, issue.testId),
        issue,
      );
    }
  });

  return { byBlockKey, byTestReferenceKey };
}

function buildTeaReviewIssueIndex(issues: TeaReviewIssue[]): TeaReviewIssueIndex {
  const byTargetId = new Map<string, TeaReviewIssue[]>();
  const byActivityId = new Map<string, TeaReviewIssue[]>();
  const byActivityRootId = new Map<string, TeaReviewIssue[]>();
  const bySubActivityId = new Map<string, TeaReviewIssue[]>();
  const byBlockId = new Map<string, TeaReviewIssue[]>();
  const activityIssues: TeaReviewIssue[] = [];

  issues.forEach((issue) => {
    if (issue.targetId) {
      addIssueToMap(byTargetId, issue.targetId, issue);
    }

    if (issue.activityId) {
      addIssueToMap(byActivityId, issue.activityId, issue);
    }

    if (issue.subActivityId) {
      addIssueToMap(bySubActivityId, issue.subActivityId, issue);
    } else if (issue.activityId) {
      addIssueToMap(byActivityRootId, issue.activityId, issue);
    }

    if (issue.blockId) {
      addIssueToMap(byBlockId, issue.blockId, issue);
    }

    if (issue.tab === "activities" && issue.activityId) {
      activityIssues.push(issue);
    }
  });

  return {
    byTargetId,
    byActivityId,
    byActivityRootId,
    bySubActivityId,
    byBlockId,
    activityIssues,
    hasActivityIssues: activityIssues.length > 0,
  };
}

function splitTeaActivityReviewIssues(issues: TeaReviewIssue[]): {
  rootIssues: TeaReviewIssue[];
  bySubActivityId: Map<string, TeaReviewIssue[]>;
} {
  const rootIssues: TeaReviewIssue[] = [];
  const bySubActivityId = new Map<string, TeaReviewIssue[]>();

  issues.forEach((issue) => {
    if (issue.subActivityId) {
      addIssueToMap(bySubActivityId, issue.subActivityId, issue);
      return;
    }

    rootIssues.push(issue);
  });

  return {
    rootIssues: rootIssues.length > 0 ? rootIssues : emptyTeaReviewIssues,
    bySubActivityId,
  };
}

function groupTeaReviewIssuesByBlock(
  issues: TeaReviewIssue[],
): Map<string, TeaReviewIssue[]> {
  const byBlockId = new Map<string, TeaReviewIssue[]>();

  issues.forEach((issue) => {
    if (issue.blockId) {
      addIssueToMap(byBlockId, issue.blockId, issue);
    }
  });

  return byBlockId;
}

function getReviewIssuesByKey(
  map: Map<string, ReviewIssue[]>,
  key: string,
): ReviewIssue[] {
  return map.get(key) ?? emptyReviewIssues;
}

function getTeaReviewIssuesByKey(
  map: Map<string, TeaReviewIssue[]>,
  key: string,
): TeaReviewIssue[] {
  return map.get(key) ?? emptyTeaReviewIssues;
}

function addIssueToMap<TIssue>(
  map: Map<string, TIssue[]>,
  key: string,
  issue: TIssue,
): void {
  const current = map.get(key);

  if (current) {
    current.push(issue);
    return;
  }

  map.set(key, [issue]);
}

function buildPermissionBlockSummaryItems(block: PermissionBlock): string[] {
  const pendingCount = block.tests.filter(testHasPendingReview).length;
  const problemCount = block.tests.filter((test) => hasProblemStatus(test.result)).length;

  return [
    formatOtCount(block.tests.length, "teste", "testes"),
    formatOtCount(pendingCount, "pendente", "pendentes"),
    `${problemCount} com problema`,
  ];
}

function buildTestSummaryItems(test: PermissionBlockTest): string[] {
  const items = [
    `${getSelectedCheckKeys(test.result).length}/${checkOrder.length} checks`,
    `Legado ${test.result.legacyImages.length}`,
    `Novo ${test.result.newImages.length}`,
  ];

  if (getTestCorrection(test).corrected) {
    items.push("Corrigido");
  }

  return items;
}

function buildCorrectionGroups(
  documentData: OtDocument,
  entries: PermissionBlockEntry[],
): CorrectionGroup[] {
  const groups = new Map<string, CorrectionGroup>();

  entries.forEach((entry) => {
    const block = documentData.permissionBlocks[entry.key] ?? emptyPermissionBlock;

    block.tests.forEach((test, testIndex) => {
      if (!test.result.checks.newIssue) {
        return;
      }

      const key = getCorrectionGroupKey(entry.key, test);
      const referenceKey = createTestReferenceKey(entry.key, test.id);
      const occurrence: CorrectionOccurrence = {
        ...entry,
        test,
        testIndex,
        referenceKey,
      };
      const existing = groups.get(key);

      if (existing) {
        existing.occurrences.push(occurrence);
        return;
      }

      groups.set(key, {
        key,
        title: test.title.trim() || `Teste ${testIndex + 1} sem nome`,
        occurrences: [occurrence],
        correction: getTestCorrection(test),
      });
    });
  });

  return Array.from(groups.values());
}

function getCorrectionGroupKey(blockKey: string, test: PermissionBlockTest): string {
  const normalizedTitle = normalizeTextKey(test.title);
  return normalizedTitle ? `title:${normalizedTitle}` : `single:${blockKey}:${test.id}`;
}

function getTestCorrection(test: PermissionBlockTest): TestCorrection {
  return {
    ...createEmptyTestCorrection(),
    ...test.correction,
    beforeImages: test.correction?.beforeImages ?? [],
    afterImages: test.correction?.afterImages ?? [],
  };
}

function hasCorrectionDetails(correction: TestCorrection): boolean {
  return (
    correction.corrected ||
    Boolean(correction.hotfixTag.trim()) ||
    Boolean(correction.correctedBy.trim()) ||
    correction.cloudStage !== "none" ||
    correction.beforeImages.length > 0 ||
    correction.afterImages.length > 0
  );
}

function parseCloudStage(value: string | null): TestCorrection["cloudStage"] {
  return value === "dev" || value === "homolog" || value === "production"
    ? value
    : "none";
}

function formatCloudStage(value: TestCorrection["cloudStage"]): string {
  return cloudStageOptions.find((option) => option.value === value)?.label ?? "Nao enviado";
}

function buildOtOutlineItems(
  documentData: OtDocument,
  entries: PermissionBlockEntry[],
  reviewSummary: ReviewSummary,
  context: ActiveTab,
): DocumentOutlineGroup[] {
  const issues = reviewSummary.issues;
  const issueCounts = buildOtOutlineIssueCounts(issues);
  const tab = context === "preview" ? "preview" : "tests";

  return [
    {
      id: `ot-${tab}-tests`,
      title: "Testes",
      items: entries.flatMap((entry) => {
        const block = documentData.permissionBlocks[entry.key] ?? emptyPermissionBlock;

        return block.tests.map<DocumentOutlineItem>((test, index) => {
          const referenceKey = createTestReferenceKey(entry.key, test.id);
          const testIssueCount = issueCounts.byTestReferenceKey.get(referenceKey) ?? 0;

          return {
            id: `ot-${tab}-test-${referenceKey}`,
            title: test.title.trim() || `Teste ${index + 1} sem nome`,
            meta: `${getSelectedCheckKeys(test.result).length}/${checkOrder.length} checks - ${formatOtCount(
              getTestImageCount(test),
              "imagem",
              "imagens",
            )}`,
            tab,
            targetId:
              tab === "preview"
                ? `ot-preview-test-${toDomId(referenceKey)}`
                : `test-card-${toDomId(referenceKey)}`,
            blockKey: entry.key,
            testId: test.id,
            status: outlineStatus(testIssueCount),
          };
        });
      }),
    },
  ];

  const documentIssues = issueCounts.byTab.document;
  const permissionIssues = issueCounts.byTab.permissions;
  const testIssues = issueCounts.byTab.tests;
  const nonEmptyStepCount = documentData.accessSteps.filter((step) => step.text.trim()).length;
  const selectedMacroIds = new Set(entries.map((entry) => entry.macro.id));

  return [
    {
      id: "ot-document",
      title: "Documento",
      items: [
        {
          id: "ot-document-main",
          title: documentData.metadata.screen.trim() || "Documento sem tela definida",
          meta: documentData.metadata.responsible.trim() || "Responsável não informado",
          tab: "document",
          targetId: "ot-section-document",
          status: outlineStatus(documentIssues),
        },
        {
          id: "ot-document-steps",
          title: "Passo a passo",
          meta: formatOtCount(nonEmptyStepCount, "passo", "passos"),
          tab: "document",
          targetId: "ot-section-steps",
          status: outlineStatus(issueCounts.byId.get("missing-steps") ?? 0),
        },
      ],
    },
    {
      id: "ot-permissions",
      title: "Permissões",
      items: [
        {
          id: "ot-permissions-main",
          title: "Permissões",
          meta: formatOtCount(entries.length, "micro selecionada", "micros selecionadas"),
          tab: "permissions",
          targetId: "ot-section-permissions",
          status: outlineStatus(permissionIssues),
        },
        ...documentData.permissionGroups.map((macro, index) => ({
          id: `ot-macro-${macro.id}`,
          title: formatPermission(macro) || `Macro ${index + 1} sem nome`,
          meta: `${formatOtCount(macro.microPermissions.length, "micro", "micros")}${
            selectedMacroIds.has(macro.id) ? " • em uso" : ""
          }`,
          tab: "permissions",
          targetId: `permission-macro-${toDomId(macro.id)}`,
          level: 1 as const,
          status: undefined,
        })),
      ],
    },
    {
      id: "ot-tests",
      title: "Testes",
      items: [
        {
          id: "ot-tests-main",
          title: "Testes por permissão",
          meta: formatOtCount(reviewSummary.testCount, "teste", "testes"),
          tab: "tests",
          targetId: "ot-section-tests",
          status: outlineStatus(testIssues),
        },
        ...entries.flatMap((entry) => {
          const block = documentData.permissionBlocks[entry.key] ?? emptyPermissionBlock;
          const blockIssues = issueCounts.byBlockKey.get(entry.key) ?? 0;
          const blockItem: DocumentOutlineItem = {
            id: `ot-block-${entry.key}`,
            title: formatPermission(entry.micro) || "Micro sem nome",
            meta: `${formatPermission(entry.macro) || "Macro sem nome"} • ${formatOtCount(
              block.tests.length,
              "teste",
              "testes",
            )}`,
            tab: "tests",
            targetId: `permission-block-${toDomId(entry.key)}`,
            blockKey: entry.key,
            level: 1,
            status: outlineStatus(blockIssues),
          };
          const testItems = block.tests.map<DocumentOutlineItem>((test, index) => {
            const referenceKey = createTestReferenceKey(entry.key, test.id);
            const testIssueCount = issueCounts.byTestReferenceKey.get(referenceKey) ?? 0;

            return {
              id: `ot-test-${referenceKey}`,
              title: test.title.trim() || `Teste ${index + 1} sem nome`,
              meta: `${getSelectedCheckKeys(test.result).length}/${checkOrder.length} checks • ${formatOtCount(
                getTestImageCount(test),
                "imagem",
                "imagens",
              )}`,
              tab: "tests",
              targetId: `test-card-${toDomId(referenceKey)}`,
              blockKey: entry.key,
              testId: test.id,
              level: 2,
              status: outlineStatus(testIssueCount),
            };
          });

          return [blockItem, ...testItems];
        }),
      ],
    },
    {
      id: "ot-review",
      title: "Revisão",
      items: [
        {
          id: "ot-review-main",
          title: "Revisão",
          meta: formatOtCount(reviewSummary.issues.length, "pendência", "pendências"),
          tab: "review",
          targetId: "ot-section-review",
          status: outlineStatus(reviewSummary.issues.length),
        },
      ],
    },
    {
      id: "ot-preview",
      title: "Prévia",
      items: [
        {
          id: "ot-preview-main",
          title: "Prévia DOCX",
          meta: "Formato de exportação",
          tab: "preview",
          targetId: "ot-section-preview",
        },
      ],
    },
  ];
}

function buildTeaOutlineItems(
  documentData: TeaDocument,
  reviewSummary: TeaReviewSummary,
  context: TeaOutlineContext,
): DocumentOutlineGroup[] {
  const issues = reviewSummary.issues;
  const issueCounts = buildTeaOutlineIssueCounts(issues);
  const tab: TeaTab = context === "preview" ? "preview" : "activities";

  return [
    {
      id: `tea-${context}-activities`,
      title: "Atividades",
      items: documentData.activities.flatMap((activity, activityIndex) => {
        const activityNumber = `2.${activityIndex + 1}`;
        const activityIssueCount = issueCounts.byActivityId.get(activity.id) ?? 0;
        const activityTargetPrefix =
          context === "preview" ? "tea-preview-activity" : "tea-activity";
        const subActivityTargetPrefix =
          context === "preview" ? "tea-preview-subactivity" : "tea-subactivity";
        const activityItem: DocumentOutlineItem = {
          id: `tea-${context}-activity-outline-${activity.id}`,
          title: formatTeaEditorTitle(
            activityNumber,
            activity.title,
            "Atividade sem título",
          ),
          tab,
          targetId: `${activityTargetPrefix}-${toDomId(activity.id)}`,
          activityId: activity.id,
          level: 0,
          status: outlineStatus(activityIssueCount),
        };
        const subActivityItems = activity.subActivities.map<DocumentOutlineItem>(
          (subActivity, subIndex) => {
            const subActivityNumber = `${activityNumber}.${subIndex + 1}`;
            const subActivityIssueCount =
              issueCounts.bySubActivityId.get(subActivity.id) ?? 0;

            return {
              id: `tea-${context}-subactivity-outline-${subActivity.id}`,
              title: formatTeaEditorTitle(
                subActivityNumber,
                subActivity.title,
                "Subtópico sem título",
              ),
              tab,
              targetId: `${subActivityTargetPrefix}-${toDomId(subActivity.id)}`,
              activityId: activity.id,
              subActivityId: subActivity.id,
              level: 1,
              status: outlineStatus(subActivityIssueCount),
            };
          },
        );

        return [activityItem, ...subActivityItems];
      }),
    },
  ];
}

function buildOtOutlineIssueCounts(issues: ReviewIssue[]): {
  byTab: Record<ActiveTab, number>;
  byId: Map<string, number>;
  byBlockKey: Map<string, number>;
  byTestReferenceKey: Map<string, number>;
} {
  const counts = {
    byTab: {
      document: 0,
      permissions: 0,
      tests: 0,
      corrections: 0,
      review: 0,
      preview: 0,
    },
    byId: new Map<string, number>(),
    byBlockKey: new Map<string, number>(),
    byTestReferenceKey: new Map<string, number>(),
  };

  issues.forEach((issue) => {
    counts.byTab[issue.tab] += 1;
    incrementCount(counts.byId, issue.id);

    if (issue.blockKey) {
      incrementCount(counts.byBlockKey, issue.blockKey);

      if (issue.testId) {
        incrementCount(
          counts.byTestReferenceKey,
          createTestReferenceKey(issue.blockKey, issue.testId),
        );
      }
    }
  });

  return counts;
}

function buildTeaOutlineIssueCounts(issues: TeaReviewIssue[]): {
  byActivityId: Map<string, number>;
  bySubActivityId: Map<string, number>;
} {
  const counts = {
    byActivityId: new Map<string, number>(),
    bySubActivityId: new Map<string, number>(),
  };

  issues.forEach((issue) => {
    if (issue.activityId) {
      incrementCount(counts.byActivityId, issue.activityId);
    }

    if (issue.subActivityId) {
      incrementCount(counts.bySubActivityId, issue.subActivityId);
    }
  });

  return counts;
}

function incrementCount(map: Map<string, number>, key: string): void {
  map.set(key, (map.get(key) ?? 0) + 1);
}

function outlineStatus(issueCount: number): OutlineItemStatus | undefined {
  return issueCount > 0 ? "pending" : undefined;
}

function formatOtCount(count: number, singular: string, plural: string): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

function getCheckTone(key: CheckKey): "green" | "yellow" | "red" | "gray" {
  if (key === "sameBehavior") {
    return "green";
  }

  if (key === "errorReport") {
    return "red";
  }

  if (problemCheckKeys.includes(key)) {
    return "yellow";
  }

  return "gray";
}

function renderCheckIcon(key: CheckKey, size = 14): ReactNode {
  if (key === "sameBehavior") {
    return <CheckCircle2 size={size} />;
  }

  if (key === "possibleIssue") {
    return <CircleHelp size={size} />;
  }

  if (key === "errorReport") {
    return <FileText size={size} />;
  }

  return <AlertCircle size={size} />;
}

function buildReviewSummary(
  documentData: OtDocument,
  entries: PermissionBlockEntry[],
): ReviewSummary {
  const summary: ReviewSummary = {
    selectedPermissions: entries.length,
    testCount: 0,
    imageCount: 0,
    issues: [],
  };

  if (!documentData.metadata.screen.trim()) {
    summary.issues.push({
      id: "missing-screen",
      severity: "warning",
      label: "Tela sem nome",
      detail: "Preencha o campo Tela antes de exportar.",
      tab: "document",
    });
  }

  if (!documentData.objective.trim()) {
    summary.issues.push({
      id: "missing-objective",
      severity: "warning",
      label: "Objetivo vazio",
      detail: "Preencha o objetivo do documento.",
      tab: "document",
    });
  }

  if (documentData.accessSteps.filter((step) => step.text.trim()).length === 0) {
    summary.issues.push({
      id: "missing-steps",
      severity: "warning",
      label: "Sem passo a passo",
      detail: "Inclua ao menos uma etapa de acesso.",
      tab: "document",
    });
  }

  if (entries.length === 0) {
    summary.issues.push({
      id: "missing-permissions",
      severity: "danger",
      label: "Sem permissões selecionadas",
      detail: "Selecione macro e micro-permissões para gerar testes.",
      tab: "permissions",
    });
  }

  entries.forEach((entry) => {
    const block = documentData.permissionBlocks[entry.key] ?? createEmptyBlock();

    if (block.tests.length === 0) {
      summary.issues.push({
        id: `missing-tests-${entry.key}`,
        severity: "danger",
        label: "Permissao sem teste",
        detail:
          `${formatPermission(entry.macro)} / ${formatPermission(entry.micro)}: ` +
          "adicione ao menos um teste.",
        tab: "tests",
        targetId: `permission-block-${toDomId(entry.key)}`,
        blockKey: entry.key,
      });
    }

    block.tests.forEach((test, index) => {
      const referenceKey = createTestReferenceKey(entry.key, test.id);
      const targetId = `test-card-${toDomId(referenceKey)}`;
      const selectedCheckKeys = getSelectedCheckKeys(test.result);
      const testLabel = `${formatPermission(entry.macro)} / ${formatPermission(entry.micro)} / ${
        test.title.trim() || `Teste ${index + 1}`
      }`;

      summary.testCount += 1;
      summary.imageCount += test.result.legacyImages.length + test.result.newImages.length;

      if (!test.title.trim()) {
        summary.issues.push({
          id: `unnamed-${referenceKey}`,
          severity: "warning",
          label: "Teste sem nome",
          detail: testLabel,
          tab: "tests",
          targetId,
          blockKey: entry.key,
          testId: test.id,
        });
      }

      if (selectedCheckKeys.length === 0) {
        summary.issues.push({
          id: `missing-status-${referenceKey}`,
          severity: "danger",
          label: "Status do teste vazio",
          detail: `${testLabel}: marque o status do teste.`,
          tab: "tests",
          targetId,
          blockKey: entry.key,
          testId: test.id,
        });
      }

      if (test.result.legacyImages.length === 0) {
        summary.issues.push({
          id: `missing-legacy-image-${referenceKey}`,
          severity: "danger",
          label: "Imagem do legado ausente",
          detail: `${testLabel}: adicione evidencia em Legado.`,
          tab: "tests",
          targetId,
          blockKey: entry.key,
          testId: test.id,
        });
      }

      if (test.result.newImages.length === 0) {
        summary.issues.push({
          id: `missing-new-image-${referenceKey}`,
          severity: "danger",
          label: "Imagem do novo ausente",
          detail: `${testLabel}: adicione evidencia em Novo.`,
          tab: "tests",
          targetId,
          blockKey: entry.key,
          testId: test.id,
        });
      }

      if (hasProblemStatus(test.result) && !test.result.observations.trim()) {
        summary.issues.push({
          id: `missing-problem-observation-${referenceKey}`,
          severity: "warning",
          label: "Observacao obrigatoria",
          detail: `${testLabel}: status com problema ou erro exige observacao.`,
          tab: "tests",
          targetId,
          blockKey: entry.key,
          testId: test.id,
        });
      }
    });
  });

  return summary;
}

function buildTeaReviewSummary(documentData: TeaDocument): TeaReviewSummary {
  const summary: TeaReviewSummary = {
    activityCount: documentData.activities.length,
    imageCount:
      documentData.activityImages.length +
      documentData.activities.reduce(
        (total, activity) =>
          total +
          countTeaContentImages(activity.blocks) +
          activity.subActivities.reduce(
            (subTotal, subActivity) =>
              subTotal + countTeaContentImages(subActivity.blocks),
            0,
          ),
        0,
      ),
    issues: [],
  };

  if (!documentData.metadata.serviceOrder.trim()) {
    summary.issues.push({
      id: "missing-service-order",
      severity: "danger",
      category: "document",
      label: "Ordem de serviço vazia",
      detail: "Preencha a ordem de serviço do TEA.",
      tab: "document",
      targetId: "tea-metadata-service-order",
    });
  }

  if (!documentData.metadata.phase.trim()) {
    summary.issues.push({
      id: "missing-phase",
      severity: "warning",
      category: "document",
      label: "Fase/Etapa vazia",
      detail: "Preencha a fase ou etapa do TEA.",
      tab: "document",
      targetId: "tea-metadata-phase",
    });
  }

  if (!documentData.metadata.ticket.trim()) {
    summary.issues.push({
      id: "missing-ticket",
      severity: "warning",
      category: "document",
      label: "Chamado vazio",
      detail: "Preencha o chamado relacionado ao TEA.",
      tab: "document",
      targetId: "tea-metadata-ticket",
    });
  }

  if (!documentData.metadata.subject.trim()) {
    summary.issues.push({
      id: "missing-subject",
      severity: "danger",
      category: "document",
      label: "Assunto vazio",
      detail: "Preencha o assunto usado no cabeçalho e nome do arquivo.",
      tab: "document",
      targetId: "tea-metadata-subject",
    });
  }

  if (!documentData.metadata.date.trim()) {
    summary.issues.push({
      id: "missing-date",
      severity: "warning",
      category: "document",
      label: "Data vazia",
      detail: "Preencha a data do TEA.",
      tab: "document",
      targetId: "tea-metadata-date",
    });
  }

  if (!documentData.metadata.author.trim()) {
    summary.issues.push({
      id: "missing-author",
      severity: "warning",
      category: "document",
      label: "Elaborado por vazio",
      detail: "Preencha o responsavel pela elaboracao do TEA.",
      tab: "document",
      targetId: "tea-metadata-author",
    });
  }

  if (!documentData.overview.trim()) {
    summary.issues.push({
      id: "missing-overview",
      severity: "danger",
      category: "document",
      label: "Visão geral vazia",
      detail: "Preencha a seção 1 do documento.",
      tab: "document",
      targetId: "tea-overview",
    });
  }

  if (hasIncompleteTeaBoldMarkup(documentData.overview)) {
    summary.issues.push({
      id: "overview-incomplete-bold",
      severity: "warning",
      category: "document",
      label: "Negrito incompleto na visao geral",
      detail: "Revise os pares **texto** na secao 1.",
      tab: "document",
      targetId: "tea-overview",
    });
  }

  if (!documentData.activityIntro.trim()) {
    summary.issues.push({
      id: "missing-activity-intro",
      severity: "warning",
      category: "document",
      label: "Texto inicial vazio",
      detail: "Preencha o texto introdutor das atividades realizadas.",
      tab: "document",
      targetId: "tea-activity-intro",
    });
  }

  if (hasIncompleteTeaBoldMarkup(documentData.activityIntro)) {
    summary.issues.push({
      id: "activity-intro-incomplete-bold",
      severity: "warning",
      category: "document",
      label: "Negrito incompleto no texto inicial",
      detail: "Revise os pares **texto** no texto inicial das atividades.",
      tab: "document",
      targetId: "tea-activity-intro",
    });
  }

  documentData.activityImages.forEach((image, imageIndex) => {
    addTeaImageIssues(summary, image, {
      issuePrefix: `general-image-${image.id}`,
      detailPrefix: `Imagem geral ${imageIndex + 1}`,
      targetId: "tea-activity-intro",
      tab: "document",
      severity: "warning",
    });
  });

  if (documentData.activities.length === 0) {
    summary.issues.push({
      id: "missing-activities",
      severity: "danger",
      category: "activity",
      label: "Sem atividades",
      detail: "Adicione ao menos uma atividade realizada.",
      tab: "activities",
    });
  }

  documentData.activities.forEach((activity, index) => {
    if (!activity.title.trim()) {
      summary.issues.push({
        id: `activity-title-${activity.id}`,
        severity: "danger",
        category: "activity",
        label: "Atividade sem título",
        detail: `Atividade 2.${index + 1}`,
        tab: "activities",
        targetId: `tea-activity-title-${toDomId(activity.id)}`,
        activityId: activity.id,
      });
    }

    if (!hasUsefulTeaContent(activity.blocks)) {
      summary.issues.push({
        id: `activity-content-${activity.id}`,
        severity: "danger",
        category: "activity",
        label: "Atividade sem conteúdo",
        detail: `Atividade 2.${index + 1}: adicione texto, lista ou imagem.`,
        tab: "activities",
        targetId: `tea-activity-${toDomId(activity.id)}`,
        activityId: activity.id,
      });
    }

    activity.blocks.forEach((block, blockIndex) => {
      addTeaContentBlockIssues(summary, block, {
        issuePrefix: `activity-${activity.id}-block-${block.id}`,
        detailPrefix: `Atividade 2.${index + 1}, bloco ${blockIndex + 1}`,
        severity: "warning",
        activityId: activity.id,
      });
    });

    activity.subActivities.forEach((subActivity, subIndex) => {
      if (!subActivity.title.trim()) {
        summary.issues.push({
          id: `subactivity-title-${subActivity.id}`,
          severity: "warning",
          category: "activity",
          label: "Subtópico sem título",
          detail: `Subtópico 2.${index + 1}.${subIndex + 1}`,
          tab: "activities",
          targetId: `tea-subactivity-title-${toDomId(subActivity.id)}`,
          activityId: activity.id,
          subActivityId: subActivity.id,
        });
      }

      if (!hasUsefulTeaContent(subActivity.blocks)) {
        summary.issues.push({
          id: `subactivity-content-${subActivity.id}`,
          severity: "warning",
          category: "activity",
          label: "Subtópico sem conteúdo",
          detail: `Subtópico 2.${index + 1}.${subIndex + 1}: adicione texto, lista ou imagem.`,
          tab: "activities",
          targetId: `tea-subactivity-${toDomId(subActivity.id)}`,
          activityId: activity.id,
          subActivityId: subActivity.id,
        });
      }

      subActivity.blocks.forEach((block, blockIndex) => {
        addTeaContentBlockIssues(summary, block, {
          issuePrefix: `subactivity-${subActivity.id}-block-${block.id}`,
          detailPrefix: `Subtopico 2.${index + 1}.${subIndex + 1}, bloco ${blockIndex + 1}`,
          severity: "warning",
          activityId: activity.id,
          subActivityId: subActivity.id,
        });
      });
    });
  });

  return summary;
}

function addTeaContentBlockIssues(
  summary: TeaReviewSummary,
  block: TeaContentBlock,
  context: {
    issuePrefix: string;
    detailPrefix: string;
    severity: TeaReviewSeverity;
    activityId: string;
    subActivityId?: string;
  },
): void {
  const targetId = getTeaContentBlockReviewTargetId(block);
  const commonIssueFields = {
    tab: "activities" as const,
    category: "activity" as const,
    targetId,
    activityId: context.activityId,
    subActivityId: context.subActivityId,
    blockId: block.id,
  };

  if (block.type === "text") {
    if (!block.text.trim()) {
      summary.issues.push({
        id: `${context.issuePrefix}-empty-text`,
        severity: context.severity,
        label: "Bloco de texto vazio",
        detail: `${context.detailPrefix}: preencha ou remova o bloco de texto.`,
        ...commonIssueFields,
      });
    }

    if (hasIncompleteTeaBoldMarkup(block.text)) {
      summary.issues.push({
        id: `${context.issuePrefix}-incomplete-bold`,
        severity: "warning",
        label: "Negrito incompleto",
        detail: `${context.detailPrefix}: revise os pares **texto**.`,
        ...commonIssueFields,
      });
    }

    return;
  }

  if (block.type === "list") {
    if (!block.items.some((item) => item.text.trim())) {
      summary.issues.push({
        id: `${context.issuePrefix}-empty-list`,
        severity: context.severity,
        label: "Lista vazia",
        detail: `${context.detailPrefix}: adicione ao menos um item ou remova a lista.`,
        ...commonIssueFields,
      });
    }

    block.items.forEach((item, itemIndex) => {
      if (hasIncompleteTeaBoldMarkup(item.text)) {
        summary.issues.push({
          id: `${context.issuePrefix}-item-${item.id}-incomplete-bold`,
          severity: "warning",
          label: "Negrito incompleto na lista",
          detail: `${context.detailPrefix}, item ${itemIndex + 1}: revise os pares **texto**.`,
          ...commonIssueFields,
        });
      }
    });

    return;
  }

  if (block.images.length === 0) {
    summary.issues.push({
      id: `${context.issuePrefix}-empty-images`,
      severity: context.severity,
      ...commonIssueFields,
      category: "image",
      label: "Bloco de imagens vazio",
      detail: `${context.detailPrefix}: adicione uma imagem ou remova o bloco.`,
    });
  }

  block.images.forEach((image, imageIndex) => {
    addTeaImageIssues(summary, image, {
      issuePrefix: `${context.issuePrefix}-image-${image.id}`,
      detailPrefix: `${context.detailPrefix}, imagem ${imageIndex + 1}`,
      targetId,
      severity: context.severity,
      activityId: context.activityId,
      subActivityId: context.subActivityId,
      blockId: block.id,
    });
  });
}

function addTeaImageIssues(
  summary: TeaReviewSummary,
  image: EvidenceImage,
  context: {
    issuePrefix: string;
    detailPrefix: string;
    targetId: string;
    tab?: TeaTab;
    severity: TeaReviewSeverity;
    activityId?: string;
    subActivityId?: string;
    blockId?: string;
  },
): void {
  const commonIssueFields = {
    tab: context.tab ?? ("activities" as const),
    category: "image" as const,
    targetId: context.targetId,
    activityId: context.activityId,
    subActivityId: context.subActivityId,
    blockId: context.blockId,
  };

  if (!image.dataUrl) {
    summary.issues.push({
      id: `${context.issuePrefix}-missing-data`,
      severity: "danger",
      label: "Imagem sem dados",
      detail: `${context.detailPrefix}: recarregue ou substitua a imagem.`,
      ...commonIssueFields,
    });
  }
}

function hasIncompleteTeaBoldMarkup(value: string): boolean {
  return ((value.match(/\*\*/g) ?? []).length % 2) === 1;
}

function createBlockTest(id: string, title = ""): PermissionBlockTest {
  return {
    id,
    title,
    result: createEmptyTestResult(),
    correction: createEmptyTestCorrection(),
  };
}

function formatPermissionBulk(groups: PermissionGroup[]): string {
  return groups
    .map((macro) => {
      const macroLine = formatPermissionBulkLine(macro);
      const microLines = macro.microPermissions.map(
        (micro) => `  ${formatPermissionBulkLine(micro)}`,
      );

      return [macroLine, ...microLines].join("\n");
    })
    .join("\n");
}

function formatPermissionBulkLine(permission: PermissionItem): string {
  return [permission.code, permission.label].filter(Boolean).join(" - ");
}

function parsePermissionBulk(value: string): PermissionGroup[] {
  const groups: PermissionGroup[] = [];
  let currentMacro: PermissionGroup | undefined;

  value.split(/\r?\n/).forEach((line) => {
    if (!line.trim()) {
      return;
    }

    const permission = parsePermissionLine(line);
    const isMicro = /^\s+/.test(line) || line.trim().startsWith("-");

    if (!isMicro) {
      currentMacro = {
        id: "",
        code: permission.code,
        label: permission.label,
        selected: true,
        microPermissions: [],
      };
      groups.push(currentMacro);
      return;
    }

    if (!currentMacro) {
      return;
    }

    currentMacro.microPermissions.push({
      id: "",
      code: permission.code,
      label: permission.label,
      selected: true,
    });
  });

  return groups.filter((group) => group.code && group.microPermissions.length > 0);
}

function parsePermissionLine(line: string): Pick<PermissionItem, "code" | "label"> {
  const cleaned = line.trim().replace(/^-\s*/, "");
  const parenthesized = cleaned.match(/^([A-Za-z0-9_-]+)\s*\(([^)]+)\)$/);

  if (parenthesized) {
    return {
      code: parenthesized[1].toUpperCase(),
      label: parenthesized[2].trim(),
    };
  }

  const [code = "", ...labelParts] = cleaned.split(/\s+-\s+|:\s+/);

  return {
    code: code.trim().toUpperCase(),
    label: labelParts.join(" - ").trim(),
  };
}

function splitBulkLines(value: string): string[] {
  if (!value) {
    return [];
  }

  return value.replace(/\r\n/g, "\n").split("\n").map((line) => line.trim());
}

function teaItemsFromBulk(value: string): string[] {
  return splitBulkLines(value).filter(Boolean);
}

function updateTeaItemsFromBulk(
  currentItems: { id: string; text: string }[],
  value: string,
): { id: string; text: string }[] {
  return teaItemsFromBulk(value).map((text, index) => ({
    id: currentItems[index]?.id ?? createId(),
    text,
  }));
}

function createTeaContentBlock(type: TeaContentBlockType): TeaContentBlock {
  if (type === "text") {
    return {
      id: createId(),
      type: "text",
      text: "",
    };
  }

  if (type === "list") {
    return {
      id: createId(),
      type: "list",
      items: [],
    };
  }

  return {
    id: createId(),
    type: "images",
    images: [],
  };
}

function findTeaSubActivity(
  documentData: TeaDocument,
  activityId: string,
  subActivityId: string,
): TeaSubActivity | undefined {
  return documentData.activities
    .find((activity) => activity.id === activityId)
    ?.subActivities.find((subActivity) => subActivity.id === subActivityId);
}

function buildTeaActivitySelectOptions(
  documentData: TeaDocument,
  sourceActivityId: string,
): Array<{ value: string; label: string }> {
  return documentData.activities.flatMap((activity, index) =>
    activity.id === sourceActivityId
      ? []
      : [{
      value: activity.id,
      label: formatTeaEditorTitle(`2.${index + 1}`, activity.title, "Atividade sem título"),
    }],
  );
}

function areTeaActivityOutlineTargetsOpen(
  documentData: TeaDocument,
  collapsedActivities: Record<string, boolean>,
  collapsedSubActivities: Record<string, boolean>,
): boolean {
  return documentData.activities.every(
    (activity) =>
      !collapsedActivities[activity.id] &&
      activity.subActivities.every((subActivity) => !collapsedSubActivities[subActivity.id]),
  );
}

function summarizeTeaReviewIssues(issues: TeaReviewIssue[]): TeaInlineReviewSummary {
  return issues.reduce<TeaInlineReviewSummary>(
    (summary, issue) => ({
      total: summary.total + 1,
      danger: summary.danger + (issue.severity === "danger" ? 1 : 0),
      warning: summary.warning + (issue.severity === "warning" ? 1 : 0),
    }),
    { total: 0, danger: 0, warning: 0 },
  );
}

function getTeaReviewTone(review: TeaInlineReviewSummary): "gray" | "red" | "yellow" {
  if (review.danger > 0) {
    return "red";
  }

  if (review.warning > 0) {
    return "yellow";
  }

  return "gray";
}

function formatTeaEditorTitle(
  number: string,
  title: string,
  fallback: string,
): string {
  return `${number} ${title.trim() || fallback}`;
}

function formatTeaComposerTitle(
  kind: "activity" | "subActivity",
  number: string,
  title: string,
): string {
  const trimmedTitle = title.trim();

  if (trimmedTitle) {
    return `Blocos de ${trimmedTitle}`;
  }

  return kind === "activity"
    ? `Blocos da atividade ${number}`
    : `Blocos do subtópico ${number}`;
}

function formatConfirmationSubject(value: string, fallback: string): string {
  const trimmedValue = value.trim();

  return trimmedValue ? `"${trimmedValue}"` : fallback;
}

function getTeaBlockDisplayLabel(block: TeaContentBlock): string {
  return `O bloco de ${teaContentBlockLabels[block.type].toLowerCase()}`;
}

function buildTeaActivitySummaryItems(activity: TeaActivity): string[] {
  const blockCount =
    activity.blocks.length +
    activity.subActivities.reduce((total, subActivity) => total + subActivity.blocks.length, 0);
  const imageCount =
    countTeaContentImages(activity.blocks) +
    activity.subActivities.reduce(
      (total, subActivity) => total + countTeaContentImages(subActivity.blocks),
      0,
    );

  return [
    formatTeaCount(blockCount, "bloco", "blocos"),
    formatTeaCount(activity.subActivities.length, "subtópico", "subtópicos"),
    formatTeaCount(imageCount, "imagem", "imagens"),
  ];
}

function buildTeaSubActivitySummaryItems(subActivity: TeaSubActivity): string[] {
  return [
    formatTeaCount(subActivity.blocks.length, "bloco", "blocos"),
    formatTeaCount(countTeaContentImages(subActivity.blocks), "imagem", "imagens"),
  ];
}

function buildTeaContentBlockSummaryItem(block: TeaContentBlock): string {
  if (block.type === "text") {
    const characterCount = block.text.trim().length;
    return characterCount > 0
      ? formatTeaCount(characterCount, "caractere", "caracteres")
      : "texto vazio";
  }

  if (block.type === "list") {
    return formatTeaCount(
      block.items.filter((item) => item.text.trim()).length,
      "item",
      "itens",
    );
  }

  return formatTeaCount(block.images.length, "imagem", "imagens");
}

function formatTeaCount(count: number, singular: string, plural: string): string {
  return `${count} ${count === 1 ? singular : plural}`;
}

function hasTeaActivityRemovalContent(activity: TeaActivity): boolean {
  return (
    Boolean(activity.title.trim()) ||
    activity.blocks.length > 0 ||
    activity.subActivities.length > 0
  );
}

function hasTeaSubActivityRemovalContent(subActivity: TeaSubActivity): boolean {
  return Boolean(subActivity.title.trim()) || subActivity.blocks.length > 0;
}

function hasMeaningfulTeaBlockContent(block: TeaContentBlock): boolean {
  if (block.type === "text") {
    return Boolean(block.text.trim());
  }

  if (block.type === "list") {
    return block.items.some((item) => item.text.trim());
  }

  return block.images.length > 0;
}

function hasMeaningfulEvidenceImageContent(image: EvidenceImage): boolean {
  return Boolean(image.label.trim() || image.name.trim() || image.dataUrl);
}

async function duplicateTeaContentBlock(block: TeaContentBlock): Promise<TeaContentBlock> {
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

  let failedToPersist = false;
  const images = block.images.map((image) => ({
    ...image,
    id: createId(),
  }));

  try {
    await saveEvidenceImageDataBatch(
      images.flatMap((image) =>
        image.dataUrl ? [{ id: image.id, dataUrl: image.dataUrl }] : [],
      ),
    );
  } catch {
    failedToPersist = true;
  }

  if (failedToPersist) {
    window.alert("Nao foi possivel duplicar uma ou mais imagens no rascunho do navegador.");
  }

  return {
    ...block,
    id: createId(),
    images,
  };
}

async function duplicateTeaSubActivity(subActivity: TeaSubActivity): Promise<TeaSubActivity> {
  return {
    ...subActivity,
    id: createId(),
    blocks: await Promise.all(subActivity.blocks.map(duplicateTeaContentBlock)),
  };
}

function moveItemById<T extends { id: string }>(
  items: T[],
  itemId: string,
  direction: MoveDirection,
): T[] {
  const currentIndex = items.findIndex((item) => item.id === itemId);

  if (currentIndex < 0) {
    return items;
  }

  const nextIndex = direction === "up" ? currentIndex - 1 : currentIndex + 1;

  if (nextIndex < 0 || nextIndex >= items.length) {
    return items;
  }

  const nextItems = [...items];
  [nextItems[currentIndex], nextItems[nextIndex]] = [
    nextItems[nextIndex],
    nextItems[currentIndex],
  ];

  return nextItems;
}

function setCollapsedMapValue(
  current: Record<string, boolean>,
  itemId: string,
  collapsed: boolean,
): Record<string, boolean> {
  if (collapsed) {
    return {
      ...current,
      [itemId]: true,
    };
  }

  const { [itemId]: _removed, ...rest } = current;
  return rest;
}

function countTeaContentImages(blocks: TeaContentBlock[]): number {
  return blocks.reduce(
    (total, block) => total + (block.type === "images" ? block.images.length : 0),
    0,
  );
}

function hasUsefulTeaContent(blocks: TeaContentBlock[]): boolean {
  return blocks.some((block) => {
    if (block.type === "text") {
      return Boolean(block.text.trim());
    }

    if (block.type === "list") {
      return block.items.some((item) => item.text.trim());
    }

    return block.images.length > 0;
  });
}

function blocksByPermissionCode(documentData: OtDocument): Map<string, PermissionBlock> {
  const blocks = new Map<string, PermissionBlock>();

  documentData.permissionGroups.forEach((macro) => {
    macro.microPermissions.forEach((micro) => {
      const block = documentData.permissionBlocks[createPermissionKey(macro.id, micro.id)];
      if (block) {
        blocks.set(permissionCodeKey(macro.code, micro.code), block);
      }
    });
  });

  return blocks;
}

function permissionCodeKey(macroCode: string, microCode: string): string {
  return `${normalizeTextKey(macroCode)}:${normalizeTextKey(microCode)}`;
}

function normalizeTextKey(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function draftStatusColor(status: DraftStatus): "green" | "yellow" | "blue" | "red" {
  if (status === "Rascunho grande demais") {
    return "red";
  }

  if (status === "Salvando...") {
    return "blue";
  }

  if (status === "Alterações pendentes") {
    return "yellow";
  }

  return "green";
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

function createEmptyBlock(): PermissionBlock {
  return { tests: [] };
}

function createTestReferenceKey(blockKey: string, testId: string): string {
  return `${blockKey}::${testId}`;
}

function getMacroIdFromBlockKey(blockKey: string): string {
  return blockKey.split(":")[0] ?? "";
}

function toDomId(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]/g, "-");
}

function removePermissionBlocks(
  blocks: Record<string, PermissionBlock>,
  shouldRemove: (key: string) => boolean,
): Record<string, PermissionBlock> {
  const removedKeys = new Set(Object.keys(blocks).filter(shouldRemove));

  return Object.fromEntries(
    Object.entries(blocks)
      .filter(([key]) => !removedKeys.has(key))
      .map(([key, block]) => [key, block]),
  );
}

function createId(): string {
  return window.crypto.randomUUID();
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${Math.round(bytes / 1024)} KB`;
  }

  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function getPastedImageFiles(clipboardData: DataTransfer): File[] {
  const itemFiles = Array.from(clipboardData.items)
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter((file): file is File => file !== null && isImageFile(file));

  if (itemFiles.length > 0) {
    return itemFiles;
  }

  return Array.from(clipboardData.files).filter(isImageFile);
}

function isImageFile(file: File): boolean {
  return (
    file.type.startsWith("image/") ||
    /\.(avif|bmp|gif|jfif|jpe?g|png|svg|webp)$/i.test(file.name)
  );
}
