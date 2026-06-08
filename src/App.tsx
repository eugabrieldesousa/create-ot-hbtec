import { memo, useCallback, useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import type { ClipboardEvent, DragEvent, ReactNode } from "react";
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
  Modal,
  Paper,
  SegmentedControl,
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
  FileText,
  FileUp,
  ImagePlus,
  ListChecks,
  Moon,
  Plus,
  RotateCcw,
  Save,
  SearchCheck,
  Sun,
  Trash2,
  Wand2,
  X,
} from "lucide-react";
import { checkLabels, checkOrder, createEmptyTestResult, createPermissionKey } from "./defaultDocument";
import { exportOtDocument, exportTeaDocument } from "./docxExport";
import { parseDocxFile } from "./docxImport";
import type { DocxImportResult } from "./docxImport";
import {
  deleteEvidenceImageData,
  hydrateDocumentImages,
  hydrateTeaDocumentImages,
  persistEmbeddedEvidenceImages,
  persistEmbeddedTeaImages,
  saveEvidenceImageData,
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
  TeaDocument,
  TeaSubActivity,
  TeaTextItem,
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

type ActiveTab = "document" | "permissions" | "tests" | "review" | "help";
type TeaTab = "document" | "activities" | "review";
type DocumentKind = "ot" | "tea";

type DraftStatus =
  | "Alterações pendentes"
  | "Salvando..."
  | "Rascunho salvo"
  | "Rascunho grande demais";

type ReviewSummary = {
  selectedPermissions: number;
  testCount: number;
  imageCount: number;
  issues: ReviewIssue[];
};

type ReviewIssue = {
  id: string;
  label: string;
  detail: string;
  tab: Exclude<ActiveTab, "help">;
  targetId?: string;
  blockKey?: string;
  testId?: string;
};

type TeaReviewSummary = {
  activityCount: number;
  imageCount: number;
  issues: TeaReviewIssue[];
};

type TeaReviewIssue = {
  id: string;
  label: string;
  detail: string;
  tab: TeaTab;
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

type PermissionBlockEditorProps = Omit<PermissionBlockGroupProps, "macro" | "entries"> & {
  blockKey: string;
  entry: PermissionBlockEntry;
  block: PermissionBlock;
  sourceBlock: PermissionBlock;
};

type BlockTestEditorProps = {
  blockKey: string;
  index: number;
  test: PermissionBlockTest;
  selfReferenceKey: string;
  isExpanded: boolean;
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

const problemCheckKeys: CheckKey[] = ["possibleIssue", "bothIssue", "newIssue", "errorReport"];

const quickCheckLabels: Record<CheckKey, string> = {
  sameBehavior: "OK",
  possibleIssue: "Possível",
  bothIssue: "Ambos",
  newIssue: "Novo",
  errorReport: "Erros",
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
          "A Revisão mostra pendências, contadores de testes e imagens. Use Próxima pendência para ir direto ao primeiro ponto incompleto.",
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
        title: "Padrão em todos",
        description:
          "Adiciona Criação, Edição, Consulta e Exclusão em todos os blocos de permissão selecionados, sem duplicar títulos que já existem no bloco.",
      },
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
  const [expandedTests, setExpandedTests] = useState<Record<string, boolean>>({});
  const [activeTab, setActiveTab] = useState<ActiveTab>("document");
  const [teaActiveTab, setTeaActiveTab] = useState<TeaTab>("document");
  const [testBlockFilter, setTestBlockFilter] = useState<TestBlockFilter>("all");
  const [isExporting, setIsExporting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [importPreview, setImportPreview] = useState<DocxImportResult | null>(null);
  const [draftStatus, setDraftStatus] = useState<DraftStatus>("Rascunho salvo");
  const [permissionBulkText, setPermissionBulkText] = useState(() =>
    formatPermissionBulk(documentData.permissionGroups),
  );
  const documentDataRef = useRef(documentData);
  const teaDataRef = useRef(teaData);
  const documentKindRef = useRef<DocumentKind>(documentKind);
  const saveTimerRef = useRef<number | undefined>();
  const idleSaveRef = useRef<number | undefined>();
  const isDarkMode = colorScheme === "dark";
  const deferredDocumentData = useDeferredValue(documentData);
  const deferredTeaData = useDeferredValue(teaData);

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

  useEffect(() => {
    documentDataRef.current = documentData;
  }, [documentData]);

  useEffect(() => {
    teaDataRef.current = teaData;
  }, [teaData]);

  useEffect(() => {
    documentKindRef.current = documentKind;
  }, [documentKind]);

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

    void prepareImages();

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    const handleBeforeUnload = () => {
      flushDraft();
    };

    window.addEventListener("beforeunload", handleBeforeUnload);

    return () => {
      window.removeEventListener("beforeunload", handleBeforeUnload);
    };
  }, [flushDraft]);

  const reviewSummary = useMemo(
    () => buildReviewSummary(deferredDocumentData, permissionBlockEntries),
    [deferredDocumentData, permissionBlockEntries],
  );

  const teaReviewSummary = useMemo(
    () => buildTeaReviewSummary(deferredTeaData),
    [deferredTeaData],
  );

  const updateDocument = useCallback((updater: (current: OtDocument) => OtDocument): void => {
    setDocumentData((current) => updater(current));
  }, []);

  const updateTeaDocument = useCallback((updater: (current: TeaDocument) => TeaDocument): void => {
    setTeaData((current) => updater(current));
  }, []);

  function selectDocumentKind(nextKind: DocumentKind): void {
    if (nextKind === documentKindRef.current) {
      return;
    }

    flushDraft();
    setDocumentKind(nextKind);
  }

  function updateTeaMetadata(field: keyof TeaDocument["metadata"], value: string): void {
    updateTeaDocument((current) => ({
      ...current,
      metadata: {
        ...current.metadata,
        [field]: value,
      },
    }));
  }

  function updateTeaActivity(
    activityId: string,
    updater: (activity: TeaActivity) => TeaActivity,
  ): void {
    updateTeaDocument((current) => ({
      ...current,
      activities: current.activities.map((activity) =>
        activity.id === activityId ? updater(activity) : activity,
      ),
    }));
  }

  function addTeaActivity(): void {
    updateTeaDocument((current) => ({
      ...current,
      activities: [
        ...current.activities,
        {
          id: createId(),
          title: "",
          description: "",
          items: [],
          images: [],
          subActivities: [],
        },
      ],
    }));
  }

  function removeTeaActivity(activityId: string): void {
    updateTeaDocument((current) => ({
      ...current,
      activities: current.activities.filter((activity) => activity.id !== activityId),
    }));
  }

  function updateTeaActivityItems(activityId: string, value: string): void {
    const items = teaItemsFromBulk(value);

    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      items: items.map((text, index) => ({
        id: activity.items[index]?.id ?? createId(),
        text,
      })),
    }));
  }

  function addTeaSubActivity(activityId: string): void {
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: [
        ...activity.subActivities,
        {
          id: createId(),
          title: "",
          description: "",
          items: [],
          images: [],
        },
      ],
    }));
  }

  function updateTeaSubActivity(
    activityId: string,
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ): void {
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: activity.subActivities.map((subActivity) =>
        subActivity.id === subActivityId ? updater(subActivity) : subActivity,
      ),
    }));
  }

  function removeTeaSubActivity(activityId: string, subActivityId: string): void {
    updateTeaActivity(activityId, (activity) => ({
      ...activity,
      subActivities: activity.subActivities.filter(
        (subActivity) => subActivity.id !== subActivityId,
      ),
    }));
  }

  function updateTeaSubActivityItems(
    activityId: string,
    subActivityId: string,
    value: string,
  ): void {
    const items = teaItemsFromBulk(value);

    updateTeaSubActivity(activityId, subActivityId, (subActivity) => ({
      ...subActivity,
      items: items.map((text, index) => ({
        id: subActivity.items[index]?.id ?? createId(),
        text,
      })),
    }));
  }

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
    updateDocument((current) => ({
      ...current,
      accessSteps: current.accessSteps.filter((step) => step.id !== stepId),
    }));
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

  function addMacroGroup(): void {
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
  }

  function updateMacroGroup(macroId: string, updates: Partial<PermissionItem>): void {
    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.map((macro) =>
        macro.id === macroId ? { ...macro, ...updates } : macro,
      ),
    }));
  }

  function removeMacroGroup(macroId: string): void {
    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.filter((macro) => macro.id !== macroId),
      permissionBlocks: removePermissionBlocks(current.permissionBlocks, (key) =>
        key.startsWith(`${macroId}:`),
      ),
    }));
  }

  function addMicroPermission(macroId: string): void {
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
  }

  function updateMicroPermission(
    macroId: string,
    microId: string,
    updates: Partial<PermissionItem>,
  ): void {
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
  }

  function removeMicroPermission(macroId: string, microId: string): void {
    const blockKey = createPermissionKey(macroId, microId);

    updateDocument((current) => ({
      ...current,
      permissionGroups: current.permissionGroups.map((macro) =>
        macro.id === macroId
          ? {
              ...macro,
              microPermissions: macro.microPermissions.filter(
                (micro) => micro.id !== microId,
              ),
            }
          : macro,
      ),
      permissionBlocks: removePermissionBlocks(
        current.permissionBlocks,
        (key) => key === blockKey,
      ),
    }));
  }

  function applyPermissionBulk(): void {
    const parsedGroups = parsePermissionBulk(permissionBulkText);

    if (parsedGroups.length === 0) {
      return;
    }

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

    setTestExpansion(createTestReferenceKey(blockKey, testId), true);

    updateBlock(blockKey, (block) => ({
      ...block,
      tests: [...block.tests, createBlockTest(testId)],
    }));
  }, [setTestExpansion, updateBlock]);

  const addStandardTests = useCallback((blockKey: string): void => {
    const tests = standardTestTitles.map((title) => createBlockTest(createId(), title));
    const firstReferenceKey = createTestReferenceKey(blockKey, tests[0]?.id ?? "");

    if (tests[0]) {
      setTestExpansion(firstReferenceKey, true);
    }

    updateBlock(blockKey, (block) => ({
      ...block,
      tests: [...block.tests, ...tests],
    }));
  }, [setTestExpansion, updateBlock]);

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
                })),
              },
            ]),
        ),
      },
    }));
  }, [permissionBlockEntries, updateDocument]);

  const duplicateBlockTest = useCallback((blockKey: string, testId: string): void => {
    const duplicatedId = createId();

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
      };
      const tests = [...block.tests];
      tests.splice(index + 1, 0, duplicated);

      return { ...block, tests };
    });
  }, [setTestExpansion, updateBlock]);

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
    const referenceKey = createTestReferenceKey(blockKey, testId);

    setExpandedTests((current) => {
      const { [referenceKey]: _removed, ...rest } = current;
      return rest;
    });

    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.filter((test) => test.id !== testId),
    }));
  }, [updateBlock]);

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
  }, [filteredPermissionBlockGroups]);

  function handleReviewIssueClick(issue: ReviewIssue): void {
    setActiveTab(issue.tab);

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

  function goToNextIssue(): void {
    const [issue] = reviewSummary.issues;

    if (issue) {
      handleReviewIssueClick(issue);
    }
  }

  function handleTeaReviewIssueClick(issue: TeaReviewIssue): void {
    setTeaActiveTab(issue.tab);
  }

  async function handleImportFile(file: File | null): Promise<void> {
    if (!file) {
      return;
    }

    setIsImporting(true);

    try {
      setImportPreview(await parseDocxFile(file));
    } catch {
      window.alert("Nao foi possivel importar este DOCX.");
    } finally {
      setIsImporting(false);
    }
  }

  async function confirmImport(): Promise<void> {
    if (!importPreview) {
      return;
    }

    try {
      await persistEmbeddedEvidenceImages(importPreview.document);
    } catch {
      window.alert("O documento foi importado, mas algumas imagens podem nao ficar salvas no rascunho.");
    }

    setDocumentData(importPreview.document);
    setPermissionBulkText(formatPermissionBulk(importPreview.document.permissionGroups));
    setExpandedTests({});
    setActiveTab("review");
    setImportPreview(null);
  }

  async function handleExport(): Promise<void> {
    flushDraft();
    setIsExporting(true);

    try {
      if (documentKindRef.current === "tea") {
        await exportTeaDocument(teaDataRef.current);
      } else {
        await exportOtDocument(documentDataRef.current);
      }
    } finally {
      setIsExporting(false);
    }
  }

  async function handleClearDraft(): Promise<void> {
    if (!window.confirm("Limpar o rascunho atual?")) {
      return;
    }

    if (documentKindRef.current === "tea") {
      await clearTeaDraft();
      setTeaData(loadTeaDraft());
      setTeaActiveTab("document");
    } else {
      await clearDraft();
      const nextDocument = loadDraft();
      setDocumentData(nextDocument);
      setPermissionBulkText(formatPermissionBulk(nextDocument.permissionGroups));
      setExpandedTests({});
    }

    setDraftStatus("Rascunho salvo");
  }

  return (
    <>
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

            <Group gap="xs">
              <Badge
                variant="light"
                color={draftStatusColor(draftStatus)}
                leftSection={<Save size={14} />}
                h={30}
              >
                {draftStatus}
              </Badge>
              {documentKind === "ot" ? (
                <>
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
                  >
                    Importar DOCX
                  </Button>
                )}
              </FileButton>
              <Button
                variant="light"
                leftSection={<Wand2 size={17} />}
                disabled={permissionBlockEntries.length === 0}
                onClick={addStandardTestsToAll}
              >
                Padrão em todos
              </Button>
              <Button
                variant="light"
                color={reviewSummary.issues.length > 0 ? "yellow" : "gray"}
                leftSection={<SearchCheck size={17} />}
                disabled={reviewSummary.issues.length === 0}
                onClick={goToNextIssue}
              >
                Próxima pendência
              </Button>
                </>
              ) : null}
              <Tooltip label={isDarkMode ? "Ativar modo claro" : "Ativar modo escuro"}>
                <ActionIcon
                  variant="light"
                  color="gray"
                  size="lg"
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
                onClick={() => {
                  void handleClearDraft();
                }}
              >
                Limpar
              </Button>
              <Button
                leftSection={<Download size={17} />}
                onClick={handleExport}
                loading={isExporting}
              >
                Exportar DOCX
              </Button>
            </Group>
          </Group>
        </Paper>

        {documentKind === "ot" ? (
        <Tabs
          value={activeTab}
          onChange={(value) => {
            if (value) {
              setActiveTab(value as ActiveTab);
            }
          }}
          keepMounted={false}
          className="workspaceTabs"
        >
          <Tabs.List>
            <Tabs.Tab value="document" leftSection={<ClipboardList size={16} />}>
              Documento
            </Tabs.Tab>
            <Tabs.Tab value="permissions" leftSection={<ListChecks size={16} />}>
              Permissões
            </Tabs.Tab>
            <Tabs.Tab value="tests" leftSection={<CheckCircle2 size={16} />}>
              Testes
            </Tabs.Tab>
            <Tabs.Tab value="review" leftSection={<AlertCircle size={16} />}>
              Revisão
            </Tabs.Tab>
            <Tabs.Tab value="help" leftSection={<CircleHelp size={16} />}>
              FAQ
            </Tabs.Tab>
          </Tabs.List>

          <Tabs.Panel value="document" pt="md">
            <Stack gap="md">
        <Section title="Documento" tone="document">
          <Stack gap="sm">
            <div className="documentFields">
            <TextInput
              label="Tela"
              value={documentData.metadata.screen}
              onChange={(event) => updateMetadata("screen", event.currentTarget.value)}
            />
            <TextInput
              label="Responsável pelo teste"
              value={documentData.metadata.responsible}
              onChange={(event) => updateMetadata("responsible", event.currentTarget.value)}
            />
            <TextInput
              label="Data"
              type="date"
              value={documentData.metadata.date}
              onChange={(event) => updateMetadata("date", event.currentTarget.value)}
            />
            <TextInput
              label="Ambiente"
              value={documentData.metadata.environment}
              onChange={(event) => updateMetadata("environment", event.currentTarget.value)}
            />
            <TextInput
              label="Elaborada por"
              value={documentData.metadata.author}
              onChange={(event) => updateMetadata("author", event.currentTarget.value)}
            />
            </div>
            <Textarea
              label="Objetivo"
              minRows={4}
              styles={{ input: { resize: "vertical" } }}
              value={documentData.objective}
              onChange={(event) => {
                const value = event.currentTarget.value;

                updateDocument((current) => ({
                  ...current,
                  objective: value,
                }));
              }}
            />
          </Stack>
        </Section>

        <Section
          title="Passo a passo"
          tone="steps"
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={addStep}>
              Adicionar
            </Button>
          }
        >
          <Stack gap="xs">
            <Textarea
              label="Editar passos em lote"
              minRows={3}
              autosize
              value={documentData.accessSteps.map((step) => step.text).join("\n")}
              onChange={(event) => replaceAccessStepsFromBulk(event.currentTarget.value)}
            />
            {documentData.accessSteps.map((step, index) => (
              <Group key={step.id} align="flex-end" wrap="nowrap">
                <Badge color="gray" variant="outline" w={34} h={34}>
                  {index + 1}
                </Badge>
                <TextInput
                  label={index === 0 ? "Etapa" : undefined}
                  value={step.text}
                  onChange={(event) => updateStep(step.id, event.currentTarget.value)}
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
              <EmptyState actionLabel="Adicionar passo" onAction={addStep} />
            ) : null}
          </Stack>
        </Section>

            </Stack>
          </Tabs.Panel>

          <Tabs.Panel value="permissions" pt="md">
        <Section
          title="Permissões"
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
                      onClick={() =>
                        setPermissionBulkText(formatPermissionBulk(documentData.permissionGroups))
                      }
                    >
                      Carregar atual
                    </Button>
                    <Button variant="light" size="xs" onClick={applyPermissionBulk}>
                      Aplicar lista
                    </Button>
                  </Group>
                </Group>
                <Textarea
                  minRows={4}
                  autosize
                  value={permissionBulkText}
                  placeholder={"AO - Administrador Geral\n  AT - Atualização\n  SC - Somente Consulta"}
                  onChange={(event) => setPermissionBulkText(event.currentTarget.value)}
                />
              </Stack>
            </Paper>
            {documentData.permissionGroups.map((macro, index) => (
              <PermissionGroupEditor
                key={macro.id}
                index={index}
                macro={macro}
                onMacroChange={(updates) => updateMacroGroup(macro.id, updates)}
                onRemoveMacro={() => removeMacroGroup(macro.id)}
                onAddMicro={() => addMicroPermission(macro.id)}
                onMicroChange={(microId, updates) =>
                  updateMicroPermission(macro.id, microId, updates)
                }
                onRemoveMicro={(microId) => removeMicroPermission(macro.id, microId)}
              />
            ))}
            {documentData.permissionGroups.length === 0 ? (
              <EmptyState actionLabel="Adicionar macro" onAction={addMacroGroup} />
            ) : null}
          </Stack>
        </Section>

          </Tabs.Panel>

          <Tabs.Panel value="tests" pt="md">
        <Section
          title="Testes por permissão"
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
            <Group gap="xs" justify="flex-end" className="testExpansionActions">
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsDown size={15} />}
                disabled={visibleTestCount === 0}
                onClick={() => setVisibleTestsExpansion(true)}
              >
                Expandir todos
              </Button>
              <Button
                variant="light"
                size="xs"
                leftSection={<ChevronsUp size={15} />}
                disabled={visibleTestCount === 0}
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
              <Paper withBorder p="md" ta="center">
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

          <Tabs.Panel value="review" pt="md">
            <ReviewPanel summary={reviewSummary} onIssueClick={handleReviewIssueClick} />
          </Tabs.Panel>

          <Tabs.Panel value="help" pt="md">
            <FaqPanel />
          </Tabs.Panel>
        </Tabs>
        ) : (
          <TeaWorkspace
            documentData={teaData}
            activeTab={teaActiveTab}
            reviewSummary={teaReviewSummary}
            onTabChange={setTeaActiveTab}
            onMetadataChange={updateTeaMetadata}
            onOverviewChange={(overview) =>
              updateTeaDocument((current) => ({ ...current, overview }))
            }
            onActivityIntroChange={(activityIntro) =>
              updateTeaDocument((current) => ({ ...current, activityIntro }))
            }
            onActivityImagesChange={(updater) =>
              updateTeaDocument((current) => ({
                ...current,
                activityImages: updater(current.activityImages),
              }))
            }
            onAddActivity={addTeaActivity}
            onActivityChange={updateTeaActivity}
            onActivityRemove={removeTeaActivity}
            onActivityItemsChange={updateTeaActivityItems}
            onAddSubActivity={addTeaSubActivity}
            onSubActivityChange={updateTeaSubActivity}
            onSubActivityRemove={removeTeaSubActivity}
            onSubActivityItemsChange={updateTeaSubActivityItems}
            onReviewIssueClick={handleTeaReviewIssueClick}
          />
        )}
      </Stack>
    </Container>
    <ImportPreviewModal
      result={importPreview}
      onClose={() => setImportPreview(null)}
      onConfirm={() => {
        void confirmImport();
      }}
    />
    </>
  );
}

function TeaWorkspace({
  documentData,
  activeTab,
  reviewSummary,
  onTabChange,
  onMetadataChange,
  onOverviewChange,
  onActivityIntroChange,
  onActivityImagesChange,
  onAddActivity,
  onActivityChange,
  onActivityRemove,
  onActivityItemsChange,
  onAddSubActivity,
  onSubActivityChange,
  onSubActivityRemove,
  onSubActivityItemsChange,
  onReviewIssueClick,
}: {
  documentData: TeaDocument;
  activeTab: TeaTab;
  reviewSummary: TeaReviewSummary;
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
  onActivityItemsChange: (activityId: string, value: string) => void;
  onAddSubActivity: (activityId: string) => void;
  onSubActivityChange: (
    activityId: string,
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ) => void;
  onSubActivityRemove: (activityId: string, subActivityId: string) => void;
  onSubActivityItemsChange: (activityId: string, subActivityId: string, value: string) => void;
  onReviewIssueClick: (issue: TeaReviewIssue) => void;
}) {
  return (
    <Tabs
      value={activeTab}
      onChange={(value) => {
        if (value) {
          onTabChange(value as TeaTab);
        }
      }}
      keepMounted={false}
      className="workspaceTabs"
    >
      <Tabs.List>
        <Tabs.Tab value="document" leftSection={<FileText size={16} />}>
          Documento
        </Tabs.Tab>
        <Tabs.Tab value="activities" leftSection={<ListChecks size={16} />}>
          Atividades
        </Tabs.Tab>
        <Tabs.Tab value="review" leftSection={<AlertCircle size={16} />}>
          Revisão
        </Tabs.Tab>
      </Tabs.List>

      <Tabs.Panel value="document" pt="md">
        <Stack gap="md">
          <Section title="Documento TEA" tone="document">
            <Stack gap="sm">
              <div className="documentFields">
                <TextInput
                  label="Ordem de Serviço"
                  value={documentData.metadata.serviceOrder}
                  placeholder="OS2171 - Login/Menu/Prestador/Documentos Vencidos"
                  onChange={(event) =>
                    onMetadataChange("serviceOrder", event.currentTarget.value)
                  }
                />
                <TextInput
                  label="Fase/Etapa"
                  value={documentData.metadata.phase}
                  placeholder="Etapa 5"
                  onChange={(event) => onMetadataChange("phase", event.currentTarget.value)}
                />
                <TextInput
                  label="Chamado"
                  value={documentData.metadata.ticket}
                  placeholder="Chamado 202504000396"
                  onChange={(event) => onMetadataChange("ticket", event.currentTarget.value)}
                />
                <TextInput
                  label="Assunto"
                  value={documentData.metadata.subject}
                  placeholder="Telas - Novo Layout"
                  onChange={(event) => onMetadataChange("subject", event.currentTarget.value)}
                />
                <TextInput
                  label="Data"
                  type="date"
                  value={documentData.metadata.date}
                  onChange={(event) => onMetadataChange("date", event.currentTarget.value)}
                />
                <TextInput
                  label="Elaborado por"
                  value={documentData.metadata.author}
                  onChange={(event) => onMetadataChange("author", event.currentTarget.value)}
                />
              </div>
              <Textarea
                label="1. Visão geral"
                minRows={5}
                autosize
                styles={{ input: { resize: "vertical" } }}
                value={documentData.overview}
                onChange={(event) => onOverviewChange(event.currentTarget.value)}
              />
              <Textarea
                label="Texto inicial das atividades"
                minRows={2}
                autosize
                styles={{ input: { resize: "vertical" } }}
                value={documentData.activityIntro}
                onChange={(event) => onActivityIntroChange(event.currentTarget.value)}
              />
            </Stack>
          </Section>

          <Section title="Imagem geral" tone="steps">
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
          title="Atividades"
          tone="blocks"
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={onAddActivity}>
              Adicionar atividade
            </Button>
          }
        >
          <Stack gap="md">
            {documentData.activities.map((activity, index) => (
              <TeaActivityEditor
                key={activity.id}
                index={index}
                activity={activity}
                onChange={(updater) => onActivityChange(activity.id, updater)}
                onRemove={() => onActivityRemove(activity.id)}
                onItemsChange={(value) => onActivityItemsChange(activity.id, value)}
                onAddSubActivity={() => onAddSubActivity(activity.id)}
                onSubActivityChange={(subActivityId, updater) =>
                  onSubActivityChange(activity.id, subActivityId, updater)
                }
                onSubActivityRemove={(subActivityId) =>
                  onSubActivityRemove(activity.id, subActivityId)
                }
                onSubActivityItemsChange={(subActivityId, value) =>
                  onSubActivityItemsChange(activity.id, subActivityId, value)
                }
              />
            ))}

            {documentData.activities.length === 0 ? (
              <EmptyState actionLabel="Adicionar atividade" onAction={onAddActivity} />
            ) : null}
          </Stack>
        </Section>
      </Tabs.Panel>

      <Tabs.Panel value="review" pt="md">
        <TeaReviewPanel summary={reviewSummary} onIssueClick={onReviewIssueClick} />
      </Tabs.Panel>
    </Tabs>
  );
}

function TeaActivityEditor({
  index,
  activity,
  onChange,
  onRemove,
  onItemsChange,
  onAddSubActivity,
  onSubActivityChange,
  onSubActivityRemove,
  onSubActivityItemsChange,
}: {
  index: number;
  activity: TeaActivity;
  onChange: (updater: (activity: TeaActivity) => TeaActivity) => void;
  onRemove: () => void;
  onItemsChange: (value: string) => void;
  onAddSubActivity: () => void;
  onSubActivityChange: (
    subActivityId: string,
    updater: (subActivity: TeaSubActivity) => TeaSubActivity,
  ) => void;
  onSubActivityRemove: (subActivityId: string) => void;
  onSubActivityItemsChange: (subActivityId: string, value: string) => void;
}) {
  return (
    <Paper withBorder p="md" className="teaActivityCard">
      <Stack gap="sm">
        <Group justify="space-between" align="center">
          <Group gap="xs">
            <Badge variant="outline" color="gray">
              2.{index + 1}
            </Badge>
            <Text fw={800}>Atividade</Text>
          </Group>
          <Tooltip label="Remover atividade">
            <ActionIcon variant="subtle" color="red" onClick={onRemove} aria-label="Remover atividade">
              <Trash2 size={17} />
            </ActionIcon>
          </Tooltip>
        </Group>

        <TextInput
          label="Título"
          value={activity.title}
          placeholder="Seletor Situação do Prestador"
          onChange={(event) =>
            onChange((current) => ({ ...current, title: event.currentTarget.value }))
          }
        />
        <Textarea
          label="Descrição"
          minRows={4}
          autosize
          styles={{ input: { resize: "vertical" } }}
          value={activity.description}
          onChange={(event) =>
            onChange((current) => ({ ...current, description: event.currentTarget.value }))
          }
        />
        <Textarea
          label="Itens em lista"
          minRows={2}
          autosize
          placeholder="Um item por linha"
          value={activity.items.map((item) => item.text).join("\n")}
          onChange={(event) => onItemsChange(event.currentTarget.value)}
        />
        <EvidenceUploader
          title="Imagens da atividade"
          tone="new"
          images={activity.images}
          onChange={(updater) =>
            onChange((current) => ({ ...current, images: updater(current.images) }))
          }
        />

        <Divider />

        <Group justify="space-between" align="center">
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

        <Stack gap="sm">
          {activity.subActivities.map((subActivity, subIndex) => (
            <TeaSubActivityEditor
              key={subActivity.id}
              activityIndex={index}
              index={subIndex}
              subActivity={subActivity}
              onChange={(updater) => onSubActivityChange(subActivity.id, updater)}
              onRemove={() => onSubActivityRemove(subActivity.id)}
              onItemsChange={(value) => onSubActivityItemsChange(subActivity.id, value)}
            />
          ))}
        </Stack>
      </Stack>
    </Paper>
  );
}

function TeaSubActivityEditor({
  activityIndex,
  index,
  subActivity,
  onChange,
  onRemove,
  onItemsChange,
}: {
  activityIndex: number;
  index: number;
  subActivity: TeaSubActivity;
  onChange: (updater: (subActivity: TeaSubActivity) => TeaSubActivity) => void;
  onRemove: () => void;
  onItemsChange: (value: string) => void;
}) {
  return (
    <Paper withBorder p="sm" className="teaSubActivityCard">
      <Stack gap="sm">
        <Group justify="space-between" align="center">
          <Badge variant="outline" color="gray">
            2.{activityIndex + 1}.{index + 1}
          </Badge>
          <Tooltip label="Remover subtópico">
            <ActionIcon variant="subtle" color="red" onClick={onRemove} aria-label="Remover subtópico">
              <Trash2 size={17} />
            </ActionIcon>
          </Tooltip>
        </Group>

        <TextInput
          label="Título"
          value={subActivity.title}
          placeholder="Botão de Anexo"
          onChange={(event) =>
            onChange((current) => ({ ...current, title: event.currentTarget.value }))
          }
        />
        <Textarea
          label="Descrição"
          minRows={3}
          autosize
          styles={{ input: { resize: "vertical" } }}
          value={subActivity.description}
          onChange={(event) =>
            onChange((current) => ({ ...current, description: event.currentTarget.value }))
          }
        />
        <Textarea
          label="Itens em lista"
          minRows={2}
          autosize
          placeholder="Um item por linha"
          value={subActivity.items.map((item) => item.text).join("\n")}
          onChange={(event) => onItemsChange(event.currentTarget.value)}
        />
        <EvidenceUploader
          title="Imagens do subtópico"
          tone="legacy"
          images={subActivity.images}
          onChange={(updater) =>
            onChange((current) => ({ ...current, images: updater(current.images) }))
          }
        />
      </Stack>
    </Paper>
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

  return (
    <Section
      title="Revisão"
      tone="document"
      action={
        <Badge color={totalIssues > 0 ? "yellow" : "green"} variant="light">
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

        {summary.issues.length > 0 ? (
          <ReviewIssueGroup
            title="Pendências para revisar"
            tone="warning"
            issues={summary.issues}
            onIssueClick={onIssueClick}
          />
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

function Section({
  title,
  tone,
  action,
  children,
}: {
  title: string;
  tone: "document" | "steps" | "permissions" | "blocks";
  action?: ReactNode;
  children: ReactNode;
}) {
  return (
    <Card
      withBorder
      padding="lg"
      radius="md"
      className={`sectionCard sectionCard--${tone}`}
    >
      <Group justify="space-between" mb="md" align="center">
        <Title order={2} size="h4">
          {title}
        </Title>
        {action}
      </Group>
      {children}
    </Card>
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

function PermissionGroupEditor({
  index,
  macro,
  onMacroChange,
  onRemoveMacro,
  onAddMicro,
  onMicroChange,
  onRemoveMicro,
}: {
  index: number;
  macro: PermissionGroup;
  onMacroChange: (updates: Partial<PermissionItem>) => void;
  onRemoveMacro: () => void;
  onAddMicro: () => void;
  onMicroChange: (microId: string, updates: Partial<PermissionItem>) => void;
  onRemoveMicro: (microId: string) => void;
}) {
  return (
    <Paper withBorder p="md" className="permissionGroup">
      <Stack gap="sm">
        <Group justify="space-between" align="center">
          <Group gap="xs">
            <Badge variant="outline" color="gray">
              Macro {index + 1}
            </Badge>
            <Checkbox
              label="Usar"
              checked={macro.selected}
              onChange={(event) => onMacroChange({ selected: event.currentTarget.checked })}
            />
          </Group>
          <Tooltip label="Remover macro">
            <ActionIcon
              variant="subtle"
              color="red"
              onClick={onRemoveMacro}
              aria-label="Remover macro"
            >
              <Trash2 size={17} />
            </ActionIcon>
          </Tooltip>
        </Group>

        <div className="permissionFields">
          <TextInput
            label="Código"
            value={macro.code}
            placeholder="AO"
            onChange={(event) => onMacroChange({ code: event.currentTarget.value })}
          />
          <TextInput
            label="Descrição"
            value={macro.label}
            placeholder="Administrador Geral"
            onChange={(event) => onMacroChange({ label: event.currentTarget.value })}
          />
        </div>

        <Divider />

        <Group justify="space-between" align="center">
          <Text fw={700} size="sm">
            Micro-permissões
          </Text>
          <Button variant="subtle" size="xs" leftSection={<Plus size={15} />} onClick={onAddMicro}>
            Adicionar micro
          </Button>
        </Group>

        <Stack gap="xs">
          {macro.microPermissions.map((micro) => (
            <div className="microPermissionRow" key={micro.id}>
              <Checkbox
                label="Usar"
                checked={micro.selected}
                onChange={(event) =>
                  onMicroChange(micro.id, { selected: event.currentTarget.checked })
                }
              />
              <TextInput
                label="Código"
                value={micro.code}
                placeholder="AT"
                onChange={(event) =>
                  onMicroChange(micro.id, { code: event.currentTarget.value })
                }
              />
              <TextInput
                label="Descrição"
                value={micro.label}
                placeholder="Atualização"
                onChange={(event) =>
                  onMicroChange(micro.id, { label: event.currentTarget.value })
                }
              />
              <Tooltip label="Remover micro">
                <ActionIcon
                  variant="subtle"
                  color="red"
                  onClick={() => onRemoveMicro(micro.id)}
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
}

const PermissionBlockGroup = memo(function PermissionBlockGroup({
  macro,
  entries,
  expandedTests,
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
  return (
    <Paper withBorder p="md" className="blockGroup">
      <Stack gap="sm">
        <Group justify="space-between" align="center">
          <div>
            <Text size="xs" c="dimmed" fw={700} tt="uppercase">
              Macro
            </Text>
            <Title order={3} size="h5">
              {formatPermission(macro)}
            </Title>
          </div>
          <Badge variant="outline" color="gray">
            {entries.length} micro{entries.length === 1 ? "" : "s"}
          </Badge>
        </Group>

        <Stack gap="sm">
          {entries.map((entry) => (
            <PermissionBlockEditor
              key={entry.key}
              blockKey={entry.key}
              entry={entry}
              block={entry.block}
              sourceBlock={entry.sourceBlock}
              expandedTests={expandedTests}
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
  return (
    <Paper withBorder p="md" className="permissionBlock">
      <Stack gap="sm">
        <Group justify="space-between" align="center" gap="md">
          <div>
            <Group gap="xs" mb={4}>
              <Badge variant="outline" color="gray">
                Micro
              </Badge>
              <Text fw={700}>{formatPermission(entry.micro)}</Text>
            </Group>
            <Text c="dimmed" size="sm">
              {formatPermission(entry.macro)}
            </Text>
          </div>

          <Group gap="xs">
            <Button
              variant="subtle"
              size="xs"
              leftSection={<ListChecks size={15} />}
              onClick={() => onAddStandardTests(blockKey)}
            >
              Adicionar pacote padrão
            </Button>
            <Button
              variant="subtle"
              size="xs"
              leftSection={<Copy size={15} />}
              disabled={block.tests.length === 0}
              onClick={() => onDuplicateBlockStructure(blockKey)}
            >
              Copiar para vazios
            </Button>
            <Button
              variant="subtle"
              size="xs"
              leftSection={<Plus size={15} />}
              onClick={() => onAddTest(blockKey)}
            >
              Adicionar teste
            </Button>
          </Group>
        </Group>

        <Stack gap="sm">
          {block.tests.map((test, visibleIndex) => {
            const selfReferenceKey = createTestReferenceKey(entry.key, test.id);
            const sourceIndex = sourceBlock.tests.findIndex((sourceTest) => sourceTest.id === test.id);
            const testIndex = sourceIndex >= 0 ? sourceIndex : visibleIndex;

            return (
              <BlockTestEditor
                key={test.id}
                blockKey={blockKey}
                index={testIndex}
                test={test}
                selfReferenceKey={selfReferenceKey}
                isExpanded={expandedTests[selfReferenceKey] ?? false}
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
            <EmptyState actionLabel="Adicionar teste" onAction={() => onAddTest(blockKey)} />
          ) : null}
        </Stack>
      </Stack>
    </Paper>
  );
}, arePermissionBlockEditorPropsEqual);

const BlockTestEditor = memo(function BlockTestEditor({
  blockKey,
  index,
  test,
  selfReferenceKey,
  isExpanded,
  canMoveUp,
  canMoveDown,
  onTestExpansionChange,
  onTestTitleChange,
  onTestMove,
  onDuplicateTest,
  onTestRemove,
  onResultChange,
}: BlockTestEditorProps) {
  const selectedCheckCount = checkOrder.filter((key) => test.result.checks[key]).length;
  const evidenceCount = test.result.legacyImages.length + test.result.newImages.length;
  const testPanelId = `test-details-${toDomId(selfReferenceKey)}`;
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
        <div className="testHeaderGrid">
          <Group gap={6} wrap="nowrap" className="testIndexCell">
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
          </Group>
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
          <Group gap={4} wrap="wrap" className="testActions">
            <Badge color="gray" variant="outline" className="testMetaBadge">
              {selectedCheckCount}/{checkOrder.length} checks
            </Badge>
            {evidenceCount > 0 ? (
              <Badge color="gray" variant="outline" className="testMetaBadge">
                {evidenceCount} img
              </Badge>
            ) : null}
            <Tooltip label="Duplicar teste">
              <ActionIcon
                variant="subtle"
                onClick={() => onDuplicateTest(blockKey, test.id)}
                aria-label="Duplicar teste"
              >
                <Copy size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Mover para cima">
              <ActionIcon
                variant="subtle"
                disabled={!canMoveUp}
                onClick={() => onTestMove(blockKey, index, -1)}
                aria-label="Mover para cima"
              >
                <ArrowUp size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Mover para baixo">
              <ActionIcon
                variant="subtle"
                disabled={!canMoveDown}
                onClick={() => onTestMove(blockKey, index, 1)}
                aria-label="Mover para baixo"
              >
                <ArrowDown size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Remover teste">
              <ActionIcon
                variant="subtle"
                color="red"
                onClick={() => onTestRemove(blockKey, test.id)}
                aria-label="Remover teste"
              >
                <Trash2 size={17} />
              </ActionIcon>
            </Tooltip>
          </Group>
        </div>

        <Group gap={6} className="quickChecks">
          {checkOrder.map((key) => (
            <Tooltip key={key} label={checkLabels[key]}>
              <button
                type="button"
                className={`quickCheck ${test.result.checks[key] ? "quickCheck--active" : ""}`}
                onClick={() => toggleCheck(key)}
              >
                {quickCheckLabels[key]}
              </button>
            </Tooltip>
          ))}
        </Group>

        {isExpanded ? (
        <Collapse in transitionDuration={0}>
          <div id={testPanelId} className="testBody">
            <TestResultEditor
              result={test.result}
              onChange={(updater) => onResultChange(blockKey, test.id, updater)}
            />
          </div>
        </Collapse>
        ) : null}
      </Stack>
    </Paper>
  );
});

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

  function updateCheck(key: CheckKey): void {
    onChange((current) => ({
      ...current,
      checks: {
        ...current.checks,
        [key]: !current.checks[key],
      },
    }));
  }

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
      <Stack gap="xs">
        {checkOrder.map((key) => (
          <Checkbox
            key={key}
            checked={result.checks[key]}
            onChange={() => updateCheck(key)}
            label={checkLabels[key]}
          />
        ))}
      </Stack>

      <Textarea
        label="Observações"
        minRows={4}
        styles={{ input: { resize: "vertical" } }}
        value={observations.value}
        onBlur={observations.commit}
        onChange={(event) => {
          const value = event.currentTarget.value;
          observations.setValue(value);
        }}
      />

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
    if (!fileList.length) {
      return;
    }

    const evidence = await Promise.all(
      fileList
        .filter(isImageFile)
        .map(async (file) => {
          const optimized = await optimizeImageFile(file);
          const id = createId();

          try {
            await saveEvidenceImageData(id, optimized.dataUrl);
          } catch {
            window.alert("Nao foi possivel salvar a imagem no rascunho do navegador.");
          }

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

    onChange((current) => [...current, ...evidence]);
  }

  function updateImageLabel(imageId: string, label: string): void {
    onChange((current) =>
      current.map((image) => (image.id === imageId ? { ...image, label } : image)),
    );
  }

  function removeImage(imageId: string): void {
    void deleteEvidenceImageData(imageId);
    onChange((current) => current.filter((image) => image.id !== imageId));
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
      aria-label={`${title}: cole ou arraste uma imagem`}
    >
      <Stack gap="sm">
        <Group justify="space-between">
          <Group gap="xs">
            <Text fw={700}>{title}</Text>
            <Tooltip label="Aceita imagem colada ou arrastada">
              <ClipboardPaste size={16} className="pasteIndicator" aria-hidden="true" />
            </Tooltip>
          </Group>
          <FileButton
            onChange={(files) => {
              void addFiles(files);
            }}
            accept="image/*"
            multiple
          >
            {(props) => (
              <Button {...props} variant="light" leftSection={<ImagePlus size={17} />}>
                Imagem
              </Button>
            )}
          </FileButton>
        </Group>

        {images.length > 0 ? (
          <Stack gap="xs">
            {images.map((image) => (
              <Paper withBorder p="xs" key={image.id}>
                <Group align="center" wrap="nowrap">
                  {image.dataUrl ? (
                    <img className="imagePreview" src={image.dataUrl} alt={image.name} />
                  ) : (
                    <div className="imagePreview imagePreview--missing">Sem preview</div>
                  )}
                  <Stack gap={4} style={{ flex: 1, minWidth: 0 }}>
                    <TextInput
                      value={image.label}
                      placeholder="Legenda"
                      onChange={(event) =>
                        updateImageLabel(image.id, event.currentTarget.value)
                      }
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
                      onClick={() => removeImage(image.id)}
                      aria-label="Remover imagem"
                    >
                      <X size={17} />
                    </ActionIcon>
                  </Tooltip>
                </Group>
              </Paper>
            ))}
          </Stack>
        ) : (
          <Text c="dimmed" ta="center" py="md">
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
}: {
  actionLabel: string;
  onAction: () => void;
}) {
  return (
    <Paper withBorder p="md" ta="center" className="softEmpty">
      <Stack gap="xs" align="center">
        <Text c="dimmed">Nenhum item adicionado.</Text>
        <Button variant="light" size="xs" leftSection={<Plus size={15} />} onClick={onAction}>
          {actionLabel}
        </Button>
      </Stack>
    </Paper>
  );
}

function ImportPreviewModal({
  result,
  onClose,
  onConfirm,
}: {
  result: DocxImportResult | null;
  onClose: () => void;
  onConfirm: () => void;
}) {
  return (
    <Modal
      opened={result !== null}
      onClose={onClose}
      title="Prévia da importação"
      size="lg"
      centered
    >
      {result ? (
        <Stack gap="md">
          <div>
            <Text fw={800}>{result.summary.screen || result.sourceName}</Text>
            <Text size="sm" c="dimmed">
              {result.sourceName}
            </Text>
          </div>

          <div className="importMetrics">
            <ReviewMetric label="Passos" value={result.summary.accessSteps} />
            <ReviewMetric label="Permissões" value={result.summary.selectedPermissions} />
            <ReviewMetric label="Testes" value={result.summary.tests} />
            <ReviewMetric label="Imagens" value={result.summary.images} />
          </div>

          {result.warnings.length > 0 ? (
            <div className="importWarnings">
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
            <Button variant="light" color="gray" onClick={onClose}>
              Cancelar
            </Button>
            <Button leftSection={<FileUp size={17} />} onClick={onConfirm}>
              Substituir rascunho
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

  return (
    <Section
      title="Revisão"
      tone="document"
      action={
        <Badge
          color={totalIssues > 0 ? "yellow" : "green"}
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
          {summary.issues.length > 0 ? (
            <ReviewIssueGroup
              title="Pendências para revisar"
              tone="warning"
              issues={summary.issues}
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

function ReviewIssueGroup<TIssue extends { id: string; label: string; detail: string }>({
  title,
  tone,
  issues,
  onIssueClick,
}: {
  title: string;
  tone: "warning";
  issues: TIssue[];
  onIssueClick: (issue: TIssue) => void;
}) {
  return (
    <div className={`reviewIssueGroup reviewIssueGroup--${tone}`}>
      <Group gap="xs" mb={6}>
        <AlertCircle size={17} />
        <Text fw={700}>{title}</Text>
      </Group>
      <Stack gap={4}>
        {issues.map((issue) => (
          <button
            key={issue.id}
            type="button"
            className="reviewIssueButton"
            onClick={() => onIssueClick(issue)}
          >
            <Text size="sm" fw={700}>
              {issue.label}
            </Text>
            <Text size="xs" c="dimmed">
              {issue.detail}
            </Text>
          </button>
        ))}
      </Stack>
    </div>
  );
}

function useBufferedText(
  externalValue: string,
  onCommit: (value: string) => void,
  delay = 120,
): BufferedTextState {
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

  return { value, setValue, commit };
}

function arePermissionBlockGroupPropsEqual(
  previous: PermissionBlockGroupProps,
  next: PermissionBlockGroupProps,
): boolean {
  if (
    previous.macro !== next.macro ||
    previous.entries !== next.entries ||
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
    return !test.title.trim();
  }

  return problemCheckKeys.some((key) => test.result.checks[key]);
}

function getTestImageCount(test: PermissionBlockTest): number {
  return test.result.legacyImages.length + test.result.newImages.length;
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
      label: "Tela sem nome",
      detail: "Preencha o campo Tela antes de exportar.",
      tab: "document",
    });
  }

  if (!documentData.objective.trim()) {
    summary.issues.push({
      id: "missing-objective",
      label: "Objetivo vazio",
      detail: "Preencha o objetivo do documento.",
      tab: "document",
    });
  }

  if (documentData.accessSteps.filter((step) => step.text.trim()).length === 0) {
    summary.issues.push({
      id: "missing-steps",
      label: "Sem passo a passo",
      detail: "Inclua ao menos uma etapa de acesso.",
      tab: "document",
    });
  }

  if (entries.length === 0) {
    summary.issues.push({
      id: "missing-permissions",
      label: "Sem permissões selecionadas",
      detail: "Selecione macro e micro-permissões para gerar testes.",
      tab: "permissions",
    });
  }

  entries.forEach((entry) => {
    const block = documentData.permissionBlocks[entry.key] ?? createEmptyBlock();

    block.tests.forEach((test, index) => {
      const testLabel = `${formatPermission(entry.macro)} / ${formatPermission(entry.micro)} / ${
        test.title.trim() || `Teste ${index + 1}`
      }`;

      summary.testCount += 1;
      summary.imageCount += test.result.legacyImages.length + test.result.newImages.length;

      if (!test.title.trim()) {
        const referenceKey = createTestReferenceKey(entry.key, test.id);
        summary.issues.push({
          id: `unnamed-${referenceKey}`,
          label: "Teste sem nome",
          detail: testLabel,
          tab: "tests",
          targetId: `test-card-${toDomId(referenceKey)}`,
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
          activity.images.length +
          activity.subActivities.reduce(
            (subTotal, subActivity) => subTotal + subActivity.images.length,
            0,
          ),
        0,
      ),
    issues: [],
  };

  if (!documentData.metadata.serviceOrder.trim()) {
    summary.issues.push({
      id: "missing-service-order",
      label: "Ordem de serviço vazia",
      detail: "Preencha a ordem de serviço do TEA.",
      tab: "document",
    });
  }

  if (!documentData.metadata.subject.trim()) {
    summary.issues.push({
      id: "missing-subject",
      label: "Assunto vazio",
      detail: "Preencha o assunto usado no cabeçalho e nome do arquivo.",
      tab: "document",
    });
  }

  if (!documentData.overview.trim()) {
    summary.issues.push({
      id: "missing-overview",
      label: "Visão geral vazia",
      detail: "Preencha a seção 1 do documento.",
      tab: "document",
    });
  }

  if (documentData.activities.length === 0) {
    summary.issues.push({
      id: "missing-activities",
      label: "Sem atividades",
      detail: "Adicione ao menos uma atividade realizada.",
      tab: "activities",
    });
  }

  documentData.activities.forEach((activity, index) => {
    if (!activity.title.trim()) {
      summary.issues.push({
        id: `activity-title-${activity.id}`,
        label: "Atividade sem título",
        detail: `Atividade 2.${index + 1}`,
        tab: "activities",
      });
    }
  });

  return summary;
}

function createBlockTest(id: string, title = ""): PermissionBlockTest {
  return {
    id,
    title,
    result: createEmptyTestResult(),
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
