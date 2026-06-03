import { useEffect, useMemo, useState } from "react";
import type { ReactNode } from "react";
import {
  ActionIcon,
  Badge,
  Button,
  Card,
  Checkbox,
  Container,
  Divider,
  FileButton,
  Group,
  Paper,
  SegmentedControl,
  Select,
  Stack,
  Text,
  Textarea,
  TextInput,
  Title,
  Tooltip,
} from "@mantine/core";
import {
  ArrowDown,
  ArrowUp,
  Download,
  ImagePlus,
  Plus,
  RotateCcw,
  Save,
  Trash2,
  X,
} from "lucide-react";
import { checkLabels, checkOrder, createEmptyTestResult, createPermissionKey } from "./defaultDocument";
import { exportOtDocument } from "./docxExport";
import { clearDraft, loadDraft, saveDraft } from "./storage";
import type {
  CheckKey,
  EvidenceImage,
  OtDocument,
  PermissionBlock,
  PermissionTestMode,
  PermissionBlockTest,
  PermissionGroup,
  PermissionItem,
  TestResult,
} from "./types";

type PermissionBlockEntry = {
  key: string;
  macro: PermissionGroup;
  micro: PermissionItem;
};

const testModeOptions: { value: PermissionTestMode; label: string }[] = [
  { value: "test", label: "Testar" },
  { value: "idem", label: "IDEM" },
];

export default function App() {
  const [documentData, setDocumentData] = useState<OtDocument>(() => loadDraft());
  const [isExporting, setIsExporting] = useState(false);
  const [draftStatus, setDraftStatus] = useState("Rascunho salvo");

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

  const referenceOptions = useMemo(
    () =>
      permissionBlockEntries.flatMap((entry) => {
        const block = documentData.permissionBlocks[entry.key] ?? createEmptyBlock();

        return block.tests.map((test, index) => ({
          value: createTestReferenceKey(entry.key, test.id),
          label: `${formatPermission(entry.macro)} / ${formatPermission(entry.micro)} / ${
            test.title.trim() || `Teste ${index + 1}`
          }`,
        }));
      }),
    [documentData.permissionBlocks, permissionBlockEntries],
  );

  useEffect(() => {
    try {
      saveDraft(documentData);
      setDraftStatus("Rascunho salvo");
    } catch {
      setDraftStatus("Rascunho grande demais");
    }
  }, [documentData]);

  function updateDocument(updater: (current: OtDocument) => OtDocument): void {
    setDocumentData((current) => updater(current));
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

  function addBlockTest(blockKey: string): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: [
        ...block.tests,
        {
          id: createId(),
          title: "",
          mode: "test",
          result: createEmptyTestResult(),
        },
      ],
    }));
  }

  function updateBlockTestTitle(blockKey: string, testId: string, title: string): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId ? { ...test, title } : test,
      ),
    }));
  }

  function removeBlockTest(blockKey: string, testId: string): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.filter((test) => test.id !== testId),
    }));
  }

  function moveBlockTest(blockKey: string, index: number, direction: -1 | 1): void {
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
  }

  function updateBlockTestMode(
    blockKey: string,
    testId: string,
    mode: PermissionTestMode,
  ): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId
          ? {
              ...test,
              mode,
              idemReferenceKey:
                mode === "idem" &&
                test.idemReferenceKey !== createTestReferenceKey(blockKey, testId)
                  ? test.idemReferenceKey
                  : undefined,
            }
          : test,
      ),
    }));
  }

  function updateTestIdemReference(
    blockKey: string,
    testId: string,
    referenceKey: string | null,
  ): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId
          ? {
              ...test,
              idemReferenceKey:
                referenceKey && referenceKey !== createTestReferenceKey(blockKey, testId)
                  ? referenceKey
                  : undefined,
            }
          : test,
      ),
    }));
  }

  function updateTestResult(
    blockKey: string,
    testId: string,
    updater: (result: TestResult) => TestResult,
  ): void {
    updateBlock(blockKey, (block) => ({
      ...block,
      tests: block.tests.map((test) =>
        test.id === testId ? { ...test, result: updater(test.result) } : test,
      ),
    }));
  }

  function updateBlock(
    blockKey: string,
    updater: (block: PermissionBlock) => PermissionBlock,
  ): void {
    updateDocument((current) => ({
      ...current,
      permissionBlocks: {
        ...current.permissionBlocks,
        [blockKey]: updater(current.permissionBlocks[blockKey] ?? createEmptyBlock()),
      },
    }));
  }

  async function handleExport(): Promise<void> {
    const invalidIdem = permissionBlockEntries.some((entry) => {
      const block = documentData.permissionBlocks[entry.key];
      return block?.tests.some(
        (test) =>
          test.mode === "idem" &&
          !isValidTestReference(
            createTestReferenceKey(entry.key, test.id),
            test.idemReferenceKey,
            referenceOptions,
          ),
      );
    });

    if (invalidIdem) {
      window.alert("Selecione a referência de todos os testes marcados como IDEM.");
      return;
    }

    setIsExporting(true);

    try {
      await exportOtDocument(documentData);
    } finally {
      setIsExporting(false);
    }
  }

  function handleClearDraft(): void {
    if (!window.confirm("Limpar o rascunho atual?")) {
      return;
    }

    clearDraft();
    setDocumentData(loadDraft());
  }

  return (
    <Container size="xl" py="xl">
      <Stack gap="lg">
        <Paper withBorder shadow="sm" p="lg" className="topBar">
          <Group justify="space-between" align="flex-start" gap="md">
            <div>
              <Title order={1} size="h2">
                Gerador de OT
              </Title>
              <Text c="dimmed" mt={4}>
                {documentData.metadata.screen || "Documento sem tela definida"}
              </Text>
            </div>

            <Group gap="xs">
              <Badge
                variant="light"
                color={draftStatus.includes("grande") ? "red" : "green"}
                leftSection={<Save size={14} />}
                h={30}
              >
                {draftStatus}
              </Badge>
              <Button
                variant="light"
                color="gray"
                leftSection={<RotateCcw size={17} />}
                onClick={handleClearDraft}
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

        <Section title="Documento">
          <Stack gap="sm">
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
            <Textarea
              label="Objetivo"
              minRows={4}
              autosize
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
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={addStep}>
              Adicionar
            </Button>
          }
        >
          <Stack gap="xs">
            {documentData.accessSteps.map((step, index) => (
              <Group key={step.id} align="flex-end" wrap="nowrap">
                <Badge color="blue" variant="light" w={34} h={34}>
                  {index + 1}
                </Badge>
                <TextInput
                  label={index === 0 ? "Etapa" : undefined}
                  value={step.text}
                  onChange={(event) => updateStep(step.id, event.currentTarget.value)}
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

        <Section
          title="Permissões"
          action={
            <Button variant="light" leftSection={<Plus size={17} />} onClick={addMacroGroup}>
              Adicionar macro
            </Button>
          }
        >
          <Stack gap="sm">
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

        <Section title="Blocos de permissão">
          <Stack gap="md">
            {selectedGroups.map((macro) => (
              <PermissionBlockGroup
                key={macro.id}
                macro={macro}
                entries={macro.microPermissions.map((micro) => ({
                  key: createPermissionKey(macro.id, micro.id),
                  macro,
                  micro,
                }))}
                blocks={documentData.permissionBlocks}
                referenceOptions={referenceOptions}
                onAddTest={addBlockTest}
                onTestTitleChange={updateBlockTestTitle}
                onTestModeChange={updateBlockTestMode}
                onTestReferenceChange={updateTestIdemReference}
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
          </Stack>
        </Section>
      </Stack>
    </Container>
  );
}

function Section({
  title,
  action,
  children,
}: {
  title: string;
  action?: ReactNode;
  children: ReactNode;
}) {
  return (
    <Card withBorder shadow="xs" padding="lg" radius="md" className="sectionCard">
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
            <Badge variant="light" color="indigo">
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

function PermissionBlockGroup({
  macro,
  entries,
  blocks,
  referenceOptions,
  onAddTest,
  onTestTitleChange,
  onTestModeChange,
  onTestReferenceChange,
  onTestRemove,
  onTestMove,
  onResultChange,
}: {
  macro: PermissionGroup;
  entries: PermissionBlockEntry[];
  blocks: Record<string, PermissionBlock>;
  referenceOptions: { value: string; label: string }[];
  onAddTest: (blockKey: string) => void;
  onTestTitleChange: (blockKey: string, testId: string, title: string) => void;
  onTestModeChange: (
    blockKey: string,
    testId: string,
    mode: PermissionTestMode,
  ) => void;
  onTestReferenceChange: (
    blockKey: string,
    testId: string,
    referenceKey: string | null,
  ) => void;
  onTestRemove: (blockKey: string, testId: string) => void;
  onTestMove: (blockKey: string, index: number, direction: -1 | 1) => void;
  onResultChange: (
    blockKey: string,
    testId: string,
    updater: (result: TestResult) => TestResult,
  ) => void;
}) {
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
          <Badge variant="light" color="indigo">
            {entries.length} micro{entries.length === 1 ? "" : "s"}
          </Badge>
        </Group>

        <Stack gap="sm">
          {entries.map((entry) => (
            <PermissionBlockEditor
              key={entry.key}
              entry={entry}
              block={blocks[entry.key] ?? createEmptyBlock()}
              referenceOptions={referenceOptions}
              onAddTest={() => onAddTest(entry.key)}
              onTestTitleChange={(testId, title) =>
                onTestTitleChange(entry.key, testId, title)
              }
              onTestModeChange={(testId, mode) =>
                onTestModeChange(entry.key, testId, mode)
              }
              onTestReferenceChange={(testId, referenceKey) =>
                onTestReferenceChange(entry.key, testId, referenceKey)
              }
              onTestRemove={(testId) => onTestRemove(entry.key, testId)}
              onTestMove={(index, direction) => onTestMove(entry.key, index, direction)}
              onResultChange={(testId, updater) =>
                onResultChange(entry.key, testId, updater)
              }
            />
          ))}
        </Stack>
      </Stack>
    </Paper>
  );
}

function PermissionBlockEditor({
  entry,
  block,
  referenceOptions,
  onAddTest,
  onTestTitleChange,
  onTestModeChange,
  onTestReferenceChange,
  onTestRemove,
  onTestMove,
  onResultChange,
}: {
  entry: PermissionBlockEntry;
  block: PermissionBlock;
  referenceOptions: { value: string; label: string }[];
  onAddTest: () => void;
  onTestTitleChange: (testId: string, title: string) => void;
  onTestModeChange: (testId: string, mode: PermissionTestMode) => void;
  onTestReferenceChange: (testId: string, referenceKey: string | null) => void;
  onTestRemove: (testId: string) => void;
  onTestMove: (index: number, direction: -1 | 1) => void;
  onResultChange: (testId: string, updater: (result: TestResult) => TestResult) => void;
}) {
  return (
    <Paper withBorder p="md" className="permissionBlock">
      <Stack gap="sm">
        <Group justify="space-between" align="center" gap="md">
          <div>
            <Group gap="xs" mb={4}>
              <Badge variant="light" color="cyan">
                Micro
              </Badge>
              <Text fw={700}>{formatPermission(entry.micro)}</Text>
            </Group>
            <Text c="dimmed" size="sm">
              {formatPermission(entry.macro)}
            </Text>
          </div>

          <Button
            variant="subtle"
            size="xs"
            leftSection={<Plus size={15} />}
            onClick={onAddTest}
          >
            Adicionar teste
          </Button>
        </Group>

        <Stack gap="sm">
          {block.tests.map((test, index) => {
            const selfReferenceKey = createTestReferenceKey(entry.key, test.id);

            return (
              <BlockTestEditor
                key={test.id}
                index={index}
                test={test}
                referenceOptions={referenceOptions.filter(
                  (option) => option.value !== selfReferenceKey,
                )}
                selfReferenceKey={selfReferenceKey}
                canMoveUp={index > 0}
                canMoveDown={index < block.tests.length - 1}
                onTitleChange={(title) => onTestTitleChange(test.id, title)}
                onModeChange={(mode) => onTestModeChange(test.id, mode)}
                onReferenceChange={(referenceKey) =>
                  onTestReferenceChange(test.id, referenceKey)
                }
                onMoveUp={() => onTestMove(index, -1)}
                onMoveDown={() => onTestMove(index, 1)}
                onRemove={() => onTestRemove(test.id)}
                onResultChange={(updater) => onResultChange(test.id, updater)}
              />
            );
          })}

          {block.tests.length === 0 ? (
            <EmptyState actionLabel="Adicionar teste" onAction={onAddTest} />
          ) : null}
        </Stack>
      </Stack>
    </Paper>
  );
}

function BlockTestEditor({
  index,
  test,
  referenceOptions,
  selfReferenceKey,
  canMoveUp,
  canMoveDown,
  onTitleChange,
  onModeChange,
  onReferenceChange,
  onMoveUp,
  onMoveDown,
  onRemove,
  onResultChange,
}: {
  index: number;
  test: PermissionBlockTest;
  referenceOptions: { value: string; label: string }[];
  selfReferenceKey: string;
  canMoveUp: boolean;
  canMoveDown: boolean;
  onTitleChange: (title: string) => void;
  onModeChange: (mode: PermissionTestMode) => void;
  onReferenceChange: (referenceKey: string | null) => void;
  onMoveUp: () => void;
  onMoveDown: () => void;
  onRemove: () => void;
  onResultChange: (updater: (result: TestResult) => TestResult) => void;
}) {
  const hasValidReference = isValidTestReference(
    selfReferenceKey,
    test.idemReferenceKey,
    referenceOptions,
  );

  return (
    <Paper withBorder p="md" className="testCard">
      <Stack gap="sm">
        <div className="testHeaderGrid">
          <Badge color="green" variant="light" w={34} h={34}>
            {index + 1}
          </Badge>
          <TextInput
            label="Nome do teste"
            value={test.title}
            placeholder="Criação, edição, consulta..."
            onChange={(event) => onTitleChange(event.currentTarget.value)}
          />
          <Stack gap={4} className="testModeControl">
            <Text size="xs" c="dimmed" fw={700}>
              Modo do teste
            </Text>
            <SegmentedControl
              size="xs"
              data={testModeOptions}
              value={test.mode}
              onChange={(value) => onModeChange(value as PermissionTestMode)}
            />
          </Stack>
          <Group gap={4} wrap="nowrap" className="testActions">
            <Tooltip label="Mover para cima">
              <ActionIcon
                variant="subtle"
                disabled={!canMoveUp}
                onClick={onMoveUp}
                aria-label="Mover para cima"
              >
                <ArrowUp size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Mover para baixo">
              <ActionIcon
                variant="subtle"
                disabled={!canMoveDown}
                onClick={onMoveDown}
                aria-label="Mover para baixo"
              >
                <ArrowDown size={17} />
              </ActionIcon>
            </Tooltip>
            <Tooltip label="Remover teste">
              <ActionIcon
                variant="subtle"
                color="red"
                onClick={onRemove}
                aria-label="Remover teste"
              >
                <Trash2 size={17} />
              </ActionIcon>
            </Tooltip>
          </Group>
        </div>

        {test.mode === "idem" ? (
          <Select
            label="IDEM ao teste"
            placeholder="Selecione macro, micro e teste de referência"
            data={referenceOptions}
            value={hasValidReference ? test.idemReferenceKey ?? null : null}
            onChange={onReferenceChange}
            disabled={referenceOptions.length === 0}
            error={!hasValidReference ? "Selecione um teste de referência" : undefined}
          />
        ) : (
          <TestResultEditor result={test.result} onChange={onResultChange} />
        )}
      </Stack>
    </Paper>
  );
}

function TestResultEditor({
  result,
  onChange,
}: {
  result: TestResult;
  onChange: (updater: (result: TestResult) => TestResult) => void;
}) {
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
        autosize
        value={result.observations}
        onChange={(event) => {
          const value = event.currentTarget.value;

          onChange((current) => ({ ...current, observations: value }));
        }}
      />

      <EvidenceUploader
        title="Legado"
        images={result.legacyImages}
        onChange={(updater) => updateImages("legacyImages", updater)}
      />
      <EvidenceUploader
        title="Novo"
        images={result.newImages}
        onChange={(updater) => updateImages("newImages", updater)}
      />
    </Stack>
  );
}

function EvidenceUploader({
  title,
  images,
  onChange,
}: {
  title: string;
  images: EvidenceImage[];
  onChange: (updater: (images: EvidenceImage[]) => EvidenceImage[]) => void;
}) {
  async function addFiles(files: File[] | File | null): Promise<void> {
    const fileList = Array.isArray(files) ? files : files ? [files] : [];
    if (!fileList.length) {
      return;
    }

    const evidence = await Promise.all(
      fileList
        .filter((file) => file.type.startsWith("image/"))
        .map(async (file) => {
          const dataUrl = await fileToDataUrl(file);
          const size = await readImageSize(dataUrl);

          return {
            id: createId(),
            label: "",
            name: file.name,
            dataUrl,
            width: size.width,
            height: size.height,
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
    onChange((current) => current.filter((image) => image.id !== imageId));
  }

  return (
    <Paper withBorder p="sm" className="evidencePanel">
      <Stack gap="sm">
        <Group justify="space-between">
          <Text fw={700}>{title}</Text>
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
                  <img className="imagePreview" src={image.dataUrl} alt={image.name} />
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
}

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

function isValidTestReference(
  selfReferenceKey: string,
  referenceKey: string | undefined,
  referenceOptions: { value: string; label: string }[],
): boolean {
  return (
    !!referenceKey &&
    referenceKey !== selfReferenceKey &&
    referenceOptions.some((option) => option.value === referenceKey)
  );
}

function createTestReferenceKey(blockKey: string, testId: string): string {
  return `${blockKey}::${testId}`;
}

function removePermissionBlocks(
  blocks: Record<string, PermissionBlock>,
  shouldRemove: (key: string) => boolean,
): Record<string, PermissionBlock> {
  const removedKeys = new Set(Object.keys(blocks).filter(shouldRemove));

  return Object.fromEntries(
    Object.entries(blocks)
      .filter(([key]) => !removedKeys.has(key))
      .map(([key, block]) => [
        key,
        {
          ...block,
          tests: block.tests.map((test) => {
            const pointsToRemovedBlock = Array.from(removedKeys).some((removedKey) =>
              test.idemReferenceKey?.startsWith(`${removedKey}::`),
            );

            return pointsToRemovedBlock ? { ...test, idemReferenceKey: undefined } : test;
          }),
        },
      ]),
  );
}

function createId(): string {
  return window.crypto.randomUUID();
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function readImageSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
    image.onerror = () => resolve({ width: 560, height: 320 });
    image.src = dataUrl;
  });
}
