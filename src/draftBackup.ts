import JSZip from "jszip";
import {
  hydrateDocumentImages,
  hydrateTeaDocumentImages,
  stripImageDataFromDocument,
  stripImageDataFromTeaDocument,
} from "./imageStorage";
import type {
  EvidenceImage,
  OtDocument,
  PermissionBlock,
  PermissionBlockTest,
  TeaContentBlock,
  TeaDocument,
  TestResult,
} from "./types";

export type DraftBackupKind = "ot" | "tea";

export type ParsedBackupFile =
  | {
      kind: "ot";
      document: OtDocument;
      warnings: string[];
    }
  | {
      kind: "tea";
      document: TeaDocument;
      warnings: string[];
    };

export type ClearImagesResult<TDocument> = {
  document: TDocument;
  imageIds: string[];
};

type BackupImageEntry = {
  id: string;
  name: string;
  label: string;
  width: number;
  height: number;
  originalBytes?: number;
  savedBytes?: number;
  optimized?: boolean;
  path?: string;
  mimeType?: string;
  missing?: boolean;
  warning?: string;
};

type DraftBackupPayload = {
  format: "create-ot-backup";
  version: 1;
  kind: DraftBackupKind;
  exportedAt: string;
  document: OtDocument | TeaDocument;
  images: BackupImageEntry[];
  warnings: string[];
};

type LocatedImage = {
  image: EvidenceImage;
  location: string;
};

const BACKUP_FORMAT = "create-ot-backup";
const BACKUP_VERSION = 1;

export async function exportOtBackup(documentData: OtDocument): Promise<void> {
  const hydratedDocument = await hydrateDocumentImages(documentData);
  await downloadBackup(
    await createBackupBlob("ot", hydratedDocument),
    createBackupFileName("OT", hydratedDocument.metadata.screen || "Documento"),
  );
}

export async function exportTeaBackup(documentData: TeaDocument): Promise<void> {
  const hydratedDocument = await hydrateTeaDocumentImages(documentData);
  await downloadBackup(
    await createBackupBlob(
      "tea",
      hydratedDocument,
    ),
    createBackupFileName(
      "TEA",
      hydratedDocument.metadata.subject ||
        hydratedDocument.metadata.serviceOrder ||
        "Documento",
    ),
  );
}

export async function parseBackupFile(file: File): Promise<ParsedBackupFile> {
  let zip: JSZip;

  try {
    zip = await JSZip.loadAsync(await file.arrayBuffer());
  } catch {
    throw new Error("Backup invalido. Selecione um arquivo .zip gerado pelo sistema.");
  }

  const draftFile = zip.file("draft.json");

  if (!draftFile) {
    throw new Error("Backup invalido: draft.json nao encontrado.");
  }

  let draftPayload: unknown;

  try {
    draftPayload = JSON.parse(await draftFile.async("string"));
  } catch {
    throw new Error("Backup invalido: draft.json nao pode ser lido.");
  }

  const payload = validateBackupPayload(draftPayload);
  const warnings = [...payload.warnings];
  const imageDataById = new Map<string, string>();

  for (const image of payload.images) {
    if (image.missing || !image.path) {
      continue;
    }

    const imageFile = zip.file(image.path);

    if (!imageFile) {
      warnings.push(`${image.name || image.id}: arquivo ${image.path} nao encontrado no backup.`);
      continue;
    }

    const bytes = await imageFile.async("uint8array");
    imageDataById.set(
      image.id,
      `data:${image.mimeType || "application/octet-stream"};base64,${bytesToBase64(bytes)}`,
    );
  }

  if (payload.kind === "tea") {
    return {
      kind: "tea",
      document: attachTeaImageData(payload.document as TeaDocument, imageDataById),
      warnings,
    };
  }

  return {
    kind: "ot",
    document: attachOtImageData(payload.document as OtDocument, imageDataById),
    warnings,
  };
}

export function removeAllOtImages(documentData: OtDocument): ClearImagesResult<OtDocument> {
  const imageIds = collectOtImages(documentData).map(({ image }) => image.id);
  const clearResult = (result: TestResult): TestResult => ({
    ...result,
    legacyImages: [],
    newImages: [],
  });

  return {
    imageIds,
    document: {
      ...documentData,
      permissionBlocks: Object.fromEntries(
        Object.entries(documentData.permissionBlocks).map(([blockKey, block]) => [
          blockKey,
          {
            ...block,
            tests: block.tests.map((test) => ({
              ...test,
              result: clearResult(test.result),
              correction: test.correction
                ? {
                    ...test.correction,
                    beforeImages: [],
                    afterImages: [],
                  }
                : test.correction,
            })),
          },
        ]),
      ),
    },
  };
}

export function removeAllTeaImages(documentData: TeaDocument): ClearImagesResult<TeaDocument> {
  const imageIds = collectTeaImages(documentData).map(({ image }) => image.id);
  const clearBlock = (block: TeaContentBlock): TeaContentBlock =>
    block.type === "images" ? { ...block, images: [] } : block;

  return {
    imageIds,
    document: {
      ...documentData,
      activityImages: [],
      activities: documentData.activities.map((activity) => ({
        ...activity,
        blocks: activity.blocks.map(clearBlock),
        subActivities: activity.subActivities.map((subActivity) => ({
          ...subActivity,
          blocks: subActivity.blocks.map(clearBlock),
        })),
      })),
    },
  };
}

async function createBackupBlob(
  kind: DraftBackupKind,
  documentData: OtDocument | TeaDocument,
): Promise<Blob> {
  const zip = new JSZip();
  const images = kind === "tea"
    ? collectTeaImages(documentData as TeaDocument)
    : collectOtImages(documentData as OtDocument);
  const imageEntries: BackupImageEntry[] = [];
  const warnings: string[] = [];

  images.forEach(({ image, location }, index) => {
    const parsed = image.dataUrl ? parseImageDataUrl(image.dataUrl) : null;

    if (!parsed) {
      const warning = `${location}: imagem sem dados no rascunho.`;
      warnings.push(warning);
      imageEntries.push(createBackupImageEntry(image, { missing: true, warning }));
      return;
    }

    const path = `images/${String(index + 1).padStart(4, "0")}-${sanitizeFileName(
      image.name || image.id,
    )}.${parsed.extension}`;
    zip.file(path, parsed.bytes);
    imageEntries.push(createBackupImageEntry(image, {
      path,
      mimeType: parsed.mimeType,
    }));
  });

  const payload: DraftBackupPayload = {
    format: BACKUP_FORMAT,
    version: BACKUP_VERSION,
    kind,
    exportedAt: new Date().toISOString(),
    document: kind === "tea"
      ? stripImageDataFromTeaDocument(documentData as TeaDocument)
      : stripImageDataFromDocument(documentData as OtDocument),
    images: imageEntries,
    warnings,
  };

  zip.file("draft.json", JSON.stringify(payload, null, 2));

  return zip.generateAsync({ type: "blob" });
}

function createBackupImageEntry(
  image: EvidenceImage,
  extra: Pick<BackupImageEntry, "path" | "mimeType" | "missing" | "warning">,
): BackupImageEntry {
  return {
    id: image.id,
    label: image.label,
    name: image.name,
    width: image.width,
    height: image.height,
    originalBytes: image.originalBytes,
    savedBytes: image.savedBytes,
    optimized: image.optimized,
    ...extra,
  };
}

function validateBackupPayload(value: unknown): DraftBackupPayload {
  const payload = value as Partial<DraftBackupPayload>;

  if (
    payload.format !== BACKUP_FORMAT ||
    payload.version !== BACKUP_VERSION ||
    (payload.kind !== "ot" && payload.kind !== "tea") ||
    !payload.document ||
    !Array.isArray(payload.images)
  ) {
    throw new Error("Backup invalido: formato nao reconhecido.");
  }

  return {
    ...payload,
    exportedAt: typeof payload.exportedAt === "string" ? payload.exportedAt : "",
    warnings: Array.isArray(payload.warnings) ? payload.warnings : [],
  } as DraftBackupPayload;
}

function collectOtImages(documentData: OtDocument): LocatedImage[] {
  return Object.values(documentData.permissionBlocks).flatMap((block: PermissionBlock) =>
    block.tests.flatMap((test: PermissionBlockTest) => [
      ...test.result.legacyImages.map((image, index) => ({
        image,
        location: `${test.title || "Teste"} > Legado > Imagem ${index + 1}`,
      })),
      ...test.result.newImages.map((image, index) => ({
        image,
        location: `${test.title || "Teste"} > Novo > Imagem ${index + 1}`,
      })),
      ...(test.correction?.beforeImages ?? []).map((image, index) => ({
        image,
        location: `${test.title || "Teste"} > Antes > Imagem ${index + 1}`,
      })),
      ...(test.correction?.afterImages ?? []).map((image, index) => ({
        image,
        location: `${test.title || "Teste"} > Depois > Imagem ${index + 1}`,
      })),
    ]),
  );
}

function collectTeaImages(documentData: TeaDocument): LocatedImage[] {
  return [
    ...documentData.activityImages.map((image, index) => ({
      image,
      location: `Imagem geral ${index + 1}`,
    })),
    ...documentData.activities.flatMap((activity, activityIndex) => [
      ...collectTeaBlockImages(activity.blocks, `Atividade 2.${activityIndex + 1}`),
      ...activity.subActivities.flatMap((subActivity, subActivityIndex) =>
        collectTeaBlockImages(
          subActivity.blocks,
          `Subtopico 2.${activityIndex + 1}.${subActivityIndex + 1}`,
        ),
      ),
    ]),
  ];
}

function collectTeaBlockImages(blocks: TeaContentBlock[], parent: string): LocatedImage[] {
  return blocks.flatMap((block, blockIndex) =>
    block.type === "images"
      ? block.images.map((image, imageIndex) => ({
          image,
          location: `${parent} > Bloco ${blockIndex + 1} > Imagem ${imageIndex + 1}`,
        }))
      : [],
  );
}

function attachOtImageData(
  documentData: OtDocument,
  imageDataById: Map<string, string>,
): OtDocument {
  const hydrateImages = (images: EvidenceImage[]): EvidenceImage[] =>
    images.map((image) => ({
      ...image,
      dataUrl: imageDataById.get(image.id) ?? image.dataUrl,
    }));
  const hydrateResult = (result: TestResult): TestResult => ({
    ...result,
    legacyImages: hydrateImages(result.legacyImages),
    newImages: hydrateImages(result.newImages),
  });

  return {
    ...documentData,
    permissionBlocks: Object.fromEntries(
      Object.entries(documentData.permissionBlocks).map(([blockKey, block]) => [
        blockKey,
        {
          ...block,
          tests: block.tests.map((test) => ({
            ...test,
            result: hydrateResult(test.result),
            correction: test.correction
              ? {
                  ...test.correction,
                  beforeImages: hydrateImages(test.correction.beforeImages),
                  afterImages: hydrateImages(test.correction.afterImages),
                }
              : test.correction,
          })),
        },
      ]),
    ),
  };
}

function attachTeaImageData(
  documentData: TeaDocument,
  imageDataById: Map<string, string>,
): TeaDocument {
  const hydrateImages = (images: EvidenceImage[]): EvidenceImage[] =>
    images.map((image) => ({
      ...image,
      dataUrl: imageDataById.get(image.id) ?? image.dataUrl,
    }));
  const hydrateBlock = (block: TeaContentBlock): TeaContentBlock =>
    block.type === "images" ? { ...block, images: hydrateImages(block.images) } : block;

  return {
    ...documentData,
    activityImages: hydrateImages(documentData.activityImages),
    activities: documentData.activities.map((activity) => ({
      ...activity,
      blocks: activity.blocks.map(hydrateBlock),
      subActivities: activity.subActivities.map((subActivity) => ({
        ...subActivity,
        blocks: subActivity.blocks.map(hydrateBlock),
      })),
    })),
  };
}

function parseImageDataUrl(dataUrl: string):
  | { mimeType: string; extension: string; bytes: Uint8Array }
  | null {
  const match = /^data:([^;]+);base64,(.*)$/i.exec(dataUrl);

  if (!match) {
    return null;
  }

  let bytes: Uint8Array;

  try {
    bytes = base64ToBytes(match[2]);
  } catch {
    return null;
  }

  return {
    mimeType: match[1],
    extension: mimeTypeToExtension(match[1]),
    bytes,
  };
}

function base64ToBytes(value: string): Uint8Array {
  let binary = "";

  try {
    binary = window.atob(value);
  } catch {
    throw new Error("base64 invalido");
  }
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }

  return window.btoa(binary);
}

function mimeTypeToExtension(mimeType: string): string {
  if (mimeType === "image/jpeg" || mimeType === "image/jpg") {
    return "jpg";
  }

  return mimeType.replace(/^image\//, "").replace(/[^a-z0-9]+/gi, "") || "bin";
}

function createBackupFileName(prefix: string, title: string): string {
  return `${prefix} - backup - ${sanitizeFileName(title || "Documento")} - ${
    new Date().toISOString().slice(0, 10)
  }.zip`;
}

function sanitizeFileName(value: string): string {
  const sanitized = value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return sanitized || "documento";
}

async function downloadBackup(blob: Blob, fileName: string): Promise<void> {
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}
