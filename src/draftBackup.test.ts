import JSZip from "jszip";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  exportOtBackup,
  exportTeaBackup,
  parseBackupFile,
  removeAllOtImages,
  removeAllTeaImages,
} from "./draftBackup";
import type { CheckKey, EvidenceImage, OtDocument, TeaDocument, TestResult } from "./types";

const imageStorageMocks = vi.hoisted(() => ({
  hydrateDocumentImages: vi.fn(),
  hydrateTeaDocumentImages: vi.fn(),
}));

vi.mock("./imageStorage", async () => {
  const actual = await vi.importActual<typeof import("./imageStorage")>("./imageStorage");

  return {
    ...actual,
    hydrateDocumentImages: imageStorageMocks.hydrateDocumentImages,
    hydrateTeaDocumentImages: imageStorageMocks.hydrateTeaDocumentImages,
  };
});

let lastDownloadedBlob: Blob | null = null;

beforeEach(() => {
  imageStorageMocks.hydrateDocumentImages.mockReset();
  imageStorageMocks.hydrateDocumentImages.mockImplementation(
    async (documentData: OtDocument) => documentData,
  );
  imageStorageMocks.hydrateTeaDocumentImages.mockReset();
  imageStorageMocks.hydrateTeaDocumentImages.mockImplementation(
    async (documentData: TeaDocument) => documentData,
  );
  lastDownloadedBlob = null;

  vi.spyOn(URL, "createObjectURL").mockImplementation((blob) => {
    lastDownloadedBlob = blob as Blob;
    return "blob:backup";
  });
  vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);
});

describe("draft backup export", () => {
  it("exports an OT backup with draft.json and image files", async () => {
    await exportOtBackup(createOtDraftWithImages());

    const zip = await loadDownloadedZip();
    const payload = JSON.parse(await zip.file("draft.json")!.async("string"));

    expect(payload.format).toBe("create-ot-backup");
    expect(payload.kind).toBe("ot");
    expect(payload.images).toHaveLength(2);
    expect(payload.images[0].path).toMatch(/^images\/0001-/);
    expect(zip.file(payload.images[0].path)).toBeTruthy();
    expect(payload.document.permissionBlocks["macro-a:micro-a"].tests[0].result.legacyImages[0])
      .not.toHaveProperty("dataUrl");
  });

  it("exports a TEA backup even when an image block is empty", async () => {
    const documentData = createTeaDraftWithImages();
    documentData.activities[0].blocks.push({
      id: "empty-images",
      type: "images",
      images: [],
    });

    await exportTeaBackup(documentData);

    const zip = await loadDownloadedZip();
    const payload = JSON.parse(await zip.file("draft.json")!.async("string"));

    expect(payload.kind).toBe("tea");
    expect(payload.images).toHaveLength(2);
    expect(payload.warnings).toEqual([]);
  });

  it("keeps image metadata and warning when a registered image has no data", async () => {
    const documentData = createTeaDraftWithImages();
    documentData.activities[0].blocks[1] = {
      id: "missing-images",
      type: "images",
      images: [createImageWithoutData("missing-image")],
    };

    await exportTeaBackup(documentData);

    const zip = await loadDownloadedZip();
    const payload = JSON.parse(await zip.file("draft.json")!.async("string"));
    const missingEntry = payload.images.find(
      (image: { id: string }) => image.id === "missing-image",
    );

    expect(missingEntry.missing).toBe(true);
    expect(missingEntry.warning).toContain("imagem sem dados");
    expect(payload.warnings[0]).toContain("imagem sem dados");
    expect(zip.file(missingEntry.path ?? "")).toBeNull();
  });
});

describe("draft backup import", () => {
  it("imports a backup and restores image data urls", async () => {
    const sourceDocument = createTeaDraftWithImages();
    const zip = new JSZip();

    zip.file(
      "draft.json",
      JSON.stringify({
        format: "create-ot-backup",
        version: 1,
        kind: "tea",
        exportedAt: "2026-06-17T00:00:00.000Z",
        document: {
          ...sourceDocument,
          activities: [
            {
              ...sourceDocument.activities[0],
              blocks: [
                sourceDocument.activities[0].blocks[0],
                {
                  id: "activity-images",
                  type: "images",
                  images: [createImageWithoutData("activity-image")],
                },
              ],
            },
          ],
        },
        images: [
          {
            id: "activity-image",
            name: "activity-image.png",
            label: "Evidencia",
            width: 10,
            height: 10,
            path: "images/activity-image.png",
            mimeType: "image/png",
          },
        ],
        warnings: [],
      }),
    );
    zip.file("images/activity-image.png", Uint8Array.from([1, 2, 3]));

    const result = await parseBackupFile(await zipToFile(zip));

    expect(result.kind).toBe("tea");
    if (result.kind !== "tea") {
      throw new Error("Expected TEA backup");
    }
    expect(result.warnings).toEqual([]);
    expect(
      result.document.activities[0].blocks.find((block) => block.type === "images")?.images[0]
        .dataUrl,
    ).toBe("data:image/png;base64,AQID");
  });

  it("reports an invalid ZIP", async () => {
    await expect(
      parseBackupFile(new File(["invalid"], "backup.zip", { type: "application/zip" })),
    ).rejects.toThrow("Backup invalido");
  });

  it("reports a missing draft.json", async () => {
    const zip = new JSZip();
    zip.file("other.json", "{}");

    await expect(parseBackupFile(await zipToFile(zip))).rejects.toThrow("draft.json");
  });

  it("reports listed images that are missing from the ZIP", async () => {
    const documentData = createOtDraftWithImages();
    const zip = new JSZip();

    zip.file(
      "draft.json",
      JSON.stringify({
        format: "create-ot-backup",
        version: 1,
        kind: "ot",
        exportedAt: "2026-06-17T00:00:00.000Z",
        document: documentData,
        images: [
          {
            id: "legacy-image",
            name: "legacy-image.png",
            label: "Evidencia",
            width: 10,
            height: 10,
            path: "images/missing.png",
            mimeType: "image/png",
          },
        ],
        warnings: [],
      }),
    );

    const result = await parseBackupFile(await zipToFile(zip));

    expect(result.kind).toBe("ot");
    expect(result.warnings[0]).toContain("images/missing.png");
  });
});

describe("draft backup image clearing", () => {
  it("removes all OT images and returns their ids", () => {
    const result = removeAllOtImages(createOtDraftWithImages());
    const test = result.document.permissionBlocks["macro-a:micro-a"].tests[0];

    expect(result.imageIds).toEqual(["legacy-image", "new-image"]);
    expect(test.result.legacyImages).toEqual([]);
    expect(test.result.newImages).toEqual([]);
  });

  it("removes all TEA images and keeps empty image blocks", () => {
    const result = removeAllTeaImages(createTeaDraftWithImages());
    const imageBlock = result.document.activities[0].blocks.find((block) => block.type === "images");

    expect(result.imageIds).toEqual(["general-image", "activity-image"]);
    expect(result.document.activityImages).toEqual([]);
    expect(imageBlock?.type).toBe("images");
    expect(imageBlock?.images).toEqual([]);
  });
});

async function loadDownloadedZip(): Promise<JSZip> {
  expect(lastDownloadedBlob).toBeTruthy();

  return JSZip.loadAsync(await lastDownloadedBlob!.arrayBuffer());
}

async function zipToFile(zip: JSZip): Promise<File> {
  return new File([await zip.generateAsync({ type: "blob" })], "backup.zip", {
    type: "application/zip",
  });
}

function createOtDraftWithImages(): OtDocument {
  return {
    metadata: {
      screen: "Cadastro",
      responsible: "GABRIEL",
      date: "2026-06-17",
      environment: "LOCAL",
      author: "GABRIEL",
    },
    objective: "Validar tela.",
    accessSteps: [{ id: "step-1", text: "Acessar" }],
    permissionGroups: [
      {
        id: "macro-a",
        code: "MA",
        label: "Macro A",
        selected: true,
        microPermissions: [
          {
            id: "micro-a",
            code: "AT",
            label: "Atualizacao",
            selected: true,
          },
        ],
      },
    ],
    permissionBlocks: {
      "macro-a:micro-a": {
        tests: [
          {
            id: "test-a",
            title: "Teste A",
            result: {
              ...createResult({ sameBehavior: true }),
              legacyImages: [createImage("legacy-image")],
              newImages: [createImage("new-image")],
            },
          },
        ],
      },
    },
  };
}

function createTeaDraftWithImages(): TeaDocument {
  return {
    metadata: {
      serviceOrder: "OS2171",
      phase: "Etapa 5",
      ticket: "Chamado 202504000396",
      subject: "Telas",
      date: "2026-06-17",
      author: "Gabriel",
    },
    overview: "Validar telas.",
    activityIntro: "Atividades realizadas:",
    activityImages: [createImage("general-image")],
    activities: [
      {
        id: "activity-a",
        title: "Atividade A",
        blocks: [
          {
            id: "text-a",
            type: "text",
            text: "Texto",
          },
          {
            id: "activity-images",
            type: "images",
            images: [createImage("activity-image")],
          },
        ],
        subActivities: [],
      },
    ],
  };
}

function createResult(checks: Partial<Record<CheckKey, boolean>>): TestResult {
  return {
    checks: {
      sameBehavior: false,
      possibleIssue: false,
      bothIssue: false,
      newIssue: false,
      errorReport: false,
      ...checks,
    },
    observations: "",
    legacyImages: [],
    newImages: [],
    errors: [],
  };
}

function createImage(id: string): EvidenceImage {
  return {
    id,
    label: "Evidencia",
    name: `${id}.png`,
    width: 10,
    height: 10,
    dataUrl: "data:image/png;base64,AQID",
  };
}

function createImageWithoutData(id: string): EvidenceImage {
  const { dataUrl: _dataUrl, ...image } = createImage(id);

  return image;
}
