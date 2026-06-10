import { createEmptyTestCorrection } from "./defaultDocument";
import type {
  EvidenceImage,
  OtDocument,
  TeaActivity,
  TeaContentBlock,
  TeaDocument,
  TestResult,
} from "./types";

const DB_NAME = "create-ot-images";
const STORE_NAME = "evidence-images";
const DB_VERSION = 1;

type StoredImage = {
  id: string;
  dataUrl: string;
  updatedAt: number;
};

type EvidenceImageDataEntry = {
  id: string;
  dataUrl: string;
};

export async function saveEvidenceImageData(id: string, dataUrl: string): Promise<void> {
  await saveEvidenceImageDataBatch([{ id, dataUrl }]);
}

export async function saveEvidenceImageDataBatch(
  images: EvidenceImageDataEntry[],
): Promise<void> {
  if (images.length === 0) {
    return;
  }

  const db = await openImageDatabase();
  const updatedAt = Date.now();

  await runImageTransaction(db, "readwrite", (store) => {
    images.forEach(({ id, dataUrl }) => {
      store.put({ id, dataUrl, updatedAt } satisfies StoredImage);
    });
  });
}

export async function loadEvidenceImageData(id: string): Promise<string | undefined> {
  const images = await loadEvidenceImageDataBatch([id]);
  return images[id];
}

export async function loadEvidenceImageDataBatch(
  ids: string[],
): Promise<Record<string, string | undefined>> {
  const uniqueIds = Array.from(new Set(ids)).filter(Boolean);

  if (uniqueIds.length === 0) {
    return {};
  }

  try {
    const db = await openImageDatabase();
    const images = await runImageTransaction(db, "readonly", async (store) => {
      const entries = await Promise.all(
        uniqueIds.map(async (id) => [
          id,
          (await requestValue<StoredImage | undefined>(store.get(id)))?.dataUrl,
        ] as const),
      );

      return Object.fromEntries(entries);
    });

    return images;
  } catch {
    return Object.fromEntries(uniqueIds.map((id) => [id, undefined]));
  }
}

export async function deleteEvidenceImageData(id: string): Promise<void> {
  await deleteEvidenceImageDataBatch([id]);
}

export async function deleteEvidenceImageDataBatch(ids: string[]): Promise<void> {
  const uniqueIds = Array.from(new Set(ids)).filter(Boolean);

  if (uniqueIds.length === 0) {
    return;
  }

  const db = await openImageDatabase();

  await runImageTransaction(db, "readwrite", (store) => {
    uniqueIds.forEach((id) => {
      store.delete(id);
    });
  });
}

export async function clearEvidenceImageData(): Promise<void> {
  const db = await openImageDatabase();

  await runImageTransaction(db, "readwrite", (store) => {
    store.clear();
  });
}

export async function persistEmbeddedEvidenceImages(
  documentData: OtDocument,
): Promise<OtDocument> {
  const images = getDocumentImages(documentData).flatMap((image) =>
    image.dataUrl ? [{ id: image.id, dataUrl: image.dataUrl }] : [],
  );

  await saveEvidenceImageDataBatch(images);

  return documentData;
}

export async function persistEmbeddedTeaImages(documentData: TeaDocument): Promise<TeaDocument> {
  const images = getTeaDocumentImages(documentData).flatMap((image) =>
    image.dataUrl ? [{ id: image.id, dataUrl: image.dataUrl }] : [],
  );

  await saveEvidenceImageDataBatch(images);

  return documentData;
}

export async function hydrateDocumentImages(documentData: OtDocument): Promise<OtDocument> {
  const missingImageIds = getDocumentImages(documentData).flatMap((image) =>
    image.dataUrl ? [] : [image.id],
  );
  const imageDataById = await loadEvidenceImageDataBatch(missingImageIds);
  const hydrateImages = (images: EvidenceImage[]): EvidenceImage[] =>
    images.map((image) => ({
      ...image,
      dataUrl: image.dataUrl ?? imageDataById[image.id],
    }));

  const permissionBlocks = Object.fromEntries(
    Object.entries(documentData.permissionBlocks).map(([blockKey, block]) => [
      blockKey,
      {
        ...block,
        tests: block.tests.map((test) => ({
          ...test,
          result: {
            ...test.result,
            legacyImages: hydrateImages(test.result.legacyImages),
            newImages: hydrateImages(test.result.newImages),
          },
          correction: {
            ...createEmptyTestCorrection(),
            ...test.correction,
            beforeImages: hydrateImages(test.correction?.beforeImages ?? []),
            afterImages: hydrateImages(test.correction?.afterImages ?? []),
          },
        })),
      },
    ]),
  );

  return {
    ...documentData,
    permissionBlocks,
  };
}

export async function hydrateTeaDocumentImages(documentData: TeaDocument): Promise<TeaDocument> {
  const missingImageIds = getTeaDocumentImages(documentData).flatMap((image) =>
    image.dataUrl ? [] : [image.id],
  );
  const imageDataById = await loadEvidenceImageDataBatch(missingImageIds);
  const hydrateImages = (images: EvidenceImage[]): EvidenceImage[] =>
    images.map((image) => ({
      ...image,
      dataUrl: image.dataUrl ?? imageDataById[image.id],
    }));

  const hydrateBlock = (block: TeaContentBlock): TeaContentBlock =>
    block.type === "images"
      ? {
          ...block,
          images: hydrateImages(block.images),
        }
      : block;

  const hydrateActivity = (activity: TeaActivity): TeaActivity => ({
    ...activity,
    blocks: activity.blocks.map(hydrateBlock),
    subActivities: activity.subActivities.map((subActivity) => ({
        ...subActivity,
        blocks: subActivity.blocks.map(hydrateBlock),
      })),
  });

  return {
    ...documentData,
    activityImages: hydrateImages(documentData.activityImages),
    activities: documentData.activities.map(hydrateActivity),
  };
}

function getDocumentImages(documentData: OtDocument): EvidenceImage[] {
  return Object.values(documentData.permissionBlocks).flatMap((block) =>
    block.tests.flatMap((test) => [
      ...test.result.legacyImages,
      ...test.result.newImages,
      ...(test.correction?.beforeImages ?? []),
      ...(test.correction?.afterImages ?? []),
    ]),
  );
}

function getTeaDocumentImages(documentData: TeaDocument): EvidenceImage[] {
  return [
    ...documentData.activityImages,
    ...documentData.activities.flatMap((activity) => [
      ...activity.blocks.flatMap(getTeaBlockImages),
      ...activity.subActivities.flatMap((subActivity) =>
        subActivity.blocks.flatMap(getTeaBlockImages),
      ),
    ]),
  ];
}

function getTeaBlockImages(block: TeaContentBlock): EvidenceImage[] {
  return block.type === "images" ? block.images : [];
}

function openImageDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("IndexedDB indisponivel."));
      return;
    }

    const request = window.indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;

      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function runImageTransaction<T>(
  db: IDBDatabase,
  mode: IDBTransactionMode,
  work: (store: IDBObjectStore) => T | Promise<T>,
): Promise<T> {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, mode);
    const store = transaction.objectStore(STORE_NAME);
    let result: T | Promise<T>;

    transaction.oncomplete = () => {
      Promise.resolve(result).then(resolve).catch(reject);
      db.close();
    };
    transaction.onerror = () => {
      reject(transaction.error);
      db.close();
    };
    transaction.onabort = () => {
      reject(transaction.error);
      db.close();
    };

    try {
      result = work(store);
    } catch (error) {
      reject(error);
      db.close();
    }
  });
}

function requestValue<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export function stripImageDataFromDocument(documentData: OtDocument): OtDocument {
  const stripResult = (result: TestResult): TestResult => ({
    ...result,
    legacyImages: result.legacyImages.map(stripImageData),
    newImages: result.newImages.map(stripImageData),
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
            result: stripResult(test.result),
            correction: {
              ...createEmptyTestCorrection(),
              ...test.correction,
              beforeImages: (test.correction?.beforeImages ?? []).map(stripImageData),
              afterImages: (test.correction?.afterImages ?? []).map(stripImageData),
            },
          })),
        },
      ]),
    ),
  };
}

export function stripImageDataFromTeaDocument(documentData: TeaDocument): TeaDocument {
  const stripBlock = (block: TeaContentBlock): TeaContentBlock =>
    block.type === "images"
      ? {
          ...block,
          images: block.images.map(stripImageData),
        }
      : block;

  const stripActivity = (activity: TeaActivity): TeaActivity => ({
    ...activity,
    blocks: activity.blocks.map(stripBlock),
    subActivities: activity.subActivities.map((subActivity) => ({
      ...subActivity,
      blocks: subActivity.blocks.map(stripBlock),
    })),
  });

  return {
    ...documentData,
    activityImages: documentData.activityImages.map(stripImageData),
    activities: documentData.activities.map(stripActivity),
  };
}

function stripImageData(image: EvidenceImage): EvidenceImage {
  const { dataUrl: _dataUrl, ...metadata } = image;

  return metadata;
}
