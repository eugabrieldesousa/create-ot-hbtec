import type { EvidenceImage, OtDocument, TeaActivity, TeaDocument, TestResult } from "./types";

const DB_NAME = "create-ot-images";
const STORE_NAME = "evidence-images";
const DB_VERSION = 1;

type StoredImage = {
  id: string;
  dataUrl: string;
  updatedAt: number;
};

export async function saveEvidenceImageData(id: string, dataUrl: string): Promise<void> {
  const db = await openImageDatabase();

  await runImageTransaction(db, "readwrite", (store) => {
    store.put({ id, dataUrl, updatedAt: Date.now() } satisfies StoredImage);
  });
}

export async function loadEvidenceImageData(id: string): Promise<string | undefined> {
  try {
    const db = await openImageDatabase();
    const image = await runImageTransaction(db, "readonly", (store) =>
      requestValue<StoredImage>(store.get(id)),
    );

    return image?.dataUrl;
  } catch {
    return undefined;
  }
}

export async function deleteEvidenceImageData(id: string): Promise<void> {
  const db = await openImageDatabase();

  await runImageTransaction(db, "readwrite", (store) => {
    store.delete(id);
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
  const images = getDocumentImages(documentData).filter((image) => image.dataUrl);

  await Promise.all(
    images.map((image) => saveEvidenceImageData(image.id, image.dataUrl as string)),
  );

  return documentData;
}

export async function persistEmbeddedTeaImages(documentData: TeaDocument): Promise<TeaDocument> {
  const images = getTeaDocumentImages(documentData).filter((image) => image.dataUrl);

  await Promise.all(
    images.map((image) => saveEvidenceImageData(image.id, image.dataUrl as string)),
  );

  return documentData;
}

export async function hydrateDocumentImages(documentData: OtDocument): Promise<OtDocument> {
  const hydrateImages = async (images: EvidenceImage[]): Promise<EvidenceImage[]> =>
    Promise.all(
      images.map(async (image) => ({
        ...image,
        dataUrl: image.dataUrl ?? (await loadEvidenceImageData(image.id)),
      })),
    );

  const permissionBlocks = Object.fromEntries(
    await Promise.all(
      Object.entries(documentData.permissionBlocks).map(async ([blockKey, block]) => [
        blockKey,
        {
          ...block,
          tests: await Promise.all(
            block.tests.map(async (test) => ({
              ...test,
              result: {
                ...test.result,
                legacyImages: await hydrateImages(test.result.legacyImages),
                newImages: await hydrateImages(test.result.newImages),
              },
            })),
          ),
        },
      ]),
    ),
  );

  return {
    ...documentData,
    permissionBlocks,
  };
}

export async function hydrateTeaDocumentImages(documentData: TeaDocument): Promise<TeaDocument> {
  const hydrateImages = async (images: EvidenceImage[]): Promise<EvidenceImage[]> =>
    Promise.all(
      images.map(async (image) => ({
        ...image,
        dataUrl: image.dataUrl ?? (await loadEvidenceImageData(image.id)),
      })),
    );

  const hydrateActivity = async (activity: TeaActivity): Promise<TeaActivity> => ({
    ...activity,
    images: await hydrateImages(activity.images),
    subActivities: await Promise.all(
      activity.subActivities.map(async (subActivity) => ({
        ...subActivity,
        images: await hydrateImages(subActivity.images),
      })),
    ),
  });

  return {
    ...documentData,
    activityImages: await hydrateImages(documentData.activityImages),
    activities: await Promise.all(documentData.activities.map(hydrateActivity)),
  };
}

function getDocumentImages(documentData: OtDocument): EvidenceImage[] {
  return Object.values(documentData.permissionBlocks).flatMap((block) =>
    block.tests.flatMap((test) => [
      ...test.result.legacyImages,
      ...test.result.newImages,
    ]),
  );
}

function getTeaDocumentImages(documentData: TeaDocument): EvidenceImage[] {
  return [
    ...documentData.activityImages,
    ...documentData.activities.flatMap((activity) => [
      ...activity.images,
      ...activity.subActivities.flatMap((subActivity) => subActivity.images),
    ]),
  ];
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
          })),
        },
      ]),
    ),
  };
}

export function stripImageDataFromTeaDocument(documentData: TeaDocument): TeaDocument {
  const stripActivity = (activity: TeaActivity): TeaActivity => ({
    ...activity,
    images: activity.images.map(stripImageData),
    subActivities: activity.subActivities.map((subActivity) => ({
      ...subActivity,
      images: subActivity.images.map(stripImageData),
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
