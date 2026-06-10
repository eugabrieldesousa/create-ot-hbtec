import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  deleteEvidenceImageData,
  deleteEvidenceImageDataBatch,
  loadEvidenceImageData,
  loadEvidenceImageDataBatch,
  saveEvidenceImageData,
  saveEvidenceImageDataBatch,
} from "./imageStorage";

type TestRequest<T> = {
  result?: T;
  error?: Error | null;
  onsuccess?: (() => void) | null;
  onerror?: (() => void) | null;
  onupgradeneeded?: (() => void) | null;
};

type TestTransaction = {
  error?: Error | null;
  oncomplete?: (() => void) | null;
  onerror?: (() => void) | null;
  onabort?: (() => void) | null;
  objectStore: () => TestObjectStore;
};

type TestObjectStore = {
  put: (value: { id: string; dataUrl: string }) => TestRequest<undefined>;
  get: (id: string) => TestRequest<{ id: string; dataUrl: string } | undefined>;
  delete: (id: string) => TestRequest<undefined>;
  clear: () => TestRequest<undefined>;
};

type TestDatabase = {
  close: ReturnType<typeof vi.fn>;
  objectStoreNames: { contains: (name: string) => boolean };
  createObjectStore: ReturnType<typeof vi.fn>;
  transaction: ReturnType<typeof vi.fn>;
};

let storedImages: Map<string, { id: string; dataUrl: string }>;
let databases: TestDatabase[];

beforeEach(() => {
  storedImages = new Map();
  databases = [];

  Object.defineProperty(window, "indexedDB", {
    configurable: true,
    value: {
      open: vi.fn(() => {
        const request: TestRequest<TestDatabase> = {};
        const database = createTestDatabase();

        window.setTimeout(() => {
          request.result = database;
          request.onupgradeneeded?.();
          request.onsuccess?.();
        }, 0);

        return request;
      }),
    },
  });
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("imageStorage batch APIs", () => {
  it("saves, loads and deletes images in batches while preserving single-item wrappers", async () => {
    await saveEvidenceImageDataBatch([
      { id: "first", dataUrl: "data:image/png;base64,Zmlyc3Q=" },
      { id: "second", dataUrl: "data:image/png;base64,c2Vjb25k" },
    ]);
    await saveEvidenceImageData("third", "data:image/png;base64,dGhpcmQ=");

    await expect(loadEvidenceImageDataBatch(["first", "second", "missing"])).resolves.toEqual({
      first: "data:image/png;base64,Zmlyc3Q=",
      second: "data:image/png;base64,c2Vjb25k",
      missing: undefined,
    });
    await expect(loadEvidenceImageData("third")).resolves.toBe("data:image/png;base64,dGhpcmQ=");

    await deleteEvidenceImageDataBatch(["first", "missing"]);
    await deleteEvidenceImageData("third");

    await expect(loadEvidenceImageDataBatch(["first", "second", "third"])).resolves.toEqual({
      first: undefined,
      second: "data:image/png;base64,c2Vjb25k",
      third: undefined,
    });
    expect(databases.every((database) => database.close.mock.calls.length === 1)).toBe(true);
  });
});

function createTestDatabase(): TestDatabase {
  const database: TestDatabase = {
    close: vi.fn(),
    objectStoreNames: { contains: () => true },
    createObjectStore: vi.fn(),
    transaction: vi.fn(() => createTestTransaction()),
  };

  databases.push(database);
  return database;
}

function createTestTransaction(): TestTransaction {
  let pendingRequests = 0;
  let completed = false;
  const transaction: TestTransaction = {
    objectStore: () => createTestObjectStore(startRequest, finishRequest),
  };

  function startRequest(): void {
    pendingRequests += 1;
  }

  function finishRequest(): void {
    pendingRequests -= 1;
    scheduleComplete();
  }

  function scheduleComplete(): void {
    window.setTimeout(() => {
      if (!completed && pendingRequests === 0) {
        completed = true;
        transaction.oncomplete?.();
      }
    }, 0);
  }

  scheduleComplete();
  return transaction;
}

function createTestObjectStore(
  startRequest: () => void,
  finishRequest: () => void,
): TestObjectStore {
  return {
    put: (value) =>
      completeRequest(() => {
        storedImages.set(value.id, value);
        return undefined;
      }, startRequest, finishRequest),
    get: (id) =>
      completeRequest(() => storedImages.get(id), startRequest, finishRequest),
    delete: (id) =>
      completeRequest(() => {
        storedImages.delete(id);
        return undefined;
      }, startRequest, finishRequest),
    clear: () =>
      completeRequest(() => {
        storedImages.clear();
        return undefined;
      }, startRequest, finishRequest),
  };
}

function completeRequest<T>(
  readResult: () => T,
  startRequest: () => void,
  finishRequest: () => void,
): TestRequest<T> {
  const request: TestRequest<T> = {};
  startRequest();

  window.setTimeout(() => {
    request.result = readResult();
    request.onsuccess?.();
    finishRequest();
  }, 0);

  return request;
}
