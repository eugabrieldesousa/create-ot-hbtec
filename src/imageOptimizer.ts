export type OptimizedImageData = {
  dataUrl: string;
  width: number;
  height: number;
  originalBytes: number;
  savedBytes: number;
  optimized: boolean;
};

type OptimizeOptions = {
  maxDimension?: number;
  quality?: number;
  minReduction?: number;
};

const DEFAULT_MAX_DIMENSION = 1600;
const DEFAULT_QUALITY = 0.88;
const DEFAULT_MIN_REDUCTION = 0.1;

export async function optimizeImageFile(
  file: File,
  options: OptimizeOptions = {},
): Promise<OptimizedImageData> {
  const dataUrl = await fileToDataUrl(file);
  return optimizeImageDataUrl(dataUrl, options);
}

export async function optimizeImageDataUrl(
  dataUrl: string,
  options: OptimizeOptions = {},
): Promise<OptimizedImageData> {
  const originalBytes = dataUrlByteLength(dataUrl);
  const originalImage = await loadImage(dataUrl).catch(() => undefined);
  const originalSize = originalImage
    ? {
        width: originalImage.naturalWidth || 560,
        height: originalImage.naturalHeight || 320,
      }
    : { width: 560, height: 320 };
  const fallback: OptimizedImageData = {
    dataUrl,
    width: originalSize.width,
    height: originalSize.height,
    originalBytes,
    savedBytes: originalBytes,
    optimized: false,
  };

  if (!isOptimizableDataUrl(dataUrl)) {
    return fallback;
  }

  try {
    const maxDimension = options.maxDimension ?? DEFAULT_MAX_DIMENSION;
    const quality = options.quality ?? DEFAULT_QUALITY;
    const minReduction = options.minReduction ?? DEFAULT_MIN_REDUCTION;
    const scale = Math.min(maxDimension / originalSize.width, maxDimension / originalSize.height, 1);
    const width = Math.max(1, Math.round(originalSize.width * scale));
    const height = Math.max(1, Math.round(originalSize.height * scale));
    const image = originalImage;
    const canvas = window.document.createElement("canvas");
    const context = canvas.getContext("2d");

    if (!image || !context) {
      return fallback;
    }

    canvas.width = width;
    canvas.height = height;
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);

    const blob = await canvasToBlob(canvas, "image/jpeg", quality);
    if (!blob) {
      return fallback;
    }

    const savedBytes = blob.size;
    const shouldUseOptimized = savedBytes < originalBytes * (1 - minReduction);

    if (!shouldUseOptimized) {
      return fallback;
    }

    return {
      dataUrl: await blobToDataUrl(blob),
      width,
      height,
      originalBytes,
      savedBytes,
      optimized: true,
    };
  } catch {
    return fallback;
  }
}

export function dataUrlByteLength(dataUrl: string): number {
  const [, payload = ""] = dataUrl.split(",");

  if (dataUrl.includes(";base64,")) {
    const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
    return Math.max(0, Math.floor((payload.length * 3) / 4) - padding);
  }

  try {
    return new TextEncoder().encode(decodeURIComponent(payload)).length;
  } catch {
    return payload.length;
  }
}

function isOptimizableDataUrl(dataUrl: string): boolean {
  return /^data:image\/(?:png|jpe?g|webp|bmp);/i.test(dataUrl);
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Nao foi possivel ler a imagem."));
    image.src = dataUrl;
  });
}

function canvasToBlob(
  canvas: HTMLCanvasElement,
  type: string,
  quality: number,
): Promise<Blob | null> {
  return new Promise((resolve) => {
    canvas.toBlob(resolve, type, quality);
  });
}

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}
