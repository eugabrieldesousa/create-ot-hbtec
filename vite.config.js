import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
export default defineConfig({
    plugins: [react()],
    build: {
        rollupOptions: {
            output: {
                manualChunks: {
                    "vendor-docx": ["docx", "mammoth"],
                },
            },
        },
    },
    test: {
        environment: "jsdom",
    },
});
