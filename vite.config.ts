import { defineConfig } from "vite";
import { resolve } from "node:path";

// @ts-expect-error process is a nodejs global
const host = process.env.TAURI_DEV_HOST;

// https://vitejs.dev/config/
export default defineConfig(async () => ({
  base: "/ocp-IHwork/",

  // Vite options tailored for Tauri development and only applied in `tauri dev` or `tauri build`
  //
  // 1. prevent vite from obscuring rust errors
  clearScreen: false,
  // 2. tauri expects a fixed port, fail if that port is not available
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host
      ? {
          protocol: "ws",
          host,
          port: 1421,
        }
      : undefined,
    watch: {
      // 3. tell vite to ignore watching `src-tauri`
      ignored: ["**/src-tauri/**"],
    },
  },
  build: {
    rollupOptions: {
      input: {
        main: resolve(__dirname, "index.html"),
        solarProfitability: resolve(__dirname, "Solar Profitability.html"),
        storageCost: resolve(__dirname, "Storage Cost.html"),
        fullAnalysis: resolve(__dirname, "Full Analysis.html"),
        batteryParameters: resolve(__dirname, "Battery Parameters.html"),
        ocpEmsSimulator: resolve(__dirname, "OCP EMS Simulator.html"),
        scadaSupervision: resolve(__dirname, "scada_supervision.html"),
        competitives: resolve(__dirname, "COMPETITIVES.html"),
      },
    },
  },
}));
