import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

/**
 * 開發時前端跑在 5173、後端跑在 8000，是兩個不同的來源。
 * 直接 fetch 會被瀏覽器的同源政策擋下。
 *
 * 有兩種解法：後端開 CORS，或前端用 proxy。這裡選 proxy，因為
 * 開 CORS 是為了瀏覽器安全而放寬限制，只為了本機開發方便就放寬
 * 不划算；proxy 則讓瀏覽器眼中「只有一個來源」，正式環境
 * （FastAPI 直接託管 build 產物）本來就是同源，行為完全一致。
 */
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      "/api": { target: "http://127.0.0.1:8000", changeOrigin: true },
      "/openapi.json": { target: "http://127.0.0.1:8000", changeOrigin: true },
    },
  },
  build: {
    // 產物直接放進後端會託管的目錄，省掉一個複製步驟
    outDir: "../src/reconciliation/api/static",
    emptyOutDir: true,
  },
});
