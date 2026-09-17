import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";

import App from "./App";
import "./styles.css";

/**
 * TanStack Query 管理所有的 server state。
 *
 * 為什麼不用 useState + useEffect 自己抓資料：因為那要自己處理載入中、
 * 錯誤、快取失效、重新抓取、race condition⋯⋯每個元件重寫一次。
 * server state 跟 UI state 是兩種不同的東西，混在一起管理是很多
 * React 專案變複雜的起點。
 */
const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      // 對帳結果不會自己變，切回分頁時不需要重抓
      refetchOnWindowFocus: false,
      // 4xx 是請求本身有問題，重試沒有意義
      retry: (failureCount, error) => {
        const status = (error as { status?: number }).status ?? 0;
        if (status >= 400 && status < 500) return false;
        return failureCount < 2;
      },
    },
  },
});

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <QueryClientProvider client={queryClient}>
      <App />
    </QueryClientProvider>
  </StrictMode>,
);
