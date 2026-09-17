/**
 * 版面骨架與分頁切換。
 *
 * 刻意不裝 react-router。整個應用只有三個畫面、沒有深層連結需求，
 * 一個 useState 就夠了——裝一個路由函式庫只是多一份相依與多一層概念。
 * 需要「重新整理後停在同一頁」時再加也不遲。
 *
 * 這個取捨本身是可以拿來講的：選擇不用某個套件，跟選擇用它一樣是設計決定。
 */

import { useQuery } from "@tanstack/react-query";
import { useState } from "react";

import { api } from "./api/client";
import { DetailView } from "./views/DetailView";
import { ImportView } from "./views/ImportView";
import { OverviewView } from "./views/OverviewView";

type Tab = "import" | "overview" | "detail";

const TABS: { key: Tab; label: string }[] = [
  { key: "import", label: "上傳" },
  { key: "overview", label: "總覽" },
  { key: "detail", label: "明細" },
];

export default function App() {
  const [tab, setTab] = useState<Tab>("import");
  const [runId, setRunId] = useState<string | null>(null);

  const health = useQuery({ queryKey: ["health"], queryFn: api.health });

  return (
    <div className="app">
      <header className="masthead">
        <h1>多平台電商對帳引擎</h1>
        <span className="spacer" />
        <span className="sub">
          {health.data
            ? `訂單 ${health.data.orders} · 結算列 ${health.data.settlement_records}`
            : "連線中⋯"}
        </span>
      </header>

      <nav className="tabs" role="tablist">
        {TABS.map((item) => (
          <button
            key={item.key}
            type="button"
            role="tab"
            aria-selected={tab === item.key}
            onClick={() => setTab(item.key)}
          >
            {item.label}
          </button>
        ))}
      </nav>

      {tab === "import" && <ImportView />}
      {tab === "overview" && (
        <OverviewView
          runId={runId}
          onRunCreated={(id) => {
            setRunId(id);
          }}
        />
      )}
      {tab === "detail" && <DetailView runId={runId} />}

      <footer>
        前端的型別由 <code>npm run gen:api</code> 從後端的 OpenAPI 文件產生，
        沒有一行是手寫的。後端改了欄位名稱，<code>npm run typecheck</code>{" "}
        會在編譯期就指出前端哪裡壞了。
      </footer>
    </div>
  );
}
