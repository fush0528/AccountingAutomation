/**
 * 明細頁：依象限篩選，點開看單筆的完整內容。
 *
 * 這一頁的設計重點是「讓人能做決定」。對帳系統的價值不在於指出
 * 「這筆差了 80 元」——那用 Excel 也做得到——而在於回答「為什麼差」
 * 以及「所以我該做什麼」。所以每一筆差異都帶著歸因，
 * 每一個 Stage 3 的候選都帶著可讀的理由。
 */

import { keepPreviousData, useQuery } from "@tanstack/react-query";
import { useState } from "react";

import {
  ApiError,
  api,
  formatMoney,
  type MatchResult,
  type Outcome,
} from "../api/client";

const FILTERS: { key: Outcome | "all" | "review"; label: string }[] = [
  { key: "all", label: "全部" },
  { key: "review", label: "需人工確認" },
  { key: "matched", label: "已勾稽" },
  { key: "amount_variance", label: "金額差異" },
  { key: "missing_in_settlement", label: "未撥款" },
  { key: "missing_in_ledger", label: "漏記單" },
];

const PAGE_SIZE = 25;

const VARIANCE_LABELS: Record<string, string> = {
  fee_rate_change: "費率調整",
  shipping_subsidy: "運費補貼或折扣",
  partial_refund: "部分退款",
  rounding: "捨入誤差",
  unknown: "無法歸因，需人工追查",
};

export function DetailView({ runId }: { runId: string | null }) {
  const [filter, setFilter] = useState<Outcome | "all" | "review">("review");
  const [page, setPage] = useState(0);
  const [expanded, setExpanded] = useState<number | null>(null);

  const report = useQuery({
    queryKey: ["report", runId, filter, page],
    queryFn: () =>
      api.report(runId!, {
        ...(filter !== "all" && filter !== "review" ? { outcome: filter } : {}),
        ...(filter === "review" ? { needsReview: true } : {}),
        limit: PAGE_SIZE,
        offset: page * PAGE_SIZE,
      }),
    enabled: runId !== null,
    // 換頁時保留上一頁的資料，避免表格閃一下空白
    placeholderData: keepPreviousData,
  });

  if (runId === null) {
    return (
      <div className="card">
        <p className="empty">還沒有對帳結果。先到「總覽」頁執行一次對帳。</p>
      </div>
    );
  }

  const results = report.data?.results ?? [];
  const total = report.data?.total ?? 0;
  const pages = Math.ceil(total / PAGE_SIZE);

  return (
    <div className="card">
      <h2>對帳明細</h2>
      <p className="hint">
        點任一列可展開，看到結算列的原始內容與（Stage 3 的）候選理由。
      </p>

      <div className="row">
        {FILTERS.map((item) => (
          <button
            key={item.key ?? "none"}
            type="button"
            className="ghost"
            aria-pressed={filter === item.key}
            onClick={() => {
              setFilter(item.key);
              setPage(0);
              setExpanded(null);
            }}
          >
            {item.label}
          </button>
        ))}
      </div>

      {report.isError && (
        <p className="error">
          {report.error instanceof ApiError
            ? report.error.message
            : "載入失敗"}
        </p>
      )}

      {report.isLoading && <p className="loading">載入中⋯</p>}

      {!report.isLoading && results.length === 0 && (
        <p className="empty">這個象限沒有資料。</p>
      )}

      {results.length > 0 && (
        <>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>訂單編號</th>
                  <th>商品</th>
                  <th>階段</th>
                  <th style={{ textAlign: "right" }}>預期實收</th>
                  <th style={{ textAlign: "right" }}>平台實撥</th>
                  <th style={{ textAlign: "right" }}>差異</th>
                  <th>歸因</th>
                </tr>
              </thead>
              <tbody>
                {results.map((result, index) => (
                  <Row
                    key={`${result.outcome}-${index}`}
                    result={result}
                    expanded={expanded === index}
                    onToggle={() => setExpanded(expanded === index ? null : index)}
                  />
                ))}
              </tbody>
            </table>
          </div>

          <div className="row" style={{ marginTop: 14, justifyContent: "space-between" }}>
            <span style={{ fontSize: "0.85rem", color: "var(--muted)" }}>
              共 {total} 筆，第 {page + 1} / {Math.max(pages, 1)} 頁
            </span>
            <span className="row">
              <button
                type="button"
                className="ghost"
                disabled={page === 0}
                onClick={() => {
                  setPage((value) => value - 1);
                  setExpanded(null);
                }}
              >
                上一頁
              </button>
              <button
                type="button"
                className="ghost"
                disabled={page + 1 >= pages}
                onClick={() => {
                  setPage((value) => value + 1);
                  setExpanded(null);
                }}
              >
                下一頁
              </button>
            </span>
          </div>
        </>
      )}
    </div>
  );
}

function Row({
  result,
  expanded,
  onToggle,
}: {
  result: MatchResult;
  expanded: boolean;
  onToggle: () => void;
}) {
  const orderId = result.order?.external_order_id ?? "—";
  const product =
    result.order?.product_name ?? result.records[0]?.product_name ?? "—";

  return (
    <>
      <tr className="clickable" onClick={onToggle}>
        <td className="mono">{orderId}</td>
        <td>{product}</td>
        <td>
          {result.stage ? (
            <span className={`badge stage${result.stage}`}>
              Stage {result.stage}
            </span>
          ) : (
            <span className="badge">未匹配</span>
          )}
        </td>
        <td className="right">{formatMoney(result.order?.expected_net)}</td>
        <td className="right">{formatMoney(result.settled_amount)}</td>
        <td className="right">{formatMoney(result.variance)}</td>
        <td>
          {result.variance_reason
            ? (VARIANCE_LABELS[result.variance_reason] ?? result.variance_reason)
            : "—"}
        </td>
      </tr>

      {expanded && (
        <tr>
          <td className="detail" colSpan={7}>
            <div style={{ display: "grid", gap: 14 }}>
              <div>
                <strong>結算列（{result.records.length} 列）</strong>
                {result.records.length === 0 && (
                  <p style={{ color: "var(--muted)", marginTop: 4 }}>
                    平台完全沒有這筆的記錄——這就是「未撥款」的意思。
                  </p>
                )}
                {result.records.map((record, index) => (
                  <p key={index} style={{ marginTop: 4 }} className="mono">
                    第 {record.source_row} 列 · {record.settled_at} ·{" "}
                    {record.record_type} ·{" "}
                    {record.external_order_id ?? "（無訂單編號）"} · 撥款{" "}
                    {formatMoney(record.net_amount)}
                  </p>
                ))}
              </div>

              {result.candidates.length > 0 && (
                <div>
                  <strong>Stage 3 候選</strong>
                  <p style={{ color: "var(--muted)", fontSize: "0.82rem" }}>
                    分數不到門檻，系統不敢自動認。理由列出來讓你判斷——
                    只給一個 0.87 的分數，沒有人敢按確認。
                  </p>
                  <div style={{ marginTop: 6 }}>
                    {result.candidates.map((candidate) => (
                      <div className="candidate" key={candidate.order_id}>
                        <span className="score">
                          {Number(candidate.score).toFixed(2)}
                        </span>
                        <span>
                          <span className="mono">{candidate.external_order_id}</span>{" "}
                          {candidate.product_name}
                          <div className="reasons">
                            {candidate.reasons.join("、")}
                          </div>
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {result.needs_review && (
                <p style={{ color: "var(--muted)", fontSize: "0.82rem" }}>
                  這筆需要人工確認。只有 Stage 1 的精確匹配可以直接入帳。
                </p>
              )}
            </div>
          </td>
        </tr>
      )}
    </>
  );
}
