/**
 * 對帳總覽。
 *
 * 版面的優先順序是刻意的：先給「要不要採取行動」的答案（自動化率與
 * 需人工確認的筆數），再給四象限的分佈，最後才是效能數字。
 *
 * 四象限用狀態色（好／警告／嚴重／危急）而不是四個任意的分類色，
 * 因為這四格代表的是**事情的嚴重程度**，不是四個對等的類別。
 * 每一列都直接標了名稱與數字，顏色只是強化——資訊不會只靠顏色傳達。
 */

import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useState } from "react";

import {
  ApiError,
  api,
  formatPercent,
  type Metrics,
} from "../api/client";

const QUADRANTS = [
  { key: "matched", label: "已勾稽", color: "var(--good)", note: "兩邊都有、金額相符" },
  { key: "amount_variance", label: "金額差異", color: "var(--warning)", note: "兩邊都有、金額不符" },
  { key: "missing_in_settlement", label: "未撥款", color: "var(--critical)", note: "我方有、平台無" },
  { key: "missing_in_ledger", label: "漏記單", color: "var(--serious)", note: "平台有、我方無" },
] as const;

export function OverviewView({
  runId,
  onRunCreated,
}: {
  runId: string | null;
  onRunCreated: (runId: string) => void;
}) {
  const queryClient = useQueryClient();
  const [threshold, setThreshold] = useState("");

  const health = useQuery({ queryKey: ["health"], queryFn: api.health });

  const report = useQuery({
    queryKey: ["report", runId, "summary"],
    queryFn: () => api.report(runId!, { limit: 1 }),
    enabled: runId !== null,
  });

  const run = useMutation({
    mutationFn: () =>
      threshold ? api.reconcile({ accept_threshold: threshold }) : api.reconcile(),
    onSuccess: (data) => {
      onRunCreated(data.run_id);
      void queryClient.invalidateQueries({ queryKey: ["runs"] });
    },
  });

  const metrics: Metrics | undefined = report.data?.metrics;
  const ready = (health.data?.settlement_records ?? 0) > 0;

  return (
    <div className="stack">
      <div className="card">
        <h2>執行對帳</h2>
        <p className="hint">
          把資料庫裡的訂單與所有結算記錄跑一次三階段匹配。
          目前有訂單 {health.data?.orders ?? "—"} 筆、結算列{" "}
          {health.data?.settlement_records ?? "—"} 筆。
        </p>
        <div className="row">
          <button
            type="button"
            className="primary"
            disabled={run.isPending || !ready}
            onClick={() => run.mutate()}
          >
            {run.isPending ? "對帳中⋯" : "開始對帳"}
          </button>
          <label htmlFor="threshold" style={{ fontSize: "0.86rem", color: "var(--muted)" }}>
            Stage 3 門檻
          </label>
          <input
            id="threshold"
            type="text"
            value={threshold}
            placeholder="0.82"
            style={{ width: 90 }}
            onChange={(event) => setThreshold(event.target.value)}
          />
          <span style={{ fontSize: "0.82rem", color: "var(--muted)" }}>
            模糊匹配的採納分數，留空用預設值
          </span>
        </div>

        {!ready && (
          <p className="empty" style={{ paddingBottom: 8 }}>
            資料庫裡還沒有結算記錄。先到「上傳」頁匯入結算單。
          </p>
        )}

        {run.isError && (
          <p className="error" style={{ textAlign: "left", padding: "12px 0 0" }}>
            {run.error instanceof ApiError ? run.error.message : "對帳失敗"}
          </p>
        )}
      </div>

      {runId === null && (
        <div className="card">
          <p className="empty">
            還沒有對帳結果。按上面的「開始對帳」。
          </p>
        </div>
      )}

      {report.isLoading && runId !== null && (
        <div className="card">
          <p className="loading">載入報告中⋯</p>
        </div>
      )}

      {metrics && (
        <>
          <div className="card">
            <h2>結果</h2>
            <p className="hint">
              執行編號 <code>{runId}</code>
            </p>

            <div className="tiles">
              <div className="tile">
                <div className="label">自動化率</div>
                <div className="value">{formatPercent(metrics.automation_rate)}</div>
                <div className="note">僅計 Stage 1</div>
              </div>
              <div className="tile">
                <div className="label">需人工確認</div>
                <div className="value">{metrics.needs_review}</div>
                <div className="note">筆</div>
              </div>
              <div className="tile">
                <div className="label">處理筆數</div>
                <div className="value">{metrics.record_count}</div>
                <div className="note">結算列（已去重）</div>
              </div>
              <div className="tile">
                <div className="label">耗時</div>
                <div className="value">
                  {(metrics.elapsed_seconds * 1000).toFixed(0)}
                </div>
                <div className="note">毫秒</div>
              </div>
            </div>

            <h3 style={{ marginTop: 26, marginBottom: 12 }}>四象限</h3>
            <QuadrantBars metrics={metrics} />

            <div className="note-strip">
              <strong>自動化率只計 Stage 1</strong>——那是唯一可以直接入帳、
              不需要人看的情況。把 Stage 2（金額有差）與 Stage 3（模糊比對）
              算進去，這個數字會跳到 97%，但那兩種都必須有人確認過。
            </div>
          </div>

          <div className="card">
            <h2>各階段</h2>
            <p className="hint">
              階段本身就是信心水準：Stage 1 精確匹配可自動放行，
              Stage 3 是靠評分猜的，一定要人看過。
            </p>
            <div className="tiles">
              <div className="tile">
                <div className="label">Stage 1 精確</div>
                <div className="value">{metrics.stage_exact}</div>
                <div className="note">編號與金額皆符</div>
              </div>
              <div className="tile">
                <div className="label">Stage 2 容差</div>
                <div className="value">{metrics.stage_tolerant}</div>
                <div className="note">編號符、金額有差</div>
              </div>
              <div className="tile">
                <div className="label">Stage 3 模糊</div>
                <div className="value">{metrics.stage_fuzzy}</div>
                <div className="note">靠候選集評分找回</div>
              </div>
              <div className="tile">
                <div className="label">本次輸入的重複列</div>
                <div className="value">{metrics.duplicates_dropped}</div>
                <div className="note">匯入時已擋下的不計</div>
              </div>
            </div>

            <h3 style={{ marginTop: 26, marginBottom: 8 }}>索引的效果</h3>
            <p style={{ fontSize: "0.88rem", color: "var(--ink-2)" }}>
              Stage 3 實際比對了{" "}
              <strong className="num">
                {metrics.candidate_comparisons.toLocaleString()}
              </strong>{" "}
              次。若不建索引、每筆未匹配的結算列都跟所有訂單比一遍，會是{" "}
              <strong className="num">
                {(metrics.order_count * metrics.record_count).toLocaleString()}
              </strong>{" "}
              次。
            </p>
          </div>
        </>
      )}
    </div>
  );
}

/**
 * 四象限的比例長條。
 *
 * 每一列都有名稱、色點與數字——顏色不是唯一的區分方式。這很重要，
 * 因為 warning 與 serious 兩個狀態色在白底上的對比偏低，
 * 只靠顏色會有人看不出差別。
 */
function QuadrantBars({ metrics }: { metrics: Metrics }) {
  const values: Record<string, number> = {
    matched: metrics.matched,
    amount_variance: metrics.amount_variance,
    missing_in_settlement: metrics.missing_in_settlement,
    missing_in_ledger: metrics.missing_in_ledger,
  };
  const total = Object.values(values).reduce((sum, value) => sum + value, 0);

  return (
    <div className="quadrants">
      {QUADRANTS.map((quadrant) => {
        const count = values[quadrant.key] ?? 0;
        const share = total > 0 ? count / total : 0;
        return (
          <div className="quadrant" key={quadrant.key}>
            <span className="dot" style={{ background: quadrant.color }} />
            <span className="name">
              {quadrant.label}
              <span
                style={{
                  color: "var(--muted)",
                  fontSize: "0.78rem",
                  marginLeft: 8,
                }}
              >
                {quadrant.note}
              </span>
            </span>
            <span className="track">
              <span
                className="fill"
                style={{
                  width: `${share * 100}%`,
                  background: quadrant.color,
                }}
              />
            </span>
            <span className="count">
              {count}
              <span className="pct">{formatPercent(share)}</span>
            </span>
          </div>
        );
      })}
    </div>
  );
}
