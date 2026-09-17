/**
 * 型別安全的 API 客戶端。
 *
 * 這裡沒有一個型別是手寫的——全部從 `schema.d.ts` 來，而那個檔案是
 * `npm run gen:api` 從後端的 OpenAPI 文件產生的。
 *
 * 這代表什麼：如果後端把 `automation_rate` 改名成 `auto_rate`，
 * 重新產生型別之後，`npm run typecheck` 會在**編譯期**就指出前端哪幾行壞了，
 * 而不是等到使用者點開頁面看到 undefined。
 *
 * 這就是「前後端型別單一真實來源」的實際意義：真實來源是後端的
 * Pydantic schema，前端只是它的投影。
 */

import type { components, paths } from "./schema";

// ---------------------------------------------------------------------
// 從產生的 schema 取出我們要用的型別。
// 這些別名純粹是為了讓元件裡的寫法短一點，沒有任何自己的定義。
// ---------------------------------------------------------------------
export type Money = components["schemas"]["MoneyOut"];
export type ImportResponse = components["schemas"]["ImportResponse"];
export type BatchOut = components["schemas"]["BatchOut"];
export type MatchResult = components["schemas"]["MatchResultOut"];
export type Metrics = components["schemas"]["MetricsOut"];
export type ReportResponse = components["schemas"]["ReportResponse"];
export type ReconciliationResponse =
  components["schemas"]["ReconciliationResponse"];
export type Platform = components["schemas"]["PlatformOut"];
export type ErrorResponse = components["schemas"]["ErrorResponse"];

/** 四象限。直接從 OpenAPI 的查詢參數型別推出來，不是自己列的字串。 */
export type Outcome = NonNullable<
  paths["/api/reconciliations/{run_id}/report"]["get"]["parameters"]["query"]
>["outcome"];

/**
 * 後端回傳的錯誤。
 *
 * `code` 給程式判斷、`message` 給人看——後端刻意分開這兩者，
 * 前端才不需要比對中文字串來決定行為。
 */
export class ApiError extends Error {
  readonly code: string;
  readonly status: number;

  constructor(status: number, code: string, message: string) {
    super(message);
    this.name = "ApiError";
    this.code = code;
    this.status = status;
  }

  /** 認不出平台時可以請使用者手動指定，其他錯誤重試也沒用。 */
  get isRecoverable(): boolean {
    return this.code === "unrecognised_platform";
  }
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(path, init);

  if (!response.ok) {
    let code = "unknown";
    let message = `伺服器回應 ${response.status}`;
    try {
      const body = (await response.json()) as Partial<ErrorResponse>;
      if (body.code) code = body.code;
      if (body.message) message = body.message;
    } catch {
      // 回應不是 JSON（例如代理伺服器的錯誤頁），保留預設訊息
    }
    throw new ApiError(response.status, code, message);
  }

  return (await response.json()) as T;
}

// ---------------------------------------------------------------------
export const api = {
  health: () =>
    request<{
      status: string;
      orders: number;
      settlement_records: number;
      platforms: string[];
    }>("/api/health"),

  platforms: () => request<Platform[]>("/api/platforms"),

  listImports: () => request<BatchOut[]>("/api/imports"),

  /**
   * 上傳結算單。
   *
   * 重複上傳是安全的：後端以檔案內容的 SHA-256 判斷，第二次會回傳
   * 原批次且 `was_duplicate` 為真。前端因此不需要自己防連擊——
   * 那種前端防護只是讓錯誤更難重現，擋不住真正的重複。
   */
  upload: (file: File, platform?: string) => {
    const form = new FormData();
    form.append("file", file);
    const query = platform ? `?platform=${encodeURIComponent(platform)}` : "";
    return request<ImportResponse>(`/api/imports${query}`, {
      method: "POST",
      body: form,
    });
  },

  reconcile: (body: { platform?: string; accept_threshold?: string } = {}) =>
    request<ReconciliationResponse>("/api/reconciliations", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    }),

  listRuns: () => request<string[]>("/api/reconciliations"),

  report: (
    runId: string,
    options: {
      outcome?: Outcome;
      needsReview?: boolean;
      limit?: number;
      offset?: number;
    } = {},
  ) => {
    const query = new URLSearchParams();
    if (options.outcome) query.set("outcome", options.outcome);
    if (options.needsReview !== undefined) {
      query.set("needs_review", String(options.needsReview));
    }
    query.set("limit", String(options.limit ?? 50));
    query.set("offset", String(options.offset ?? 0));
    return request<ReportResponse>(
      `/api/reconciliations/${encodeURIComponent(runId)}/report?${query}`,
    );
  },
};

// ---------------------------------------------------------------------
/**
 * 金額的顯示格式。
 *
 * 後端把金額當字串傳（`"944.70"`），因為 JSON 的 number 是 IEEE 754
 * 雙精度浮點數，無法精確表示十進位小數。所以這裡**只做顯示層的加工**，
 * 絕不把它 parseFloat 之後再拿去算——那等於把後端避開的問題請回來。
 *
 * 需要計算時用 `minor_units`（整數分）。
 */
export function formatMoney(money: Money | null | undefined): string {
  if (!money) return "—";
  const negative = money.amount.startsWith("-");
  const digits = negative ? money.amount.slice(1) : money.amount;
  const [whole = "0", fraction] = digits.split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const sign = negative ? "-" : "";
  return fraction ? `${sign}${grouped}.${fraction}` : `${sign}${grouped}`;
}

export function formatPercent(value: number): string {
  return `${(value * 100).toFixed(1)}%`;
}
