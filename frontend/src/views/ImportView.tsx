/**
 * 上傳頁。
 *
 * 這一頁最重要的互動細節：**不做前端防連擊**。
 *
 * 一般前端會在上傳中把按鈕 disable，避免使用者重複送出。但那擋不住
 * 真正的重複——網路逾時後的自動重試、使用者按重新整理再傳一次、
 * 兩個分頁同時操作，前端都攔不到。後端的冪等匯入才是真的防線。
 *
 * 所以這裡反過來做：讓使用者可以重複傳，然後把 `was_duplicate`
 * 明確顯示出來。重複上傳不是要被阻止的錯誤，是一個正常且安全的結果。
 */

import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useRef, useState } from "react";

import { ApiError, api, type ImportResponse } from "../api/client";

type Entry =
  | { kind: "ok"; filename: string; data: ImportResponse }
  | { kind: "error"; filename: string; code: string; message: string };

export function ImportView() {
  const queryClient = useQueryClient();
  const [entries, setEntries] = useState<Entry[]>([]);
  const [dragging, setDragging] = useState(false);
  const [platform, setPlatform] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const platforms = useQuery({
    queryKey: ["platforms"],
    queryFn: api.platforms,
  });

  const batches = useQuery({ queryKey: ["imports"], queryFn: api.listImports });

  const upload = useMutation({
    mutationFn: (file: File) =>
      platform ? api.upload(file, platform) : api.upload(file),
  });

  const [progress, setProgress] = useState<{ done: number; total: number } | null>(
    null,
  );

  /**
   * 逐一送出，而且是**依序等待**。
   *
   * 這裡踩過一個坑值得記下來：原本寫成
   *
   *     for (const file of files) upload.mutate(file);
   *
   * 三個檔案只有最後一個出現在畫面上。原因是一個 `useMutation` 同時只
   * 追蹤一個進行中的 mutation——連續呼叫 `mutate()` 時，後面的會接管
   * observer，前面那幾個的 `onSuccess` 就不會被呼叫。請求其實都送出去了，
   * 只是結果被丟掉，所以畫面看起來像卡在「上傳中」。
   *
   * 改用 `mutateAsync` 並 await，每一份完成後才送下一份。順便得到兩個好處：
   * 結果的順序是確定的，而且不會同時打一堆請求進去。
   */
  async function handleFiles(files: FileList | null) {
    if (!files || files.length === 0) return;
    const list = Array.from(files);
    setProgress({ done: 0, total: list.length });

    for (const [index, file] of list.entries()) {
      try {
        const data = await upload.mutateAsync(file);
        setEntries((prev) => [{ kind: "ok", filename: file.name, data }, ...prev]);
      } catch (error: unknown) {
        const failure =
          error instanceof ApiError
            ? error
            : new ApiError(0, "network", "無法連上伺服器");
        setEntries((prev) => [
          {
            kind: "error",
            filename: file.name,
            code: failure.code,
            message: failure.message,
          },
          ...prev,
        ]);
      }
      setProgress({ done: index + 1, total: list.length });
    }

    setProgress(null);
    void queryClient.invalidateQueries({ queryKey: ["imports"] });
    void queryClient.invalidateQueries({ queryKey: ["health"] });
  }

  return (
    <div className="stack">
      <div className="card">
        <h2>上傳結算單</h2>
        <p className="hint">
          支援多檔同時上傳，平台由檔案內容自動判斷。
          <strong>重複上傳同一份檔案是安全的</strong>——試試看把同一份傳兩次。
        </p>

        <div
          className={dragging ? "dropzone over" : "dropzone"}
          onDragOver={(event) => {
            event.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => {
            event.preventDefault();
            setDragging(false);
            handleFiles(event.dataTransfer.files);
          }}
        >
          <p className="big">把結算單拖到這裡</p>
          <p className="small">或</p>
          <p style={{ marginTop: 10 }}>
            <button
              type="button"
              className="primary"
              onClick={() => inputRef.current?.click()}
            >
              選擇檔案
            </button>
          </p>
          <input
            ref={inputRef}
            type="file"
            multiple
            hidden
            onChange={(event) => {
              const { files } = event.target;
              void handleFiles(files);
              event.target.value = "";
            }}
          />
        </div>

        <div className="row" style={{ marginTop: 14 }}>
          <label htmlFor="platform" style={{ fontSize: "0.86rem", color: "var(--muted)" }}>
            平台
          </label>
          <select
            id="platform"
            value={platform}
            onChange={(event) => setPlatform(event.target.value)}
          >
            <option value="">自動偵測</option>
            {platforms.data?.map((item) => (
              <option key={item.code} value={item.code}>
                {item.display_name}
              </option>
            ))}
          </select>
          <span style={{ fontSize: "0.82rem", color: "var(--muted)" }}>
            平台清單來自後端的 parser registry，新增平台不用改前端
          </span>
        </div>

        {progress && (
          <p className="loading">
            上傳中⋯ {progress.done} / {progress.total}
          </p>
        )}

        {entries.length > 0 && (
          <ul className="result-list">
            {entries.map((entry, index) => (
              <li
                key={`${entry.filename}-${index}`}
                className={
                  entry.kind === "error"
                    ? "result err"
                    : entry.data.was_duplicate
                      ? "result dup"
                      : "result"
                }
              >
                <span
                  className="tag"
                  style={{
                    color:
                      entry.kind === "error"
                        ? "var(--critical)"
                        : entry.data.was_duplicate
                          ? "var(--accent)"
                          : "var(--good)",
                  }}
                >
                  {entry.kind === "error"
                    ? "失敗"
                    : entry.data.was_duplicate
                      ? "重複"
                      : "已匯入"}
                </span>
                <div>
                  <div className="name">{entry.filename}</div>
                  {entry.kind === "error" ? (
                    <div className="msg">{entry.message}</div>
                  ) : (
                    <>
                      <div className="msg">{entry.data.message}</div>
                      <div
                        className="name"
                        style={{ marginTop: 3, color: "var(--muted)" }}
                      >
                        {entry.data.platform} · {entry.data.batch_id} ·
                        SHA-256 {entry.data.content_hash.slice(0, 16)}⋯
                      </div>
                    </>
                  )}
                </div>
              </li>
            ))}
          </ul>
        )}

        <div className="note-strip">
          後端用三道防線保證冪等：檔案內容的 SHA-256、每列的 fingerprint、
          整批單一交易。所以這個頁面沒有做任何防連擊——前端擋不住網路重試
          與重新整理，能擋的只有後端。
        </div>
      </div>

      <div className="card">
        <h2>已匯入的批次</h2>
        <p className="hint">
          每一份成功匯入的檔案是一個批次。重複上傳不會產生新批次。
        </p>
        {batches.isLoading && <p className="loading">載入中⋯</p>}
        {batches.data?.length === 0 && (
          <p className="empty">還沒有任何批次。上傳一份結算單試試。</p>
        )}
        {batches.data && batches.data.length > 0 && (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>批次</th>
                  <th>平台</th>
                  <th>檔名</th>
                  <th>匯入時間</th>
                  <th style={{ textAlign: "right" }}>成功</th>
                  <th style={{ textAlign: "right" }}>失敗</th>
                </tr>
              </thead>
              <tbody>
                {batches.data.map((batch) => (
                  <tr key={batch.batch_id}>
                    <td className="mono">{batch.batch_id}</td>
                    <td>{batch.platform}</td>
                    <td>{batch.filename}</td>
                    <td className="mono">
                      {new Date(batch.imported_at).toLocaleString("zh-TW")}
                    </td>
                    <td className="right">{batch.accepted_count}</td>
                    <td className="right">{batch.rejected_count}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
