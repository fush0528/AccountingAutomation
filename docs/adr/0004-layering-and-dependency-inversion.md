# ADR 0004：分層與依賴反轉，而且用測試守住

- 狀態：已採用
- 日期：2026-09-17
- 相關：[ADR 0001](0001-money-as-integer-minor-units.md)、[ADR 0003](0003-idempotent-import.md)

## 背景

v1（保留在 `legacy/`）是一個 902 行的 CLI 工具，結構是
`main.py` → `ExcelHandler` → `AccountingEntry`。這個結構沒有「錯」，
但它有一個性質：**每一個元件都知道資料存在 Excel 裡**。
`AccountingEntry` 的欄位是 `year`／`month`／`day`／`time` 四個字串，
因為 Excel 表格的欄位長那樣。業務概念被儲存格式塑形了。

v2 要解的問題複雜得多——三階段匹配、差異歸因、候選評分。
這些邏輯的正確性是整個專案的價值所在，而它們的測試如果需要啟動資料庫、
需要 HTTP client、需要一個真的 Excel 檔，就會慢到沒有人想跑，
然後就不會有人跑。

更具體的風險是：一旦對帳邏輯裡出現 `session.query(...)`，
那段邏輯就再也無法在沒有資料庫的情況下驗證。而對帳邏輯正是最需要
被大量、快速、反覆驗證的部分（`tests/domain/test_properties.py`
用 Hypothesis 產生數百組隨機輸入）。

## 決策

五層，依賴**只能由外向內**：

```
api/          FastAPI routers、Pydantic DTO。只處理 HTTP
services/     使用案例編排、交易邊界
parsers/      各平台實作 ＋ registry
repositories/ 介面 ＋ SQLAlchemy 實作
domain/       Money、領域模型、三階段匹配 ── 依賴圖的終點
```

**唯一的硬規則：`domain/` 不 import 任何框架。**
沒有 FastAPI、沒有 SQLAlchemy、沒有 pandas、沒有 openpyxl，只有標準函式庫。

三個推論：

**一、持久化透過介面反轉。**
`repositories/` 定義介面，SQLAlchemy 實作它。依賴方向是
「實作 → 介面 → domain」，而不是「domain → SQLAlchemy」。
這讓 `domain/` 完全不知道資料存在哪裡，也讓換資料庫變成換一個連線字串。

**二、DTO 不等於領域模型。**
`api/schemas.py` 是獨立的一層。`Money` 內部存整數分是為了算術正確；
但 API 回傳 `{"amount": "944.70", "currency": "TWD"}` 對前端才友善。
若讓領域模型直接序列化，遲早有人為了讓 JSON 好看去改 `Money`——那是本末倒置。
API 是對外承諾，領域模型是內部設計，兩者的變動頻率與變動原因完全不同。

**三、不寫原生 SQL，不用資料庫特有型別。**
沒有 `JSONB`、沒有 `ARRAY`、沒有 `SERIAL`。Alembic 在 SQLite 上啟用
`render_as_batch` 繞過它不支援的 `ALTER TABLE`。
這三件事是「換連線字串就能換資料庫」能成立的全部前提。

## 規則必須可執行，否則會腐爛

這是本 ADR 最重要的一段。

「`domain/` 不 import 框架」寫在 README 上是一句**宣稱**，而宣稱會腐爛。
某天有人為了趕時間在領域層 import 了 SQLAlchemy，程式照跑、測試照過，
文件上那句話就默默變成謊言，而且沒有任何人會發現。

所以規則寫成 `tests/architecture/test_layering.py`：用 AST 直接讀原始碼，
檢查四件事——

1. `domain/` 只 import 標準函式庫（用 `sys.stdlib_module_names` 判定）
2. 每一層只依賴允許清單裡的內部層，`domain/` 的允許清單是**空的**
3. `repositories/` 不依賴 `services/` 或 `api/`
4. 全專案沒有 SQLAlchemy 的 `text()`（原生 SQL 的入口）

這些測試不驗行為，驗的是結構；它們唯一的價值是在有人越界的那一刻讓 CI 變紅。
驗證方式是實際注入一個違規（在 `domain/money.py` 加上 `import sqlalchemy`），
確認測試真的失敗並印出可讀的訊息，再還原。**沒有驗證過會失敗的測試，
等於沒有測試。**

它擋不住蓄意規避（`importlib.import_module` 就繞過去了），
但這不是防惡意的機制，是防疏忽的機制——而疏忽正是分層腐爛的實際原因。

## 考慮過的其他方案

**不分層，單一 `app/` 套件。** 對這個規模（4,300 行）是合理的選擇，
而且少很多樣板。放棄的理由是對帳邏輯需要在沒有任何外部服務的情況下
被大量隨機輸入測試——那個需求直接推導出「領域層必須是純的」，
而一旦領域層是純的，分層就已經存在了，只是有沒有寫成目錄結構的差別。

**用 Repository 但不做介面，直接依賴 SQLAlchemy 的 Session。**
少一層抽象，實務上很常見。但這樣「換資料庫只要改連線字串」就變回一句無法驗證的話——
而本專案想主張的正是那句話。

**讓 Pydantic model 同時當領域模型與 DTO。** FastAPI 生態的常見做法，
可以省掉一整層轉換。代價是領域模型從此依賴 Pydantic（違反硬規則），
而且 API 的欄位命名與序列化需求會開始影響業務物件的設計。

**用 import-linter 這類現成工具管制分層。** 功能更完整，設定也更簡潔。
選擇自己用 AST 寫的理由是：這幾條規則很短，自己寫可以讓失敗訊息
直接解釋**為什麼**有這條規則（見上面那段 assert 訊息），
而不是印出一行設定檔違規。對一個要拿來說明設計決策的專案，
這個差別是有價值的。

## 後果

**好的：**

- 領域層的測試不需要資料庫，跑完一輪不到 0.2 秒——所以真的會被跑。
- 「SQLite 換 PostgreSQL 只要改連線字串」從宣稱變成**已驗證的事實**：
  同一份 `tests/api` 在記憶體 SQLite 與真實 PostgreSQL 16 上都是 24 passed，
  同一份 Alembic migration 兩邊都跑得過，對帳結果完全一致。
- 新增平台只需要在 `parsers/` 加一個檔案，不動任何既有程式
  （`tests/parsers/test_registry.py` 用執行期才定義的 parser 證明了這點）。
- 架構規則有 CI 守著，不靠 code review 的自律。

**代價：**

- 多一層 DTO 轉換，`api/schemas.py` 有大量看起來像樣板的 `of()` 方法。
  這是真實的成本，只有在 API 與領域模型開始各自演化時才會回本。
- 領域模型與 ORM row 之間要手寫 mapper（`repositories/mappers.py`）。
- 對只想改一個欄位的人來說，要動的檔案比單層架構多。
- 架構測試本身是額外要維護的東西；層一旦重新命名，
  `ALLOWED_INTERNAL_DEPENDENCIES` 就要跟著改。

## 驗證

```bash
python -m pytest tests/architecture                  # 分層規則
python -m pytest tests/domain                        # 領域層，不需要任何外部服務
python -m pytest tests/api                           # 記憶體 SQLite
RECONCILIATION_TEST_DATABASE_URL="postgresql+psycopg://…" \
    python -m pytest tests/api                       # 同一份測試，真的 PostgreSQL 16
```

最後一條是整個 ADR 的關鍵驗證：**可替換性不是被主張的，是被執行過的。**
測試的資料庫 URL 由環境變數決定，沒設就用記憶體 SQLite，
所以這條驗證不需要任何人記得手動去跑——把環境變數設進 CI 就一直有效。
