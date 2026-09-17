"""產生一組「每種情況各一次」的小樣本結算單。

跟 ``gen_fixtures.py`` 的差別
-----------------------------

``gen_fixtures.py`` 產生的是**統計性**的資料：幾百筆訂單、依機率注入各種
狀況，用來量測匹配率與效能。那份資料沒辦法用肉眼檢查。

這支腳本產生的是**教學性**的資料：12 筆訂單，每一種對帳情況恰好出現一次，
而且每一列都在 ``SAMPLES.md`` 裡註明它是為了示範什麼。用途是：

* 在 Excel 裡打開來看，理解結算單長什麼樣子
* 上傳到網頁介面做示範，結果小到可以逐列講解
* 面試時指著某一列說「這一列是為了測這個情況」

兩者都要有。只有大資料沒辦法解釋，只有小資料沒辦法證明。

執行::

    python scripts/gen_samples.py

輸出在 ``data/samples/``。
"""

from __future__ import annotations

import csv
import sys
import zlib
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from reconciliation.domain import Money  # noqa: E402

OUTPUT_DIR = ROOT / "data" / "samples"

FEE_RATES = {
    "platform_a": Decimal("0.0553"),
    "platform_b": Decimal("0.0480"),
    "platform_c": Decimal("0.0615"),
}
PREFIX = {"platform_a": "SP-", "platform_b": "MO-", "platform_c": "PC-"}


@dataclass
class Order:
    order_id: str
    platform: str
    code: str  # 不含前綴的編號
    ordered_at: datetime
    product: str
    quantity: int
    gross: Money
    scenario: str  # 這筆是為了示範什麼

    @property
    def fee(self) -> Money:
        return self.gross * FEE_RATES[self.platform]

    @property
    def net(self) -> Money:
        return self.gross - self.fee

    @property
    def external(self) -> str:
        return PREFIX[self.platform] + self.code


@dataclass
class Row:
    """結算單上的一列。"""

    external: str  # 已含前綴，可能被刻意改成小寫或空字串
    product: str
    quantity: int
    gross: Money
    fee: Money
    settled_at: date
    kind: str  # sale / refund
    note: str


def money(value: str) -> Money:
    return Money.from_str(value)


# ----------------------------------------------------------------------
def build() -> tuple[list[Order], dict[str, list[Row]], list[tuple[str, str, str]]]:
    """回傳（訂單、各平台的結算列、給文件用的說明表）。"""
    base = datetime(2026, 3, 2, 10, 0)

    orders = [
        # --- 平台 A -------------------------------------------------
        Order("ORD-0001", "platform_a", "A1B2C3D4", base, "無線滑鼠", 1, money("1000"), "完全相符"),
        Order(
            "ORD-0002",
            "platform_a",
            "A2B3C4D5",
            base + timedelta(days=1),
            "機械鍵盤",
            1,
            money("2500"),
            "手續費與我方預期不同",
        ),
        Order(
            "ORD-0003",
            "platform_a",
            "A3B4C5D6",
            base + timedelta(days=2),
            "USB 集線器",
            2,
            money("1360"),
            "部分退款（兩列）",
        ),
        Order(
            "ORD-0004",
            "platform_a",
            "A4B5C6D7",
            base + timedelta(days=3),
            "螢幕支架",
            1,
            money("1200"),
            "平台重複匯出同一列",
        ),
        # --- 平台 B -------------------------------------------------
        Order(
            "ORD-0005",
            "platform_b",
            "B1C2D3E4",
            base,
            "藍牙耳機",
            1,
            money("1800"),
            "結算單把編號寫成小寫",
        ),
        Order(
            "ORD-0006",
            "platform_b",
            "B2C3D4E5",
            base + timedelta(days=1),
            "行動電源",
            1,
            money("890"),
            "平台認列的訂單金額較高（運費補貼）",
        ),
        Order(
            "ORD-0007",
            "platform_b",
            "B3C4D5E6",
            base + timedelta(days=2),
            "網路攝影機",
            1,
            money("2200"),
            "結算單的訂單編號欄是空的",
        ),
        Order(
            "ORD-0008",
            "platform_b",
            "B4C5D6E7",
            base + timedelta(days=3),
            "筆電散熱座",
            1,
            money("760"),
            "平台尚未撥款",
        ),
        # --- 平台 C -------------------------------------------------
        Order(
            "ORD-0009",
            "platform_c",
            "C1D2E3F4",
            base,
            "手機保護殼",
            3,
            money("1170"),
            "編號含前綴且為全形",
        ),
        Order(
            "ORD-0010",
            "platform_c",
            "C2D3E4F5",
            base + timedelta(days=1),
            "充電線組",
            2,
            money("640"),
            "差異小於一元（捨入誤差）",
        ),
        Order(
            "ORD-0011",
            "platform_c",
            "C3D4E5F6",
            base + timedelta(days=2),
            "無線滑鼠",
            1,
            money("980"),
            "編號空白、金額與商品都對不太上，只能列候選",
        ),
        Order(
            "ORD-0012",
            "platform_c",
            "C4D5E6F7",
            base + timedelta(days=3),
            "螢幕支架",
            1,
            money("1500"),
            "退貨全額退款",
        ),
    ]
    by_id = {o.order_id: o for o in orders}
    rows: dict[str, list[Row]] = {p: [] for p in FEE_RATES}
    notes: list[tuple[str, str, str]] = []

    def add(order_id: str, row: Row, expect: str) -> None:
        order = by_id[order_id]
        rows[order.platform].append(row)
        notes.append((order.external, row.note, expect))

    def settled(order: Order, days: int = 13) -> date:
        return (order.ordered_at + timedelta(days=days)).date()

    # === 平台 A ======================================================
    o = by_id["ORD-0001"]
    add(
        "ORD-0001",
        Row(
            o.external,
            o.product,
            o.quantity,
            o.gross,
            o.fee,
            settled(o),
            "sale",
            "編號、金額完全相符",
        ),
        "Stage 1 精確匹配 → 已勾稽",
    )

    # 手續費多收 40 元，訂單金額一致 → 歸因為費率調整
    o = by_id["ORD-0002"]
    add(
        "ORD-0002",
        Row(
            o.external,
            o.product,
            o.quantity,
            o.gross,
            o.fee + money("40"),
            settled(o),
            "sale",
            "手續費比我方預期多 40 元",
        ),
        "Stage 2 容差匹配 → 金額差異，歸因「費率調整」",
    )

    # 銷售 + 退款各一列，同一個訂單編號
    o = by_id["ORD-0003"]
    add(
        "ORD-0003",
        Row(o.external, o.product, o.quantity, o.gross, o.fee, settled(o), "sale", "銷售列"),
        "與下一列合併成同一筆結論",
    )
    half = Decimal("0.5")
    add(
        "ORD-0003",
        Row(
            o.external,
            o.product,
            o.quantity,
            -(o.gross * half),
            # 平台 A 的退款不退手續費——金流商已經收走了。這不是簡化，
            # 是這類對帳檔常見的實際規則，也讓歸因更有東西可講。
            money("0"),
            settled(o, 18),
            "refund",
            "退款寫在銷售列的「退款金額」欄，不另起一列",
        ),
        "同一列拆成兩筆記錄 → 歸因「部分退款」",
    )

    # 完全相同的兩列
    o = by_id["ORD-0004"]
    dup = Row(
        o.external,
        o.product,
        o.quantity,
        o.gross,
        o.fee,
        settled(o),
        "sale",
        "與下一列一模一樣（平台重複匯出）",
    )
    add("ORD-0004", dup, "第二次出現會被 fingerprint 擋下")
    rows["platform_a"].append(dup)

    # === 平台 B ======================================================
    # 編號寫成小寫
    o = by_id["ORD-0005"]
    add(
        "ORD-0005",
        Row(
            o.external.lower(),
            o.product,
            o.quantity,
            o.gross,
            o.fee,
            settled(o),
            "sale",
            "編號被寫成小寫",
        ),
        "正規化後 Stage 1 仍然命中",
    )

    # 平台認列的訂單金額多 60，手續費一致
    o = by_id["ORD-0006"]
    add(
        "ORD-0006",
        Row(
            o.external,
            o.product,
            o.quantity,
            o.gross + money("60"),
            o.fee,
            settled(o),
            "sale",
            "訂單金額比我方多 60 元（運費補貼）",
        ),
        "Stage 2 → 歸因「運費補貼或折扣」",
    )

    # 編號欄空白，但金額與商品都對得上
    o = by_id["ORD-0007"]
    add(
        "ORD-0007",
        Row("", o.product, o.quantity, o.gross, o.fee, settled(o), "sale", "訂單編號欄是空的"),
        "Stage 3 模糊匹配靠金額＋日期＋商品名稱救回",
    )

    # ORD-0008 刻意不產生任何結算列 → 未撥款

    # 平台有、我方無
    ghost = Row(
        "MO-ZZZZ9999",
        "不明商品",
        1,
        money("450"),
        money("450") * FEE_RATES["platform_b"],
        date(2026, 3, 20),
        "sale",
        "我方查無此訂單編號",
    )
    rows["platform_b"].append(ghost)
    notes.append(("MO-ZZZZ9999", ghost.note, "→ 漏記單（平台有、我方無）"))

    # === 平台 C ======================================================
    # 全形編號
    o = by_id["ORD-0009"]
    fullwidth = "".join(chr(ord(c) + 0xFEE0) if "0" <= c <= "z" else c for c in o.external)
    add(
        "ORD-0009",
        Row(fullwidth, o.product, o.quantity, o.gross, o.fee, settled(o), "sale", "編號是全形字元"),
        "NFKC 正規化後 Stage 1 命中",
    )

    # 手續費差 0.5 元 → 捨入誤差
    o = by_id["ORD-0010"]
    add(
        "ORD-0010",
        Row(
            o.external,
            o.product,
            o.quantity,
            o.gross,
            o.fee + money("0.50"),
            settled(o),
            "sale",
            "手續費差 0.5 元",
        ),
        "Stage 2 → 歸因「捨入誤差」，可忽略",
    )

    # 編號空白、金額差 90 元（落在容差 100 內）、商品名稱也不一樣。
    #
    # 這裡的數字是調過的：差異必須**落在容差內**才會被算成候選，
    # 但總分要低於門檻 0.82 才不會被自動採納。差太多（例如 300 元）
    # 會連候選都產不出來，那示範的就變成「完全沒線索」而不是
    # 「有線索但不敢認」——後者才是 Stage 3 真正要處理的情況。
    o = by_id["ORD-0011"]
    add(
        "ORD-0011",
        Row(
            "",
            "藍牙耳機",
            1,
            o.gross + money("90"),
            o.fee,
            settled(o),
            "sale",
            "編號空白，金額差 90 元、商品名稱也對不上",
        ),
        "Stage 3 分數不足 → 漏記單，但附上候選與理由",
    )

    # 全額退款：銷售 + 等額退款
    o = by_id["ORD-0012"]
    add(
        "ORD-0012",
        Row(o.external, o.product, o.quantity, o.gross, o.fee, settled(o), "sale", "銷售列"),
        "與下一列合併",
    )
    add(
        "ORD-0012",
        Row(
            o.external,
            o.product,
            o.quantity,
            -o.gross,
            -o.fee,
            settled(o, 17),
            "refund",
            "全額退款列",
        ),
        "淨撥款為 0 → 歸因「部分退款」",
    )

    return orders, rows, notes


# ----------------------------------------------------------------------
def write_orders(path: Path, orders: list[Order]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            [
                "訂單編號",
                "平台",
                "平台訂單編號",
                "訂單時間",
                "商品名稱",
                "數量",
                "訂單金額",
                "預期手續費",
            ]
        )
        for o in orders:
            writer.writerow(
                [
                    o.order_id,
                    o.platform,
                    o.external,
                    o.ordered_at.isoformat(),
                    o.product,
                    o.quantity,
                    o.gross.to_decimal(),
                    o.fee.to_decimal(),
                ]
            )


#: 三個扣款科目的拆分比例。合計必須是 100，實際金額用 Money.allocate
#: 依最大餘數法分配——三個整數分不盡的那一分錢會落在第一個科目，
#: 而不是憑空多出或少掉。
FEE_SPLIT = (55, 35, 10)
PLATFORM_A_HEADER = (
    "交易日期",
    "廠商訂單編號",
    "金流交易編號",
    "付款方式",
    "手續費率",
    "商品名稱",
    "交易金額",
    "金流手續費",
    "平台手續費",
    "金流處理費",
    "退款日期",
    "退款金額",
    "應收款項(淨額)",
    "結算日期",
    "撥款日期",
)


def merge_platform_a(rows: list[Row]) -> list[tuple[Row, Row | None]]:
    """把退款列併回它的銷售列。

    平台 A 的對帳檔把「賣了又退一部分」壓縮在同一列——銷售列上多一個
    退款金額欄。產生器內部統一用兩個 Row 表示（跟平台 C 一致），
    要寫檔時才在這裡壓縮回去。parser 做的是反向的同一件事。
    """
    merged: list[tuple[Row, Row | None]] = []
    for row in rows:
        if row.kind == "refund":
            for index in range(len(merged) - 1, -1, -1):
                sale, existing = merged[index]
                if sale.external == row.external and existing is None:
                    merged[index] = (sale, row)
                    break
            else:  # pragma: no cover - 樣本資料裡不會發生
                raise ValueError(f"退款列找不到對應的銷售列：{row.external}")
        else:
            merged.append((row, None))
    return merged


def write_platform_a(path: Path, rows: list[Row]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(PLATFORM_A_HEADER)
        for sale, refund in merge_platform_a(rows):
            gateway_fee, platform_fee, handling_fee = sale.fee.allocate(FEE_SPLIT)
            refund_amount = -refund.gross if refund else money("0")
            rate = (
                (sale.fee.to_decimal() / sale.gross.to_decimal() * 100).quantize(
                    Decimal("0.01")
                )
                if not sale.gross.is_zero
                else Decimal("0.00")
            )
            writer.writerow(
                [
                    (sale.settled_at - timedelta(days=13)).isoformat(),
                    sale.external,
                    # 金流端的流水號由內容推導，不用列序——這樣「平台重複
                    # 匯出同一列」產生的兩列才會真的一模一樣。
                    # 用 crc32 而不是內建 hash()：後者每個 process 都不同
                    # （PYTHONHASHSEED 隨機化），產生的檔案就不會固定。
                    f"EC{sale.settled_at:%Y%m%d}"
                    f"{zlib.crc32(sale.external.encode()) % 10000:04d}",
                    "信用卡",
                    f"{rate}%",
                    sale.product,
                    sale.gross.to_decimal(),
                    gateway_fee.to_decimal(),
                    platform_fee.to_decimal(),
                    handling_fee.to_decimal(),
                    refund.settled_at.isoformat() if refund else "",
                    refund_amount.to_decimal(),
                    (sale.gross - sale.fee - refund_amount).to_decimal(),
                    (sale.settled_at - timedelta(days=2)).isoformat(),
                    sale.settled_at.isoformat(),
                ]
            )


def write_platform_b(path: Path, rows: list[Row]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    labels = {"sale": "銷售", "refund": "折讓"}
    wb = Workbook()
    ws = wb.active
    ws.title = "結算明細"

    ws["A1"] = "平台 B 商城　賣家結算明細表"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:H1")
    ws["A2"] = "結算期間：2026/03/01 - 2026/03/31"
    ws.merge_cells("A2:H2")
    # 第 3 列刻意留白

    ws["A4"] = "訂單資訊"
    ws.merge_cells("A4:C4")
    ws["D4"] = "金額明細"
    ws.merge_cells("D4:F4")
    ws["G4"] = "結算"
    ws.merge_cells("G4:H4")
    for cell in ("A4", "D4", "G4"):
        ws[cell].font = Font(bold=True)
        ws[cell].alignment = Alignment(horizontal="center")

    header = ["訂單編號", "商品", "數量", "訂單金額", "手續費", "撥款金額", "結算日", "類別"]
    for col, name in enumerate(header, start=1):
        ws.cell(row=5, column=col, value=name).font = Font(bold=True)

    for offset, r in enumerate(rows, start=6):
        ws.cell(row=offset, column=1, value=r.external)
        ws.cell(row=offset, column=2, value=r.product)
        ws.cell(row=offset, column=3, value=r.quantity)
        ws.cell(row=offset, column=4, value=float(r.gross.to_decimal()))
        ws.cell(row=offset, column=5, value=float(r.fee.to_decimal()))
        ws.cell(row=offset, column=6, value=float((r.gross - r.fee).to_decimal()))
        ws.cell(row=offset, column=7, value=r.settled_at.strftime("%Y/%m/%d"))
        ws.cell(row=offset, column=8, value=labels[r.kind])

    ws.cell(row=len(rows) + 7, column=1, value="※ 本表僅供對帳參考")
    for column, width in zip("ABCDEFGH", (16, 14, 6, 12, 10, 12, 12, 8), strict=True):
        ws.column_dimensions[column].width = width
    wb.save(path)


def write_platform_c(path: Path, rows: list[Row]) -> None:
    labels = {"sale": "銷貨", "refund": "退貨"}

    def roc(d: date) -> str:
        return f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"

    def amount(m: Money) -> str:
        return f"{abs(m.to_decimal()):,}"

    lines = [["訂單序號", "交易別", "品名", "數量", "銷售金額", "服務費", "實撥金額", "撥款日"]]
    for r in rows:
        lines.append(
            [
                r.external,
                labels[r.kind],
                r.product,
                str(r.quantity),
                amount(r.gross),
                amount(r.fee),
                amount(r.gross - r.fee),
                roc(r.settled_at),
            ]
        )
    text = (
        "\r\n".join(",".join(f'"{c}"' if "," in c else c for c in line) for line in lines) + "\r\n"
    )
    path.write_bytes(text.encode("big5", errors="replace"))


def write_guide(path: Path, orders: list[Order], notes: list[tuple[str, str, str]]) -> None:
    lines = [
        "# 小樣本說明",
        "",
        "這組資料的每一列都是為了示範某一種對帳情況。**這是合成資料**——",
        "作者手上沒有真實的結算單，所有金額與編號都是產生的。",
        "欄位規格則取自公開技術文件，出處見",
        "[ADR 0005](../../docs/adr/0005-sample-data-provenance.md)。",
        "",
        "由 `python scripts/gen_samples.py` 產生，內容固定不會變動。",
        "",
        "> **這組資料的自動化率沒有參考價值。**",
        ">",
        "> 12 筆訂單裡刻意塞滿了各種例外，所以 Stage 1 只佔 28.6%。",
        "> 真實情況下絕大多數交易是單純相符的——`gen_fixtures.py` 產生的",
        "> 300 筆資料比例才接近現實，自動化率 79.6%。",
        ">",
        "> 這一組的用途是**逐列講解**，不是量測效能。兩者不要混用。",
        "",
        "## 檔案",
        "",
        "| 檔案 | 現實中誰給的 | 格式 | 主要難點 |",
        "|---|---|---|---|",
        "| `orders.csv` | 我方自己的系統 | UTF-8 CSV | 對帳的另一半，不是上傳的對象 |",
        "| `platform_a_sample.csv` | **第三方金流商**的撥款對帳檔 | UTF-8 CSV "
        "| 手續費拆成三欄要相加；退款寫在銷售列的欄位裡；檔案自帶淨額欄可交叉驗證 |",
        "| `platform_b_sample.xlsx` | **電商平台**賣家後台匯出的報表 | Excel "
        "| 報表標題 ＋ 兩層表頭 ＋ 合併儲存格，表頭列號要用找的；金額是數值型 |",
        "| `platform_c_sample.csv` | **舊 ERP／供應商**的對帳單 | **Big5** CSV "
        "| 民國年日期；金額有千分位且是字串；退貨列印正數要自己轉負 |",
        "",
        "三個平台的欄位名稱、日期格式、編碼、數字寫法、甚至「退款」的表示法",
        "全都不同，但解析出來都是同一種 `SettlementRecord`。",
        "",
        "特別值得看的是退款：平台 A 把它放在銷售列的一個欄位，平台 C 另起一列。",
        "同一件商業事實、兩種表示法——parser 把它們都收斂成「一筆銷售記錄 ＋",
        "一筆退款記錄」，所以下游的歸因邏輯完全不需要知道資料來自哪個平台。",
        "",
        "## 我方訂單",
        "",
        "| 訂單 | 平台 | 平台編號 | 商品 | 訂單金額 | 預期手續費 | 預期實收 | 這筆是為了示範 |",
        "|---|---|---|---|---:|---:|---:|---|",
    ]
    for o in orders:
        lines.append(
            f"| {o.order_id} | {o.platform.removeprefix('platform_').upper()} "
            f"| `{o.external}` | {o.product} | {o.gross.to_decimal()} "
            f"| {o.fee.to_decimal()} | {o.net.to_decimal()} | {o.scenario} |"
        )

    lines += [
        "",
        "## 結算單上的每一列",
        "",
        "| 結算單上的編號 | 這一列的特色 | 預期的對帳結果 |",
        "|---|---|---|",
    ]
    for external, note, expect in notes:
        shown = external if external else "（空白）"
        lines.append(f"| `{shown}` | {note} | {expect} |")

    lines += [
        "",
        "## 怎麼用",
        "",
        "```bash",
        "python scripts/serve.py            # 啟動網頁介面",
        "# 在「上傳」頁把三個 platform_*_sample 檔案傳上去",
        "# 同一份再傳一次，看 was_duplicate 變成 true",
        "# 到「總覽」頁按開始對帳，再到「明細」頁逐列對照上表",
        "```",
        "",
        "資料小到可以在 Excel 裡打開逐列看完——這正是它跟",
        "`gen_fixtures.py` 產生的幾百筆統計資料的差別。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


# ----------------------------------------------------------------------
def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    orders, rows, notes = build()

    write_orders(OUTPUT_DIR / "orders.csv", orders)
    write_platform_a(OUTPUT_DIR / "platform_a_sample.csv", rows["platform_a"])
    write_platform_b(OUTPUT_DIR / "platform_b_sample.xlsx", rows["platform_b"])
    write_platform_c(OUTPUT_DIR / "platform_c_sample.csv", rows["platform_c"])
    write_guide(OUTPUT_DIR / "SAMPLES.md", orders, notes)

    print(f"輸出目錄：{OUTPUT_DIR}\n")
    print(f"  orders.csv                  我方訂單 {len(orders)} 筆")
    for platform, platform_rows in sorted(rows.items()):
        suffix = "xlsx" if platform == "platform_b" else "csv"
        print(f"  {platform}_sample.{suffix:<5}   結算列 {len(platform_rows)} 筆")
    print("  SAMPLES.md                  逐列說明")
    print("\n每一列都在 SAMPLES.md 裡註明它是為了示範什麼情況。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
