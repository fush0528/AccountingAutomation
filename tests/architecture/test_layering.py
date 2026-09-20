"""把分層規則變成會失敗的測試。

README 與 ADR 0004 都寫著「``domain/`` 不 import 任何框架」。
那是一句宣稱——宣稱會腐爛。某天有人為了趕時間在領域層 import 了
SQLAlchemy，程式照跑、測試照過，而文件上那句話就默默變成謊言。

所以這裡用 AST 直接讀原始碼，把規則寫成可執行的檢查。這些測試不驗行為，
驗的是**結構**；它們唯一的價值是在有人越界的那一刻讓 CI 變紅。

檢查的是靜態 import，所以規避得了（``importlib.import_module`` 就繞過去了）。
但這不是防惡意的機制，是防疏忽的機制——而疏忽正是分層腐爛的實際原因。
"""

from __future__ import annotations

import ast
import sys
from pathlib import Path

import pytest

PACKAGE_ROOT = Path(__file__).resolve().parents[2] / "src" / "reconciliation"

#: 層與層之間允許的依賴方向。鍵不可以 import 值以外的內部層。
#:
#: domain 的允許清單是空的——它是依賴圖的終點，誰都可以指向它，
#: 它不指向任何人。這正是依賴反轉的具體形狀。
ALLOWED_INTERNAL_DEPENDENCIES: dict[str, set[str]] = {
    "domain": set(),
    "parsers": {"domain"},
    "repositories": {"domain"},
    "services": {"domain", "parsers", "repositories"},
    "api": {"domain", "parsers", "repositories", "services"},
}

LAYERS = tuple(ALLOWED_INTERNAL_DEPENDENCIES)


def python_files(layer: str) -> list[Path]:
    return sorted((PACKAGE_ROOT / layer).rglob("*.py"))


def imported_modules(path: Path) -> set[str]:
    """這個檔案 import 了哪些**頂層**模組名稱。

    相對 import（``from ..domain.models import X``）要另外處理：AST 裡
    ``node.module`` 只有 ``domain.models``，真正的層級資訊在 ``node.level``。
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                found.add(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            found.add(node.module.split(".")[0])
    return found


def internal_layers_used(path: Path) -> set[str]:
    """這個檔案用到了本套件的哪幾層（含相對與絕對 import）。"""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    used: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom):
            if node.level > 0 and node.module:
                # from ..domain.models import X  →  domain
                used.add(node.module.split(".")[0])
            elif node.level == 0 and node.module and node.module.startswith("reconciliation."):
                parts = node.module.split(".")
                if len(parts) > 1:
                    used.add(parts[1])
    return used & set(LAYERS)


class TestDomainHasNoFrameworkDependency:
    """領域層只能依賴標準函式庫。

    這是整個架構唯一的一條硬規則。它成立，領域層的測試就不需要資料庫、
    不需要 HTTP、不需要任何外部服務——實測跑完一輪不到 0.2 秒。
    """

    def test_domain_imports_only_stdlib(self) -> None:
        stdlib = sys.stdlib_module_names
        offenders: list[str] = []
        for path in python_files("domain"):
            for module in imported_modules(path):
                if module not in stdlib:
                    offenders.append(f"{path.relative_to(PACKAGE_ROOT)} → {module}")
        assert not offenders, (
            "domain/ 只能 import 標準函式庫，發現外部依賴：\n  "
            + "\n  ".join(offenders)
            + "\n\n這條規則不是潔癖：領域層一旦依賴框架，就再也無法在沒有那個框架的"
            "情況下測試，而對帳邏輯的正確性正是這個專案的全部價值所在。"
        )

    @pytest.mark.parametrize("framework", ["sqlalchemy", "fastapi", "pydantic", "openpyxl"])
    def test_specific_frameworks_are_absent(self, framework: str) -> None:
        """把幾個最可能被誤用的名字單獨列出來，讓失敗訊息更好讀。"""
        for path in python_files("domain"):
            assert framework not in imported_modules(path), (
                f"{path.relative_to(PACKAGE_ROOT)} import 了 {framework}。"
                f"若領域邏輯需要它提供的東西，正確做法是把那個能力抽成介面，"
                f"由外層注入實作。"
            )


class TestDependenciesPointInwards:
    """依賴只能由外層指向內層，不能反過來。"""

    @pytest.mark.parametrize("layer", LAYERS)
    def test_layer_respects_allowed_dependencies(self, layer: str) -> None:
        allowed = ALLOWED_INTERNAL_DEPENDENCIES[layer] | {layer}
        violations: list[str] = []
        for path in python_files(layer):
            for used in internal_layers_used(path) - allowed:
                violations.append(f"{path.relative_to(PACKAGE_ROOT)} → {used}/")
        assert not violations, (
            f"{layer}/ 只允許依賴 {sorted(allowed - {layer}) or '（無）'}，"
            f"但發現：\n  " + "\n  ".join(violations)
        )

    def test_repositories_do_not_import_services_or_api(self) -> None:
        """持久層不該知道誰在用它。

        這條若被打破，`repositories/` 就不再是可替換的元件——
        換一個實作會連帶影響上層。
        """
        for path in python_files("repositories"):
            used = internal_layers_used(path)
            assert not (used & {"services", "api"}), (
                f"{path.relative_to(PACKAGE_ROOT)} 依賴了上層：{sorted(used & {'services', 'api'})}"
            )


class TestNoRawSql:
    """不寫原生 SQL，否則 SQLite ↔ PostgreSQL 的可替換性無法保證。

    這是 ADR 0003／0004 都提到的前提。``text()`` 是 SQLAlchemy 裡
    「我要自己寫 SQL」的入口，出現它就代表那個保證出現缺口。
    """

    def test_no_sqlalchemy_text_construct(self) -> None:
        offenders: list[str] = []
        for layer in LAYERS:
            for path in python_files(layer):
                tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
                for node in ast.walk(tree):
                    if (
                        isinstance(node, ast.Call)
                        and isinstance(node.func, ast.Name)
                        and node.func.id == "text"
                    ):
                        offenders.append(f"{path.relative_to(PACKAGE_ROOT)}:{node.lineno}")
        assert not offenders, (
            "發現 SQLAlchemy 的 text()（原生 SQL）：\n  "
            + "\n  ".join(offenders)
            + "\n\n原生 SQL 會綁定特定資料庫的方言，讓「換連線字串就能換資料庫」"
            "這個已驗證的性質失效。"
        )
