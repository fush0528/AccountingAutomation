.PHONY: help install test lint type check fmt cov clean

help:
	@echo "install  安裝開發相依套件"
	@echo "test     執行測試"
	@echo "lint     ruff 靜態檢查"
	@echo "type     mypy 型別檢查（strict）"
	@echo "fmt      ruff 自動格式化"
	@echo "check    lint + type + test，提交前跑這個"
	@echo "cov      測試並產生覆蓋率報告"

install:
	pip install -e ".[dev]"

test:
	pytest

lint:
	ruff check .

type:
	mypy src

fmt:
	ruff check --fix .
	ruff format .

check: lint type test

cov:
	pytest --cov --cov-report=term-missing --cov-report=html

clean:
	rm -rf .pytest_cache .mypy_cache .ruff_cache htmlcov .coverage
