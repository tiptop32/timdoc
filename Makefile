.PHONY: sync lint test eval run build

sync:
	uv sync --all-packages --extra desktop --extra dev --extra build

lint:
	uv run ruff check .
	uv run ruff format --check .
	uv run mypy src contracts/src services/act-parser/src services/customer-store/src services/document-generator/src
	node --check src/timdoc_app/ui/app.js

test:
	uv run pytest

eval:
	uv run pytest -m eval --no-cov

run:
	uv run timdoc

build:
	uv run pyinstaller packaging/timdoc.spec --clean --noconfirm
