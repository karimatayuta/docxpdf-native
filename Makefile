.PHONY: sync test coverage lint format-check typecheck benchmark check

sync:
	uv sync

test:
	uv run pytest

coverage:
	uv run pytest --cov=docxpdf_native --cov-report=term-missing

lint:
	uv run ruff check .

format-check:
	uv run ruff format --check .

typecheck:
	uv run mypy src

benchmark:
	uv run python -m benchmark.validate_reference

check: test coverage lint format-check typecheck
