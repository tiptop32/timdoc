# Document generator

Подготавливает docxtpl-шаблоны, преобразует старые DOC через Microsoft Word и
создает три DOCX на каждую единицу техники вместе с ZIP.

Тесты: `uv run pytest services/document-generator/tests --no-cov`.
Оценка заполнения: `uv run pytest services/document-generator/evals -m eval --no-cov`.
