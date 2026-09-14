# Act parser

Извлекает дату, стороны, технику и заводские номера из текстового PDF Saby.
Не выполняет OCR. Тесты: `uv run pytest services/act-parser/tests --no-cov`.
Оценка реального акта: `uv run pytest services/act-parser/evals -m eval --no-cov`.
