from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import asdict
from pathlib import Path

from timdoc_contracts import Customer, ServiceSettings


class CustomerStore:
    def __init__(self, database_path: Path) -> None:
        self.database_path = Path(database_path)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def upsert_customer(self, customer: Customer) -> Customer:
        normalized_name = _normalize_name(customer.name)
        if not normalized_name:
            raise ValueError("Укажите название хозяйства")
        if customer.inn and not re.fullmatch(r"\d{10}|\d{12}", customer.inn):
            raise ValueError("ИНН должен содержать 10 или 12 цифр")
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO customers (
                    name, normalized_name, inn, address, country, postal_code,
                    region, district, locality, street, house, representative_name,
                    representative_position, phone, email
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(normalized_name) DO UPDATE SET
                    name=excluded.name,
                    inn=excluded.inn,
                    address=excluded.address,
                    country=excluded.country,
                    postal_code=excluded.postal_code,
                    region=excluded.region,
                    district=excluded.district,
                    locality=excluded.locality,
                    street=excluded.street,
                    house=excluded.house,
                    representative_name=excluded.representative_name,
                    representative_position=excluded.representative_position,
                    phone=excluded.phone,
                    email=excluded.email
                """,
                (
                    customer.name.strip(),
                    normalized_name,
                    customer.inn.strip(),
                    customer.address.strip(),
                    customer.country.strip(),
                    customer.postal_code.strip(),
                    customer.region.strip(),
                    customer.district.strip(),
                    customer.locality.strip(),
                    customer.street.strip(),
                    customer.house.strip(),
                    customer.representative_name.strip(),
                    customer.representative_position.strip(),
                    customer.phone.strip(),
                    customer.email.strip(),
                ),
            )
        return customer

    def find_customer(self, *, name: str = "", inn: str = "") -> Customer | None:
        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            row = None
            if inn:
                row = connection.execute(
                    "SELECT * FROM customers WHERE inn = ?", (inn.strip(),)
                ).fetchone()
            if row is None and name:
                row = connection.execute(
                    "SELECT * FROM customers WHERE normalized_name = ?",
                    (_normalize_name(name),),
                ).fetchone()
        return _customer_from_row(row) if row else None

    def list_customers(self) -> list[Customer]:
        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute("SELECT * FROM customers ORDER BY normalized_name").fetchall()
        return [_customer_from_row(row) for row in rows]

    def get_settings(self) -> ServiceSettings:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT value FROM app_settings WHERE key = 'service_settings'"
            ).fetchone()
        if row is None:
            return ServiceSettings()
        stored = json.loads(row[0])
        defaults = asdict(ServiceSettings())
        return ServiceSettings(**(defaults | stored))

    def save_settings(self, settings: ServiceSettings) -> ServiceSettings:
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO app_settings (key, value) VALUES ('service_settings', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (json.dumps(asdict(settings), ensure_ascii=False),),
            )
        return settings

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.database_path, timeout=10)

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS customers (
                    name TEXT NOT NULL,
                    normalized_name TEXT PRIMARY KEY,
                    inn TEXT NOT NULL DEFAULT '',
                    address TEXT NOT NULL DEFAULT '',
                    country TEXT NOT NULL DEFAULT 'РФ',
                    postal_code TEXT NOT NULL DEFAULT '',
                    region TEXT NOT NULL DEFAULT '',
                    district TEXT NOT NULL DEFAULT '',
                    locality TEXT NOT NULL DEFAULT '',
                    street TEXT NOT NULL DEFAULT '',
                    house TEXT NOT NULL DEFAULT '',
                    representative_name TEXT NOT NULL DEFAULT '',
                    representative_position TEXT NOT NULL DEFAULT '',
                    phone TEXT NOT NULL DEFAULT '',
                    email TEXT NOT NULL DEFAULT ''
                )
                """
            )
            connection.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS customers_inn_unique
                ON customers(inn) WHERE inn <> ''
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )


def _normalize_name(value: str) -> str:
    normalized = value.casefold().replace("ё", "е")
    normalized = re.sub(r"[«»\"'.,()]+", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def _customer_from_row(row: sqlite3.Row) -> Customer:
    return Customer(
        name=row["name"],
        inn=row["inn"],
        address=row["address"],
        country=row["country"],
        postal_code=row["postal_code"],
        region=row["region"],
        district=row["district"],
        locality=row["locality"],
        street=row["street"],
        house=row["house"],
        representative_name=row["representative_name"],
        representative_position=row["representative_position"],
        phone=row["phone"],
        email=row["email"],
    )
