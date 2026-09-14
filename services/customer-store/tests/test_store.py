from pathlib import Path

import pytest

from timdoc_contracts import Customer, ServiceSettings
from timdoc_customer_store import CustomerStore


def test_customer_is_retrievable_by_inn_and_normalized_name(tmp_path: Path) -> None:
    store = CustomerStore(tmp_path / "timdoc.db")
    customer = Customer(
        name="ООО «Рощинский»",
        inn="0268104130",
        address="Республика Башкортостан",
    )

    store.upsert_customer(customer)

    assert store.find_customer(inn="0268104130") == customer
    assert store.find_customer(name='ооо "рощинский"') == customer


def test_customer_update_does_not_create_duplicate(tmp_path: Path) -> None:
    store = CustomerStore(tmp_path / "timdoc.db")
    store.upsert_customer(Customer(name="ООО «Рощинский»", inn="0268104130"))
    updated = Customer(name="ООО «Рощинский»", inn="0268104130", phone="+7 900 000-00-00")

    store.upsert_customer(updated)

    assert store.list_customers() == [updated]


def test_service_settings_have_defaults_and_persist(tmp_path: Path) -> None:
    store = CustomerStore(tmp_path / "timdoc.db")
    assert store.get_settings().organization == "ООО «Акрос РБ»"

    changed = ServiceSettings(organization="ООО «Другой центр»", employee_name="Иванов И.И.")
    store.save_settings(changed)

    reopened = CustomerStore(tmp_path / "timdoc.db")
    assert reopened.get_settings() == changed


def test_customer_validation_rejects_blank_name_and_bad_inn(tmp_path: Path) -> None:
    store = CustomerStore(tmp_path / "timdoc.db")
    with pytest.raises(ValueError, match="название"):
        store.upsert_customer(Customer(name="   ", inn="0268104130"))
    with pytest.raises(ValueError, match="10 или 12 цифр"):
        store.upsert_customer(Customer(name="ООО «Рощинский»", inn="12345"))
    assert store.list_customers() == []
    assert store.find_customer(name="", inn="") is None
