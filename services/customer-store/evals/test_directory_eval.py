from pathlib import Path

import pytest

from timdoc_contracts import Customer
from timdoc_customer_store import CustomerStore


@pytest.mark.eval
def test_directory_resolves_common_quote_and_case_variants(tmp_path: Path) -> None:
    store = CustomerStore(tmp_path / "timdoc.db")
    expected = Customer(name="ООО «Рощинский»", inn="0268104130")
    store.upsert_customer(expected)

    variants = ["ООО Рощинский", 'ооо "рощинский"', "  ООО   «РОЩИНСКИЙ» "]
    assert [store.find_customer(name=value) for value in variants] == [expected] * 3
