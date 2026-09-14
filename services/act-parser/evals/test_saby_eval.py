from pathlib import Path

import pytest

from timdoc_act_parser import ActParser

REAL_ACT = Path("/Users/tiptop32/my_git_reps/timdoc/doc_templates/Акт.pdf")


@pytest.mark.eval
@pytest.mark.skipif(not REAL_ACT.exists(), reason="real acceptance fixture is outside Git")
def test_real_act_extraction_quality_is_complete_for_document_fields() -> None:
    result = ActParser().parse(REAL_ACT)

    required = [result.customer_name, result.document_date, result.service_center_name]
    assert all(required)
    assert len(result.equipment) == 6
    assert all(item.name and item.serial_number for item in result.equipment)
