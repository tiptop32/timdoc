from pathlib import Path

from timdoc_contracts import Customer, EquipmentItem, GenerationRequest, ServiceSettings


def test_generation_request_keeps_explicit_paths() -> None:
    request = GenerationRequest(
        customer=Customer(name="ООО «Рощинский»", inn="0268104130"),
        settings=ServiceSettings(),
        equipment=[EquipmentItem(name="RSM SS-780-13", serial_number="M0S07013002106")],
        document_date="2026-09-07",
        template_directory=Path("templates"),
        output_directory=Path("out"),
    )

    assert request.template_directory == Path("templates")
    assert request.output_directory == Path("out")
