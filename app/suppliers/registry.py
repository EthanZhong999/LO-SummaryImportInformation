"""当前MVP供应商静态注册；元数据不修改Parser业务输出。"""
from collections.abc import Callable
from dataclasses import dataclass

from app.core.models import InvoiceLine, OCRItem, ParseDiagnostics, ShipmentData
from app.suppliers.nexperia.parser import extract_invoice_data, extract_waybill_data


@dataclass(frozen=True)
class SupplierDefinition:
    key: str
    display_name: str
    shipper: str
    parse_waybill: Callable[[list[OCRItem]], tuple[ShipmentData, ParseDiagnostics]]
    parse_invoice: Callable[[list[OCRItem]], tuple[list[InvoiceLine], str, ParseDiagnostics]]

    def __post_init__(self) -> None:
        if not isinstance(self.shipper, str) or not self.shipper.strip():
            raise ValueError("供应商Shipper必须是非空文本。")


_SUPPLIERS = {
    "NEXPERIA": SupplierDefinition(
        "NEXPERIA", "NEXPERIA", "Nexperia B.V.",
        extract_waybill_data, extract_invoice_data,
    ),
}


def get_supplier(key: str) -> SupplierDefinition:
    if not isinstance(key, str) or key not in _SUPPLIERS:
        raise ValueError(f"不支持的供应商：{key!r}。")
    return _SUPPLIERS[key]


def list_suppliers() -> list[SupplierDefinition]:
    """返回新列表，调用方不能借列表操作改变注册表。"""
    return list(_SUPPLIERS.values())
