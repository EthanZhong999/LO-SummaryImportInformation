"""统一业务数据与OCR来源；模型不负责布局解析、标准化或Excel写入。"""
from dataclasses import dataclass, field
from decimal import Decimal
from datetime import date

Quad = tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]]


@dataclass
class ShipmentData:
    gross_weight: Decimal
    package: int
    port_of_dispatch: str
    hawb_bl: str
    etd: date | None = None
    carrier: str = ""


@dataclass
class InvoiceLine:
    inv_no: str
    po_no: str
    material_no: str
    origin: str
    hs_code: str
    value: Decimal
    qty: int
    net_weight: Decimal


@dataclass
class InvoiceRecord:
    inv_no: str
    po_no: str
    material_no: str
    shipper: str
    origin: str
    hs_code: str
    value: Decimal
    currency: str
    qty: int
    gross_weight: Decimal
    net_weight: Decimal
    package: int
    port_of_dispatch: str
    hawb_bl: str
    etd: date | None = None
    carrier: str = ""
    volume_m3: Decimal | None = None

    @classmethod
    def from_parts(cls, shipment: ShipmentData, line: InvoiceLine, *, shipper: str,
                   currency: str) -> "InvoiceRecord":
        """公共值逐条复制，不对重复的毛重和包数求和。"""
        return cls(
            inv_no=line.inv_no, po_no=line.po_no, material_no=line.material_no,
            shipper=shipper, origin=line.origin, hs_code=line.hs_code,
            value=line.value, currency=currency, qty=line.qty,
            gross_weight=shipment.gross_weight, net_weight=line.net_weight,
            package=shipment.package, port_of_dispatch=shipment.port_of_dispatch,
            hawb_bl=shipment.hawb_bl,
            etd=shipment.etd, carrier=shipment.carrier,
        )


@dataclass
class OCRItem:
    text: str
    confidence: float
    page_number: int  # 原PDF页码，从1开始。
    x1: float
    y1: float
    x2: float
    y2: float
    original_box: Quad
    page_width: float
    page_height: float
    # 框与页面尺寸均采用渲染页面像素坐标。


@dataclass
class FieldEvidence:
    field_name: str
    confidence: float
    page_number: int
    box: Quad
    line_index: int | None = None  # 明细从0开始；None表示整票字段。


@dataclass
class RecognitionIssue:
    field_name: str
    message: str
    line_index: int | None = None
    page_number: int | None = None
    candidates: tuple[str, ...] = ()


@dataclass
class ParseDiagnostics:
    evidence: list[FieldEvidence] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    issues: list[RecognitionIssue] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    manual_fields: set[tuple[int | None, str]] = field(default_factory=set)
    shipment_preview: ShipmentData | None = None
