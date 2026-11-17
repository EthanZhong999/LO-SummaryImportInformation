"""标准化后的数据验证及显式低分门禁；不转换或修复业务值。"""
from collections import Counter
from decimal import Decimal
from datetime import date
import math
import re

from app.core.models import FieldEvidence, InvoiceLine, ParseDiagnostics, ShipmentData


class ValidationError(ValueError):
    """数据或来源证据无法通过当前阶段已确认规则。"""


def _text(value: object, name: str) -> None:
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{name}必须是非空文本。")


def _decimal(value: object, name: str) -> None:
    if not isinstance(value, Decimal) or not value.is_finite():
        raise ValidationError(f"{name}必须是有限Decimal数值。")


def _integer(value: object, name: str) -> None:
    # bool继承int，必须显式排除；未确认正数业务规则，不检查大小。
    if not isinstance(value, int) or isinstance(value, bool):
        raise ValidationError(f"{name}必须是整数，不能是布尔值。")


def validate_shipment_data(shipment: ShipmentData) -> None:
    if not isinstance(shipment, ShipmentData):
        raise ValidationError("运单数据必须是ShipmentData。")
    _decimal(shipment.gross_weight, "gross_weight")
    _integer(shipment.package, "package")
    _text(shipment.port_of_dispatch, "port_of_dispatch")
    _text(shipment.hawb_bl, "hawb_bl")
    if shipment.etd is not None and type(shipment.etd) is not date:
        raise ValidationError("etd必须是有效日期。")
    if shipment.carrier:
        _text(shipment.carrier, "carrier")
    if any(c.isspace() for c in shipment.hawb_bl):
        raise ValidationError("hawb_bl仍含空白，请先执行标准化。")


def validate_invoice_line(line: InvoiceLine) -> None:
    if not isinstance(line, InvoiceLine):
        raise ValidationError("明细必须是InvoiceLine。")
    for name in ("inv_no", "po_no", "material_no", "origin", "hs_code"):
        _text(getattr(line, name), name)
    if re.fullmatch(r"[0-9]{3}\.[0-9]{3}\.[0-9]{3}-[0-9]{2}", line.material_no) is None:
        raise ValidationError("material_no必须是XXX.XXX.XXX-XX，请先执行标准化。")
    if re.fullmatch(r"[0-9]{10}", line.hs_code) is None:
        raise ValidationError("hs_code必须是10位ASCII数字文本，请先执行标准化。")
    # ABC//仅去一个尾斜杠后为ABC/，不能在此增加全部尾斜杠禁令。
    _decimal(line.value, "value")
    _decimal(line.net_weight, "net_weight")
    _integer(line.qty, "qty")


def validate_invoice_lines(lines: list[InvoiceLine]) -> None:
    if not isinstance(lines, list) or not lines:
        raise ValidationError("发票必须包含至少一条明细。")
    for index, line in enumerate(lines):
        try:
            validate_invoice_line(line)
        except ValidationError as exc:
            raise ValidationError(f"第{index + 1}条明细：{exc}") from exc


def validate_currency(currency: str) -> None:
    _text(currency, "currency")


def _evidence(diagnostics: ParseDiagnostics) -> list[FieldEvidence]:
    if not isinstance(diagnostics, ParseDiagnostics) or not isinstance(diagnostics.evidence, list):
        raise ValidationError("来源诊断必须是ParseDiagnostics及Evidence列表。")
    for item in diagnostics.evidence:
        if not isinstance(item, FieldEvidence):
            raise ValidationError("来源列表必须包含FieldEvidence。")
        _text(item.field_name, "Evidence.field_name")
        if item.line_index is not None and (type(item.line_index) is not int or item.line_index < 0):
            raise ValidationError("Evidence.line_index必须是非负整数或None。")
    return diagnostics.evidence


def validate_shipment_evidence(diagnostics: ParseDiagnostics) -> None:
    counts = Counter((e.line_index, e.field_name) for e in _evidence(diagnostics))
    required = {(None, name) for name in ("gross_weight", "package", "port_of_dispatch", "hawb_bl")}
    for key in diagnostics.manual_fields:
        counts.setdefault(key, 1)
    if (None, "etd") in counts or (None, "carrier") in counts:
        required |= {(None, "etd"), (None, "carrier")}
    if set(counts) != required or any(counts[key] != 1 for key in required):
        raise ValidationError("运单字段Evidence必须各有一个整票来源，不得缺失、重复或关联错误。")


def validate_invoice_evidence(diagnostics: ParseDiagnostics, line_count: int) -> None:
    if type(line_count) is not int or line_count < 1:
        raise ValidationError("Evidence验证需要正整数明细条数。")
    names = ("inv_no", "po_no", "material_no", "origin", "hs_code", "value", "qty", "net_weight")
    required = {(i, name) for i in range(line_count) for name in names}
    counts = Counter((e.line_index, e.field_name) for e in _evidence(diagnostics))
    for key in diagnostics.manual_fields:
        counts.setdefault(key, 1)
    currency_key = (None, "currency")
    shipper_key = (None, "shipper")
    volume_key = (None, "volume_m3")
    optional = {key for key in (shipper_key, volume_key) if key in counts}
    # 续页允许币种重复来源，Parser已确认文本一致；其余字段当前是单框来源。
    if (set(counts) != required | {currency_key} | optional or counts[currency_key] < 1 or counts[volume_key] > 1
            or any(counts[key] != 1 for key in required)):
        raise ValidationError("发票每行八字段Evidence或整票Currency来源缺失、重复或关联错误。")


def low_confidence_evidence(diagnostics: ParseDiagnostics, threshold: float) -> list[FieldEvidence]:
    if (isinstance(threshold, bool) or not isinstance(threshold, (int, float))
            or not 0 <= threshold <= 1 or not math.isfinite(threshold)):
        raise ValidationError("confidence阈值必须是0到1之间的有限数字，不能是布尔值。")
    # 分数几何合法性由OCR接口保证，此处只执行调用方显式传入的策略。
    return [e for e in _evidence(diagnostics) if e.confidence < threshold]


def validate_confidence(diagnostics: ParseDiagnostics, threshold: float) -> None:
    low = low_confidence_evidence(diagnostics, threshold)
    if low:
        fields = ", ".join(f"line_index={e.line_index}:{e.field_name}" for e in low)
        raise ValidationError(f"存在{len(low)}项低置信度字段（阈值{threshold}）：{fields}")
