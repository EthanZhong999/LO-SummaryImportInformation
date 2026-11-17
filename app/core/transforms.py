"""只执行已确认字符规则；保留原模型及全部来源证据。"""
from dataclasses import replace
import re

from app.core.models import InvoiceLine, ShipmentData
from app.core.validators import ValidationError


def normalize_po_number(value: str) -> str:
    if not isinstance(value, str):
        raise ValidationError("PO必须是文本。")
    cleaned = value.strip()
    # 只删除一个业务尾斜杠，内部字符（包括Tab/换行）保持原样。
    return cleaned[:-1] if cleaned.endswith("/") else cleaned


def normalize_material_number(value: str) -> str:
    cleaned = normalize_po_number(value)
    if re.fullmatch(r"[0-9]{11}", cleaned) is None:
        raise ValidationError("Material清理后必须是11位数字文本，请检查原始值。")
    # 用字符串切片保护前导零，不通过int中转或纠正OCR字符。
    return f"{cleaned[:3]}.{cleaned[3:6]}.{cleaned[6:9]}-{cleaned[9:]}"


def normalize_hs_code(value: str) -> str:
    if not isinstance(value, str):
        raise ValidationError("HS Code必须是文本。")
    cleaned = value.strip()
    if re.fullmatch(r"[0-9]{6}", cleaned) is None:
        raise ValidationError("HS Code原始值必须是6位ASCII数字文本，请检查原始值。")
    # 已确认业务规则只在Transform执行；拒绝10位输入，避免重复标准化。
    return cleaned + "0000"


def normalize_hawb(value: str) -> str:
    if not isinstance(value, str):
        raise ValidationError("HAWB必须是文本。")
    cleaned = "".join(value.split())
    if not cleaned:
        raise ValidationError("HAWB删除空白后不能为空。")
    return cleaned


def normalize_shipment_data(shipment: ShipmentData) -> ShipmentData:
    return replace(shipment, hawb_bl=normalize_hawb(shipment.hawb_bl))


def normalize_invoice_line(line: InvoiceLine) -> InvoiceLine:
    return replace(line, po_no=normalize_po_number(line.po_no),
                   material_no=normalize_material_number(line.material_no),
                   hs_code=normalize_hs_code(line.hs_code))


def normalize_invoice_lines(lines: list[InvoiceLine]) -> list[InvoiceLine]:
    return [normalize_invoice_line(line) for line in lines]
