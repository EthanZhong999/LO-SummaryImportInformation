"""预览修正校验本次改动，保留原识别证据。"""
from dataclasses import replace
from datetime import date
from decimal import Decimal, InvalidOperation
from math import isfinite
import re

from app.core.models import InvoiceLine, ShipmentData
from app.core.transforms import (
    normalize_hawb, normalize_hs_code, normalize_material_number, normalize_po_number,
)
from app.core.validators import validate_currency, validate_invoice_line, validate_shipment_data

SHIPMENT_FIELDS = ('hawb_bl', 'etd', 'port_of_dispatch', 'gross_weight', 'package', 'currency')
LINE_FIELDS = tuple(InvoiceLine.__dataclass_fields__)
WAYBILL_FIELDS = set(ShipmentData.__dataclass_fields__)


def edit_records(records, index, inputs, *, issues=(), blocked=False, review_field=None):
    pending = {issue.field_name for issue in issues if issue.line_index == index}
    fields = (*SHIPMENT_FIELDS, *(pending & {'shipper', 'carrier', 'volume_m3'})) if index is None else LINE_FIELDS
    if not inputs.keys() <= set(fields):
        raise ValueError('表单包含不可编辑字段。')
    current = records[0 if index is None else index]
    changes = {}
    for name, raw in inputs.items():
        current_value = getattr(current, name)
        if raw == ('' if current_value is None else str(current_value)) and not (
                name == review_field and name in pending):
            continue
        try:
            if not raw.strip():
                raise ValueError('不能为空。')
            if name in ('gross_weight', 'net_weight', 'value', 'volume_m3'):
                value = Decimal(raw)
                if not value.is_finite():
                    raise ValueError('请输入有限数值。')
                if name == 'volume_m3' and value < 0:
                    raise ValueError('请输入非负的立方米数值。')
            elif name in ('qty', 'package'):
                if not re.fullmatch(r'[+-]?[0-9]+', raw.strip()):
                    raise ValueError('请输入整数。')
                value = int(raw)
            elif name == 'etd':
                value = date.fromisoformat(raw.strip())
                if value.isoformat() != raw.strip():
                    raise ValueError('请使用 YYYY-MM-DD 格式。')
            elif name == 'hawb_bl':
                value = normalize_hawb(raw)
            elif name == 'po_no':
                value = normalize_po_number(raw)
                if not value.strip():
                    raise ValueError('清理后不能为空。')
            elif name == 'material_no':
                value = raw.strip()
                if not re.fullmatch(r'[0-9]{3}\.[0-9]{3}\.[0-9]{3}-[0-9]{2}', value):
                    value = normalize_material_number(raw)
            elif name == 'hs_code':
                value = raw.strip()
                if not re.fullmatch(r'[0-9]{10}', value):
                    value = normalize_hs_code(raw)
            else:
                value = raw
        except (ValueError, InvalidOperation) as exc:
            raise ValueError(f'{FIELD_NAMES[name]}输入无效：{exc}') from exc
        changes[name] = value

    candidates = [
        replace(record, **changes) if index is None or index == i else record
        for i, record in enumerate(records)
    ]
    if blocked or any(issue.line_index != index or issue.field_name not in changes for issue in issues):
        return candidates, changes.keys()
    for record in candidates:
        validate_shipment_data(ShipmentData(**{name: getattr(record, name) for name in WAYBILL_FIELDS}))
        validate_invoice_line(InvoiceLine(**{name: getattr(record, name) for name in LINE_FIELDS}))
        validate_currency(record.currency)
    return candidates, changes.keys()


FIELD_NAMES = {
    'inv_no': '发票号', 'po_no': 'PO 号', 'material_no': '物料号',
    'origin': '原产地', 'hs_code': 'HS 编码', 'value': '金额',
    'qty': '数量', 'net_weight': '净重', 'currency': '币种',
    'gross_weight': '毛重', 'volume_m3': '体积(m³)', 'package': '件数', 'port_of_dispatch': '发运地',
    'hawb_bl': '运单号', 'etd': '预计发运日期（ETD）', 'carrier': '承运商', 'shipper': '供应商',
    'volume_m3': '体积（m³）',
}


def evidence_summary(state, evidence, threshold, technical=False):
    index = 0 if evidence.line_index is None else evidence.line_index
    name = evidence.field_name
    original = getattr(state.original_records[index], name)
    current = getattr(state.current_records[index], name)
    document = '运单' if name in WAYBILL_FIELDS else '发票／装箱单'
    status = '待复核' if evidence.confidence < threshold else '通过'
    text = f'{FIELD_NAMES[name]}｜原识别值：{original}｜{document}第 {evidence.page_number} 页｜{status}'
    if original != current:
        text += f'｜当前使用值：{current}｜已人工修正'
    if technical:
        text += f'\n    {name}｜置信度：{evidence.confidence}｜识别框：{evidence.box}'
    return text

MASTER_FIELD_NAMES = {
    'english_name': '英文品名', 'chinese_name': '中文品名', 'duty_rate': '税率（小数）',
    'certificate': '证件', 'process': '工艺', 'term': '贸易条款', 'mode': '运输方式',
    'forwarder': '货代', 'port_of_arrival': '到达港', 'flight_voyage': '航班／航次',
}


def edit_master_row(original, inputs):
    """保留 Excel 数值税率和文本字段类型，整行通过后交由界面更新快照。"""
    values = original.copy()
    for name, raw in inputs.items():
        label = MASTER_FIELD_NAMES[name]
        text = raw.strip()
        if not text:
            raise ValueError(f'{label}不能为空。')
        if text == str(original[name]):
            continue
        if name == 'duty_rate':
            try:
                rate = float(text)
            except ValueError as exc:
                raise ValueError(f'{label}请输入有限数值。') from exc
            if not isfinite(rate):
                raise ValueError(f'{label}请输入有限数值。')
            values[name] = rate
        else:
            values[name] = text
    return values
