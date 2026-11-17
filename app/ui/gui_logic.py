"""GUI纯展示逻辑，不初始化Tk或OCR。"""
from dataclasses import dataclass, field
from pathlib import Path
from app.ui.preview_edit import FIELD_NAMES, edit_records
from app.core.models import InvoiceRecord, ParseDiagnostics
from app.core.validators import low_confidence_evidence, validate_invoice_evidence, validate_shipment_evidence

# 用户本轮明确授权：0.80只作复核提醒，低分不再阻止整批写入。
EXCEL_CONFIDENCE_THRESHOLD: float | None = 0.80


def validate_excel_submission(records, diagnostics, threshold):
    if diagnostics.blockers:
        raise ValueError(diagnostics.blockers[0])
    if diagnostics.issues:
        raise ValueError(f'还有 {len(diagnostics.issues)} 项识别问题，请在预览中修正。')
    if threshold is None:
        raise ValueError('Excel提交置信度阈值尚未确认。')
    waybill_fields = {'gross_weight', 'package', 'port_of_dispatch', 'hawb_bl', 'etd', 'carrier'}
    shipment = ParseDiagnostics(
        [e for e in diagnostics.evidence if e.field_name in waybill_fields],
        manual_fields={key for key in diagnostics.manual_fields if key[1] in waybill_fields})
    invoice = ParseDiagnostics(
        [e for e in diagnostics.evidence if e.field_name not in waybill_fields],
        manual_fields=diagnostics.manual_fields - shipment.manual_fields)
    validate_shipment_evidence(shipment)
    validate_invoice_evidence(invoice, len(records))
    if any(r.etd is not None or r.carrier for r in records):
        # 兼容旧阶段内存模型；Phase11记录必须带新增命名字段与供应商的真实来源。
        names = ({e.field_name for e in diagnostics.evidence if e.line_index is None}
                 | {name for index, name in diagnostics.manual_fields if index is None})
        if not {'etd', 'carrier', 'shipper'} <= names:
            raise ValueError('ETD、承运商或发票供应商来源缺失，请重新识别。')
    low_confidence_evidence(diagnostics, threshold)


def confidence_review_text(records, diagnostics, threshold) -> str:
    rows = []
    for evidence in low_confidence_evidence(diagnostics, threshold):
        index = 0 if evidence.line_index is None else evidence.line_index
        scope = '整票' if evidence.line_index is None else f'第{index + 1}条明细'
        value = getattr(records[index], evidence.field_name)
        rows.append(f'{scope}｜{FIELD_NAMES.get(evidence.field_name, evidence.field_name)}｜识别值：{value}'
                    f'｜置信度：{evidence.confidence:.6f}｜第{evidence.page_number}页')
    if not rows:
        return ''
    return f'以下{len(rows)}项低于{threshold:.2f}，允许写入，请用户自行核对原单据：\n' + '\n'.join(rows)


def validate_pdf_path(text: str) -> Path:
    path = Path(text.strip())
    if not text.strip() or path.suffix.lower() != '.pdf' or not path.is_file():
        raise ValueError('请选择存在的PDF文件（扩展名.pdf）。')
    return path


def record_row(record: InvoiceRecord, *, needs_review: bool = False, first_row: bool = True) -> tuple[str, ...]:
    return ('待复核' if needs_review else '通过', *(str(getattr(record, name)) for name in (
        'inv_no', 'po_no', 'material_no', 'origin', 'hs_code',
        'value', 'currency', 'qty', 'net_weight')),
            str(record.gross_weight) if first_row else '', str(record.package) if first_row else '')


def evidence_text(diagnostics: ParseDiagnostics, line_index: int | None) -> str:
    rows = [f'{e.field_name} | confidence={e.confidence} | 第{e.page_number}页 | box={e.box}'
            for e in diagnostics.evidence if e.line_index == line_index]
    return '\n'.join(rows) or '暂无来源信息'


@dataclass
class PreviewState:
    current_records: list[InvoiceRecord] = field(default_factory=list)
    current_diagnostics: ParseDiagnostics = field(default_factory=ParseDiagnostics)
    current_confirmed: bool = False
    running: bool = False
    excel_running: bool = False
    current_excel_written: bool = False
    excel_retry_safe: bool = True
    inserted_rows: list[int] = field(default_factory=list)
    excel1_written_path: Path | None = None
    current_excel2_written: bool = False
    excel2_inserted_rows: list[int] = field(default_factory=list)
    excel2_retry_safe: bool = True

    original_records: tuple[InvoiceRecord, ...] = ()

    @property
    def can_edit(self) -> bool:
        return bool(self.current_records and not self.busy
                    and not self.current_excel_written and self.excel_retry_safe)

    def apply_edit(self, index, inputs, *, review_field=None) -> bool:
        if not self.can_edit:
            raise ValueError('当前批次无法编辑。')
        diagnostics = self.current_diagnostics
        candidates, edited = edit_records(
            self.current_records, index, inputs, issues=diagnostics.issues,
            blocked=bool(diagnostics.blockers), review_field=review_field)
        corrected = {(index, issue.field_name) for issue in diagnostics.issues
                     if issue.line_index == index and issue.field_name in edited}
        if candidates == self.current_records and not corrected:
            return False
        self.current_records = candidates
        diagnostics.manual_fields.update(corrected)
        diagnostics.issues = [issue for issue in diagnostics.issues
                              if (issue.line_index, issue.field_name) not in corrected]
        self.current_confirmed = False
        return True

    def invalidate(self) -> None:
        self.original_records = ()
        self.current_records = []
        self.current_diagnostics = ParseDiagnostics()
        self.current_confirmed = False
        self.current_excel_written = False
        self.excel_retry_safe = True
        self.inserted_rows = []
        self.excel1_written_path = None
        self.current_excel2_written = False
        self.excel2_inserted_rows = []
        self.excel2_retry_safe = True

    def begin(self) -> bool:
        if self.busy:
            return False
        self.invalidate()
        self.running = True
        return True

    def succeed(self, records, diagnostics) -> None:
        self.original_records = tuple(records)
        self.current_records = list(records)
        self.current_diagnostics = diagnostics
        self.running = False

    def fail(self) -> None:
        self.invalidate()
        self.running = False

    def confirm(self) -> bool:
        if self.busy or not self.current_records or self.current_diagnostics.issues or self.current_diagnostics.blockers:
            return False
        self.current_confirmed = True
        return True

    @property
    def busy(self) -> bool:
        return self.running or self.excel_running

    def can_write(self, path_valid: bool) -> bool:
        return bool(self.current_records and self.current_confirmed and path_valid
                    and not self.busy and not self.current_excel_written and self.excel_retry_safe)

    def begin_excel(self, path_valid: bool) -> bool:
        if not self.can_write(path_valid):
            return False
        self.excel_running = True
        return True

    def finish_excel(self, rows: list[int], path: Path | None = None) -> None:
        self.excel_running = False
        self.current_excel_written = True
        self.inserted_rows = rows
        self.excel1_written_path = path

    def fail_excel(self, retry_safe: bool) -> None:
        self.excel_running = False
        self.excel_retry_safe = retry_safe

    def can_sync(self, path_valid: bool) -> bool:
        return bool(self.current_confirmed and self.current_excel_written and self.inserted_rows
                    and self.excel1_written_path and path_valid and not self.busy
                    and not self.current_excel2_written and self.excel2_retry_safe)

    def begin_sync(self, path_valid: bool) -> bool:
        if not self.can_sync(path_valid):
            return False
        self.excel_running = True
        return True

    def finish_sync(self, rows: list[int]) -> None:
        self.excel_running = False
        self.current_excel2_written = True
        self.excel2_inserted_rows = rows

    def fail_sync(self, retry_safe: bool) -> None:
        self.excel_running = False
        self.excel2_retry_safe = retry_safe
