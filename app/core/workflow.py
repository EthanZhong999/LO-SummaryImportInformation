"""同步识别编排，仅返回内存记录及证据；GUI线程由调用方管理。"""
from dataclasses import replace
import logging
from pathlib import Path
from time import perf_counter

from app.core.models import InvoiceLine, InvoiceRecord, OCRItem, ParseDiagnostics, RecognitionIssue, ShipmentData
from app.core.transforms import (
    normalize_hawb, normalize_hs_code, normalize_invoice_lines, normalize_material_number,
    normalize_po_number, normalize_shipment_data,
)
from app.core.validators import (
    validate_confidence, validate_currency, validate_invoice_evidence,
    validate_invoice_lines, validate_shipment_data, validate_shipment_evidence,
)
from app.services.ocr_service import DocumentOCRError, OCRService
from app.suppliers.registry import get_supplier
from app.suppliers.nexperia.parser import ParseError, extract_invoice_shipper, extract_packing_volume


def _compose_records(
    shipment: ShipmentData, lines: list[InvoiceLine], currency: str,
    waybill_diagnostics: ParseDiagnostics, invoice_diagnostics: ParseDiagnostics,
    *, shipper: str, confidence_threshold: float | None = None,
) -> tuple[list[InvoiceRecord], ParseDiagnostics]:
    shipment = normalize_shipment_data(shipment)
    lines = normalize_invoice_lines(lines)
    validate_shipment_data(shipment)
    validate_invoice_lines(lines)
    validate_currency(currency)
    # 标准化文本不替换原始来源；证据仍指向Parser选取的OCR框。
    validate_shipment_evidence(waybill_diagnostics)
    validate_invoice_evidence(invoice_diagnostics, len(lines))
    diagnostics = ParseDiagnostics(
        evidence=waybill_diagnostics.evidence + invoice_diagnostics.evidence,
        warnings=waybill_diagnostics.warnings + invoice_diagnostics.warnings,
    )
    # None只表示本次不启用分数门禁，生产阈值仍需调用方明确决定。
    if confidence_threshold is not None:
        validate_confidence(diagnostics, confidence_threshold)
    records = [InvoiceRecord.from_parts(shipment, line, shipper=shipper, currency=currency)
               for line in lines]
    return records, diagnostics


def build_records_from_ocr(
    supplier_key: str, waybill_items: list[OCRItem], invoice_items: list[OCRItem],
    *, confidence_threshold: float | None = None,
) -> tuple[list[InvoiceRecord], ParseDiagnostics]:
    """Parser→标准化→验证→可选分数门禁→动态记录；失败直接传播。"""
    supplier = get_supplier(supplier_key)
    shipment, waybill_diagnostics = supplier.parse_waybill(waybill_items)
    lines, currency, invoice_diagnostics = supplier.parse_invoice(invoice_items)
    shipper, sources = extract_invoice_shipper(invoice_items)
    invoice_diagnostics.evidence.extend(sources)
    volume, evidence = extract_packing_volume(invoice_items)
    if evidence is not None:
        invoice_diagnostics.evidence.append(evidence)
    records, diagnostics = _compose_records(
        shipment, lines, currency, waybill_diagnostics, invoice_diagnostics,
        shipper=shipper, confidence_threshold=confidence_threshold,
    )
    return [replace(record, volume_m3=volume) for record in records], diagnostics


def recognize_documents(
    supplier_key: str, waybill_pdf: Path, invoice_pdf: Path, ocr_service: OCRService,
    *, confidence_threshold: float | None = None, preview: bool = False,
) -> tuple[list[InvoiceRecord], ParseDiagnostics]:
    """使用调用方已准备的服务，按明确角色识别两份PDF，不自动交换。"""
    get_supplier(supplier_key)  # 未知供应商在耗时OCR开始前拒绝。
    logger = logging.getLogger(__name__)
    blockers = []

    def read_pdf(path, label):
        started = perf_counter()
        try:
            return ocr_service.recognize_pdf(path)
        except DocumentOCRError as exc:
            if not preview:
                raise
            logger.exception("%s OCR未完成", label)
            blockers.append(f"{label}：{exc}")
            return exc.items
        finally:
            logger.info("Workflow%s OCR耗时 seconds=%.4f", label, perf_counter() - started)

    waybill_items = read_pdf(waybill_pdf, "运单")
    invoice_items = read_pdf(invoice_pdf, "发票")
    started = perf_counter()
    if preview:
        records, diagnostics = build_preview_from_ocr(supplier_key, waybill_items, invoice_items)
        diagnostics.blockers[:0] = blockers
    else:
        records, diagnostics = build_records_from_ocr(
            supplier_key, waybill_items, invoice_items, confidence_threshold=confidence_threshold)
    logger.info("Workflow解析标准化验证组合耗时 seconds=%.4f", perf_counter() - started)
    return records, diagnostics


def build_preview_from_ocr(supplier_key, waybill_items, invoice_items):
    """生成待复核草稿。非法字段暂存 OCR 文本，提交前必须人工修正。"""
    supplier = get_supplier(supplier_key)
    diagnostics = ParseDiagnostics()
    try:
        shipment, waybill = supplier.parse_waybill(waybill_items, recover=True)
    except ParseError as exc:
        shipment = ShipmentData("", "", "", "", "", "")
        waybill = ParseDiagnostics(blockers=[str(exc)])
    try:
        lines, currency, invoice = supplier.parse_invoice(invoice_items, recover=True)
    except ParseError as exc:
        lines, currency = [], ""
        invoice = ParseDiagnostics(blockers=[str(exc)])

    diagnostics.evidence = waybill.evidence + invoice.evidence
    diagnostics.issues = waybill.issues + invoice.issues
    diagnostics.blockers = waybill.blockers + invoice.blockers
    diagnostics.warnings = waybill.warnings + invoice.warnings

    def normalize(value, name, transform, index=None):
        if any(issue.field_name == name and issue.line_index == index for issue in diagnostics.issues):
            return value
        try:
            result = transform(value)
            if not result.strip():
                raise ValueError("标准化后不能为空。")
            return result
        except ValueError as exc:
            source = next((e for e in diagnostics.evidence
                           if e.field_name == name and e.line_index == index), None)
            diagnostics.issues.append(RecognitionIssue(
                name, str(exc), index, source.page_number if source else None, (value,)))
            return value

    if not waybill.blockers:
        shipment = replace(shipment, hawb_bl=normalize(shipment.hawb_bl, "hawb_bl", normalize_hawb))
    diagnostics.shipment_preview = shipment
    shipper, volume = "", None
    if lines:
        try:
            shipper, sources = extract_invoice_shipper(invoice_items)
            diagnostics.evidence.extend(sources)
        except ParseError as exc:
            diagnostics.issues.append(RecognitionIssue(
                "shipper", str(exc), page_number=exc.page_number, candidates=exc.candidates))
        try:
            volume, source = extract_packing_volume(invoice_items)
            if source is not None:
                diagnostics.evidence.append(source)
            else:
                diagnostics.issues.append(RecognitionIssue("volume_m3", "缺少包装表 Total 体积，请填写立方米数值。"))
        except ParseError as exc:
            diagnostics.issues.append(RecognitionIssue(
                "volume_m3", str(exc), page_number=exc.page_number, candidates=exc.candidates))
    records = []
    for index, line in enumerate(lines):
        line = replace(line,
                       po_no=normalize(line.po_no, "po_no", normalize_po_number, index),
                       material_no=normalize(line.material_no, "material_no", normalize_material_number, index),
                       hs_code=normalize(line.hs_code, "hs_code", normalize_hs_code, index))
        record = InvoiceRecord.from_parts(shipment, line, shipper=shipper, currency=currency)
        records.append(replace(record, volume_m3=volume))
    return records, diagnostics
