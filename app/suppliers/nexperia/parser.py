"""解析当前 NEXPERIA 运单及双层发票表格；不调用 OCR 或业务标准化。"""
import re
from decimal import Decimal
from datetime import date

from app.core.models import FieldEvidence, InvoiceLine, OCRItem, ParseDiagnostics, RecognitionIssue, ShipmentData


class ParseError(ValueError):
    """文档类型、定位或字段格式无法可靠确认。"""

    def __init__(self, message, *, page_number=None, candidates=()):
        super().__init__(message)
        self.page_number = page_number
        self.candidates = candidates


def _anchor_text(text: str) -> str:
    return " ".join(text.casefold().split()).rstrip(" :")


def _center(item: OCRItem) -> tuple[float, float]:
    return ((item.x1 + item.x2) / (2 * item.page_width),
            (item.y1 + item.y2) / (2 * item.page_height))


def _unique(candidates: list[OCRItem], name: str) -> OCRItem:
    if not candidates:
        raise ParseError(f"缺少{name}，请检查运单文字和字段位置。")
    if len(candidates) != 1:
        raise ParseError(f"{name}存在多个候选，无法唯一确认。")
    return candidates[0]


def _anchor(items: list[OCRItem], text: str) -> OCRItem:
    return _unique([i for i in items if _anchor_text(i.text) == text], text)


def _waybill_pages(items: list[OCRItem]) -> list[int]:
    pages = []
    for number in sorted({i.page_number for i in items}):
        page = [i for i in items if i.page_number == number]
        labels = {_anchor_text(i.text) for i in page}
        if not {"shipper", "receiver", "shipment details"} <= labels:
            continue
        if not labels.intersection({"*waybill doc*", "waybill"}):
            continue
        shippers = [i for i in page if _anchor_text(i.text) == "shipper"]
        receivers = [i for i in page if _anchor_text(i.text) == "receiver"]
        if len(shippers) != 1 or len(receivers) != 1:
            continue
        shipper, receiver = shippers[0], receivers[0]
        # 供应商标识必须在发件地址内，收件方或其他区域的名称不能证明供应商。
        if any(re.search(r"\bnexperia\b", i.text, re.IGNORECASE)
               and _center(i)[0] < 0.5
               and shipper.y2 / shipper.page_height < _center(i)[1]
               < receiver.y1 / receiver.page_height for i in page):
            pages.append(number)
    return pages


def is_waybill(items: list[OCRItem]) -> bool:
    """内容组合判断；仅接受一个目标运单页，不使用文件名。"""
    return len(_waybill_pages(items)) == 1


def _below(items: list[OCRItem], anchor: OCRItem) -> list[OCRItem]:
    # 当前版式的值紧接标签下方；横向限制在标签覆盖的列中。
    left, right = anchor.x1 / anchor.page_width, anchor.x2 / anchor.page_width
    bottom = anchor.y2 / anchor.page_height
    return [i for i in items if i is not anchor
            and left - 0.02 <= _center(i)[0] <= right + 0.02
            and i.y1 / i.page_height >= bottom - 0.005
            and bottom < _center(i)[1] <= bottom + 0.03]


def _read_field(diagnostics, recover, field, candidates, label, convert=lambda text: text,
                *, page_number=None, line_index=None):
    """保留唯一来源及原文，字段失败时由预览记录候选与具体原因。"""
    try:
        source = _unique(candidates, label)
        diagnostics.evidence.append(FieldEvidence(
            field, source.confidence, source.page_number, source.original_box, line_index))
        if not source.text.strip():
            raise ParseError(f"{label}为空。")
        return convert(source.text)
    except ValueError as exc:
        if not recover:
            raise ParseError(str(exc)) from exc
        diagnostics.issues.append(RecognitionIssue(
            field, str(exc), line_index, page_number,
            tuple(i.text for i in candidates)))
        return candidates[0].text if len(candidates) == 1 else ""


def extract_waybill_data(items: list[OCRItem], *, recover=False) -> tuple[ShipmentData, ParseDiagnostics]:
    """确认版式后独立定位各字段，预览模式保留可修正的结果。"""
    pages = _waybill_pages(items)
    if len(pages) != 1:
        raise ParseError("不是唯一的 NEXPERIA Waybill：缺少必要文档锚点或存在多个运单页。")
    page = [i for i in items if i.page_number == pages[0]]
    shipper = _anchor(page, "shipper")
    receiver = _anchor(page, "receiver")
    shipment = _anchor(page, "shipment details")
    if not (_center(shipper)[1] < _center(receiver)[1] < _center(shipment)[1]):
        raise ParseError("运单锚点位置不符合当前支持的版式。")

    diagnostics = ParseDiagnostics()

    def read(field, candidates, label, convert=lambda text: text):
        return _read_field(diagnostics, recover, field, candidates, label, convert,
                           page_number=pages[0])

    top = [i for i in page if _center(i)[0] < 0.5
           and i.y2 / i.page_height <= shipper.y1 / shipper.page_height + 0.005]
    date_pattern = (r"\s*([0-9](?:\s*[0-9]){3}\s*-\s*"
                    r"[0-9]\s*[0-9]\s*-\s*[0-9]\s*[0-9])(?=\s|$)")
    date_sources = [i for i in top if re.match(date_pattern, i.text)]
    if recover and not date_sources:
        date_sources = [i for i in top if re.match(r"\s*[0-9][0-9OIl\s]*-", i.text)]

    def parse_etd(text):
        try:
            match = re.match(date_pattern, text)
            if match:
                return date.fromisoformat("".join(match.group(1).split()))
        except ValueError:
            pass
        raise ParseError("ETD日期格式或日历日期无效。")

    etd = read("etd", date_sources, "ETD", parse_etd)
    carrier = read("carrier", [i for i in top if re.search(r"\bMyDHL\b|\bDHL\b", i.text, re.I)],
                   "承运商DHL", lambda _: "DHL")

    weight_labels = [i for i in page if _anchor_text(i.text) == "cust decl shpt wgt (uom) / dim wgt (uom)"]
    piece_labels = [i for i in page if _anchor_text(i.text) == "pieces"]
    merged_waybills = [(i, re.fullmatch(r"\s*waybill(?:\s*:\s*|\s+)(\S.*?)\s*", i.text, re.I))
                       for i in page]
    merged_values = {id(i): match.group(1) for i, match in merged_waybills if match}
    waybills = [i for i in page if _anchor_text(i.text) == "waybill" or id(i) in merged_values]
    if len(weight_labels) == len(piece_labels) == len(waybills) == 1:
        weight_label, pieces, waybill = weight_labels[0], piece_labels[0], waybills[0]
        if not (_center(shipment)[1] < _center(weight_label)[1] < _center(waybill)[1]
                and abs(_center(weight_label)[1] - _center(pieces)[1]) < 0.02
                and _center(weight_label)[0] < 0.5 < _center(pieces)[0]
                and _center(waybill)[1] > 0.6):
            raise ParseError("运单锚点位置不符合当前支持的版式。")

    def weight_value(text):
        match = re.fullmatch(r"\s*([+-]?[0-9]+(?:\.[0-9]+)?)\s*kg\s*", text, re.I)
        if match is None:
            raise ParseError("Gross Weight格式无法解析，预期数字及kg单位。")
        return Decimal(match.group(1))

    def piece_value(text):
        if re.fullmatch(r"\s*[+-]?[0-9]+\s*", text) is None:
            raise ParseError("Pieces / Package格式无法解析，预期整数。")
        return int(text)

    def below_label(labels, field, label, convert):
        if len(labels) == 1:
            return read(field, _below(page, labels[0]), label, convert)
        return read(field, [], f"{label}（标签缺失或不唯一）", convert)

    weight = below_label(weight_labels, "gross_weight", "Gross Weight", weight_value)
    package = below_label(piece_labels, "package", "Pieces / Package", piece_value)

    # 用文本框左边界判断地址列，长文本跨过页面中线也能识别。
    port_candidates = [
        i for i in page if i.x1 / i.page_width < 0.5
        and i.y1 / i.page_height > shipper.y2 / shipper.page_height
        and 0 < receiver.y1 / receiver.page_height - i.y2 / i.page_height <= 0.025
        and i.text.strip()
    ]
    # 窗口内可能有多行地址，取最靠下的一行并保留 OCR 原文。
    last_row = max((_center(i)[1] for i in port_candidates), default=None)
    port = read("port_of_dispatch", [i for i in port_candidates if _center(i)[1] == last_row],
                "Port of Dispatch")
    if len(waybills) == 1:
        waybill = waybills[0]
        wx, wy = _center(waybill)
        if id(waybill) in merged_values:
            hawb = read("hawb_bl", [waybill], "Waybill号码", lambda _: merged_values[id(waybill)])
        else:
            hawb = read("hawb_bl", [i for i in page if i is not waybill
                                    and i.x1 / i.page_width >= waybill.x2 / waybill.page_width - 0.015
                                    and wx < _center(i)[0] <= waybill.x2 / waybill.page_width + 0.35
                                    and abs(_center(i)[1] - wy) <= 0.012
                                    and i.text.strip()], "Waybill号码")
    else:
        if not recover:
            _unique(waybills, "waybill")
        diagnostics.issues.append(RecognitionIssue(
            "hawb_bl", "运单号标签缺失或存在多个候选。", page_number=pages[0],
            candidates=tuple(i.text for i in waybills)))
        hawb = ""
    return ShipmentData(weight, package, port, hawb, etd, carrier), diagnostics


def extract_invoice_shipper(items: list[OCRItem]) -> tuple[str, list[FieldEvidence]]:
    """只从已确认发票页的左上供应商抬头取值，不取发货地公司或页脚。"""
    sources = []
    for number in sorted({i.page_number for i in items}):
        page = [i for i in items if i.page_number == number]
        if not _invoice_identity(page):
            continue
        candidates = [i for i in page if _center(i)[0] < 0.3
                      and 0.07 < _center(i)[1] < 0.16
                      and re.fullmatch(r"nexperia\s+b\.?v\.?", i.text.strip(), re.I)]
        if len(candidates) != 1:
            raise ParseError("发票供应商缺失或存在多个候选。", page_number=number,
                             candidates=tuple(i.text for i in candidates))
        source = candidates[0]
        sources.append(source)
    if not sources or len({i.text.strip().casefold() for i in sources}) != 1:
        raise ParseError("发票供应商缺失或各页不一致。",
                         candidates=tuple(f"第{i.page_number}页：{i.text}" for i in sources))
    return sources[0].text.strip(), [FieldEvidence("shipper", i.confidence, i.page_number,
                                                 i.original_box) for i in sources]


# 两层表头分别建列：非业务列也保留，避免单价、供应商物料号等串入目标列。
_PRIMARY_HEADERS = (
    ("item", "item"), ("supplier_material", "nexperia material number"),
    ("description", "material description"), ("eccn", "eccn"),
    ("hs_code", "hs code"), ("origin", "country of origin"),
    ("inv_no", "packslip number"), ("qty", "quantity (pc)"),
    ("unit_price", "unit price"), ("value", "amount"),
    ("net_weight", "net weight"),
)
_SECONDARY_HEADERS = (
    ("material_no", "customer part number"), ("invoice_number", "invoice number"),
    ("po_no", "customerorder number"), ("supplier_order", "nexperia order number"),
)
_PRIMARY_FIELDS = {"hs_code", "origin", "inv_no", "qty", "value", "net_weight"}
_LINE_FIELDS = ("inv_no", "po_no", "material_no", "origin", "hs_code",
                "value", "qty", "net_weight")


def _invoice_identity(page: list[OCRItem]) -> bool:
    labels = {_anchor_text(i.text) for i in page}
    return ("consolidated invoice & packing list" in labels
            and "nexperia" in labels)


def _invoice_header(page: list[OCRItem], label: str, reference: OCRItem) -> list[OCRItem]:
    # 表头区域围绕 HS Code，排除地址、正文和页脚中的同名文字。
    nearby = [i for i in page if abs(_center(i)[1] - _center(reference)[1]) <= 0.045]
    complete = [i for i in nearby if _anchor_text(i.text) == label]
    splits = {"packslip number": ("packslip", "number"),
              "country of origin": ("country of", "origin")}
    candidates = [[i] for i in complete]
    if label in splits:
        first, second = splits[label]
        for a in nearby:
            for b in nearby:
                if (_anchor_text(a.text) == first and _anchor_text(b.text) == second
                        and abs(_center(a)[0] - _center(b)[0]) <= 0.025
                        and 0 < _center(b)[1] - _center(a)[1] <= 0.03):
                    candidates.append([a, b])
    if len(candidates) != 1:
        reason = "缺少" if not candidates else "存在多个候选："
        raise ParseError(f"第{reference.page_number}页表头{reason}{label}。")
    return candidates[0]


def _invoice_columns(page: list[OCRItem]):
    reference = _anchor(page, "hs code")
    tiers = []
    for specification in (_PRIMARY_HEADERS, _SECONDARY_HEADERS):
        headers = [(name, _invoice_header(page, label, reference))
                   for name, label in specification]
        centers = [sum(_center(i)[0] for i in group) / len(group)
                   for _, group in headers]
        if any(a >= b for a, b in zip(centers, centers[1:])):
            raise ParseError(f"第{reference.page_number}页表头列顺序不符合当前版式。")
        boundaries = [max(0.0, centers[0] - (centers[1] - centers[0]) / 2)]
        boundaries += [(a + b) / 2 for a, b in zip(centers, centers[1:])]
        boundaries += [min(1.0, centers[-1] + (centers[-1] - centers[-2]) / 2)]
        tiers.append((headers, boundaries))
    primary, secondary = tiers
    primary_y = sum(_center(i)[1] for _, group in primary[0] for i in group) / sum(
        len(group) for _, group in primary[0])
    secondary_y = sum(_center(i)[1] for _, group in secondary[0] for i in group) / sum(
        len(group) for _, group in secondary[0])
    if not 0.012 < secondary_y - primary_y < 0.05:
        raise ParseError(f"第{reference.page_number}页双层表头位置异常。")
    start = max(i.y2 / i.page_height for _, group in secondary[0] for i in group)
    return primary, secondary, start


def _invoice_pages(items: list[OCRItem], *, recover=False) -> list[list[OCRItem]]:
    pages = [[i for i in items if i.page_number == n]
             for n in sorted({i.page_number for i in items})]
    detail_pages = []
    for page in pages:
        labels = {_anchor_text(i.text) for i in page}
        # 包装页虽有同一标题，但没有任何目标明细表头。
        if not labels.intersection({"hs code", "customer part number", "customerorder number"}):
            continue
        if not recover:
            _invoice_columns(page)  # 不把残缺明细页静默当包装页跳过。
        detail_pages.append(page)
    confirmed = (any(_invoice_identity(page) for page in pages) if recover
                 else bool(detail_pages and _invoice_identity(detail_pages[0])))
    if not detail_pages or not confirmed:
        raise ParseError("不是 NEXPERIA Invoice/Packing List：缺少同页标题、供应商或明细表头。")
    return detail_pages


def is_invoice_packing_list(items: list[OCRItem]) -> bool:
    """按内容及双层表头判断，不使用文件名或固定页码。"""
    try:
        _invoice_pages(items)
    except ParseError:
        return False
    return True


def extract_packing_volume(items: list[OCRItem]) -> tuple[Decimal | None, FieldEvidence | None]:
    matches = []
    for label in items:
        if _anchor_text(label.text).replace(" ", "") != "volume(cm3)":
            continue
        x, y = _center(label)
        for total in items:
            if total.page_number != label.page_number or _anchor_text(total.text) != "total":
                continue
            tx, ty = _center(total)
            if tx < x and 0 < ty - y <= 0.045:
                matches.append((label, ty))
    if not matches:
        return None, None
    if len(matches) != 1:
        raise ParseError("包装表存在多个 Total 体积区域，无法唯一确认。",
                         candidates=tuple(f"第{label.page_number}页：{label.text}" for label, _ in matches))
    label, total_y = matches[0]
    _, label_y = _center(label)
    # 只取与 Total 同一行、位于 Volume 标签下方的框，避开下一行的件数合计。
    candidates = [item for item in items if item.page_number == label.page_number
                  and label.x1 <= _center(item)[0] * label.page_width <= label.x2
                  and _center(item)[1] > label_y
                  and abs(_center(item)[1] - total_y) <= 0.008]
    if len(candidates) != 1 or not re.fullmatch(r"(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?",
                                               candidates[0].text.strip()):
        raise ParseError("包装表 Total 的 Volume(CM3) 缺失、格式异常或存在多个候选。请核对并填写立方米数值。",
                         page_number=label.page_number,
                         candidates=tuple(f"{i.text} cm³" for i in candidates))
    item = candidates[0]
    value = Decimal(item.text.strip().replace(",", "")).scaleb(-6)
    return value, FieldEvidence("volume_m3", item.confidence, item.page_number, item.original_box)


def _column_cells(items: list[OCRItem], tier, fields: set[str]) -> list[tuple[str, OCRItem]]:
    headers, boundaries = tier
    result = []
    for item in items:
        x, _ = _center(item)
        for index, (name, _) in enumerate(headers):
            if boundaries[index] <= x < boundaries[index + 1]:
                if name in fields:
                    result.append((name, item))
                break
    return result


def _group_invoice_rows(cells: list[tuple[str, OCRItem]]) -> list[list[tuple[str, OCRItem]]]:
    # 真实上层同一行中心差约0.003；下层偏移约0.031、物料间距约0.072。
    # 与组首比较避免逐个相邻比较导致噪声链把两行连在一起。
    groups = []
    for cell in sorted(cells, key=lambda c: (_center(c[1])[1], _center(c[1])[0])):
        if not groups or _center(cell[1])[1] - _center(groups[-1][0][1])[1] > 0.008:
            groups.append([])
        groups[-1].append(cell)
    return groups


def _invoice_number(text: str, *, integer: bool, name: str):
    number = r"[+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)"
    pattern = number if integer else number + r"(?:\.[0-9]+)?"
    if re.fullmatch(pattern, text.strip()) is None:
        raise ParseError(f"{name}格式无法解析，预期{'整数或合法千位格式' if integer else '十进制数字'}。")
    value = text.strip().replace(",", "")
    return int(value) if integer else Decimal(value)


def extract_invoice_data(items: list[OCRItem], *, recover=False) -> tuple[list[InvoiceLine], str, ParseDiagnostics]:
    """恢复双层明细。预览模式保留已定位的行，记录无法确认的区域。"""
    pages = _invoice_pages(items, recover=recover)
    diagnostics = ParseDiagnostics()
    if recover:
        detail_numbers = {page[0].page_number for page in pages}
        other_numbers = {i.page_number for i in items} - detail_numbers
        for number in sorted(other_numbers):
            labels = {_anchor_text(i.text) for i in items if i.page_number == number}
            if "master carton id" not in labels:
                diagnostics.blockers.append(f"发票第{number}页无法判断是明细页还是包装页。")
    lines = []
    currencies = []
    currency_errors = []
    for page_index, page in enumerate(pages):
        number = page[0].page_number
        try:
            primary, secondary, start = _invoice_columns(page)
        except ParseError as exc:
            if not recover:
                raise
            diagnostics.blockers.append(f"发票第{number}页无法定位明细：{exc}")
            continue
        currency_labels = [i for i in page if _anchor_text(i.text) == "currency"]
        if currency_labels:
            candidates = [i for label in currency_labels for i in page if i is not label
                          and 0 < _center(i)[0] - _center(label)[0] < 0.15
                          and i.x1 / i.page_width >= label.x2 / label.page_width - 0.005
                          and abs(_center(i)[1] - _center(label)[1]) <= 0.012
                          and _center(i)[1] < start]
            if len(currency_labels) == len(candidates) == 1 and candidates[0].text.strip():
                currencies.append(candidates[0])
            else:
                message = f"第{number}页Currency标签或值缺失、为空或存在多个候选。"
                if not recover:
                    raise ParseError(message)
                currency_errors.append(RecognitionIssue(
                    "currency", message, page_number=number,
                    candidates=tuple(i.text for i in candidates)))
        try:
            # 整票Total可能只在最后一张明细页出现；已有逐页Total版式仍然支持。
            totals = [i for i in page if _anchor_text(i.text) == "total"
                      and _center(i)[1] > start
                      and primary[1][6] <= _center(i)[0] < primary[1][7]]
            if totals:
                total = _unique(totals, f"第{number}页表尾Total")
                end = total.y1 / total.page_height
            else:
                # 仅下一物理页已确认有完整明细表头时允许续读，末页缺Total仍报错。
                if (page_index + 1 == len(pages)
                        or pages[page_index + 1][0].page_number != number + 1):
                    raise ParseError(f"缺少第{number}页表尾Total，且无法确认连续明细续页。")
                # 用明确的条款页脚界定本页正文，不能把整页剩余内容当作物料。
                footers = [i for i in page if _center(i)[1] > start
                           and "thisinvoiceissubjectto" in _anchor_text(i.text).replace(" ", "")
                           and "nexperiawebsite" in _anchor_text(i.text).replace(" ", "")]
                footer = _unique(footers, f"第{number}页明细续页条款页脚")
                end = footer.y1 / footer.page_height
            body = [i for i in page if start < _center(i)[1] < end]
            primary_cells = _column_cells(body, primary, _PRIMARY_FIELDS)
            secondary_cells = _column_cells(body, secondary, {"po_no", "material_no"})
            # 右侧列只在上层出现；下层PO、Nexperia Order的X位置
            # 与上层ECCN、Origin重叠，不能把它们作为新的上层行种子。
            seeds = _column_cells(body, primary, {"inv_no", "qty", "unit_price", "value", "net_weight"})
            seed_groups = _group_invoice_rows(seeds)
            groups = [[cell for cell in primary_cells
                       if abs(_center(cell[1])[1] - sum(_center(i)[1] for _, i in seed) / len(seed)) <= 0.008]
                      for seed in seed_groups]
            if not groups or any(not group for group in groups):
                raise ParseError(f"第{number}页没有完整物料明细。")
            # 上层供应商物料号/ECCN与下层目标列也共享X范围，先按层排除。
            row_centers = [sum(_center(i)[1] for _, i in group) / len(group) for group in groups]
            secondary_cells = [(name, i) for name, i in secondary_cells
                               if all(abs(_center(i)[1] - y) > 0.008 for y in row_centers)]
            used_secondary = set()
            for row_number, group in enumerate(groups, 1):
                y = sum(_center(i)[1] for _, i in group) / len(group)
                next_y = min(_center(i)[1] for _, i in groups[row_number]) if row_number < len(groups) else end
                # 下层必须在当前上层之下且在下一物料之前，禁止从上一行填充。
                lower = [(name, i) for name, i in secondary_cells
                         if y + 0.012 < _center(i)[1] < min(y + 0.055, next_y - 0.008)]
                used_secondary.update(id(i) for _, i in lower)
                values = {}
                for name in _LINE_FIELDS:
                    label = f"第{number}页第{row_number}条明细{name}"
                    if name in ("value", "net_weight", "qty"):
                        def convert(text, field=name, context=label):
                            return _invoice_number(text, integer=field == "qty", name=context)
                    else:
                        def convert(text):
                            return text
                    values[name] = _read_field(
                        diagnostics, recover, name, [i for field, i in group + lower if field == name],
                        label, convert, page_number=number, line_index=len(lines))
                lines.append(InvoiceLine(**values))
            if any(id(i) not in used_secondary for _, i in secondary_cells):
                raise ParseError(f"第{number}页存在无法关联到上层物料的下层字段。")
        except ParseError as exc:
            if not recover:
                raise
            diagnostics.blockers.append(str(exc))
    currency_values = {i.text for i in currencies}
    if len(currency_values) != 1 or currency_errors:
        if not recover:
            raise ParseError("缺少Currency标签或值。" if not currencies
                             else "Currency存在多个冲突值，无法确认整票币种。")
        if len(currency_values) != 1:
            currency_errors.append(RecognitionIssue(
                "currency", "缺少Currency标签或值。" if not currencies
                else "Currency存在多个冲突值，无法确认整票币种。",
                page_number=pages[0][0].page_number,
                candidates=tuple(f"第{i.page_number}页：{i.text}" for i in currencies)))
        diagnostics.issues.extend(currency_errors)
        return lines, "", diagnostics
    diagnostics.evidence.extend(FieldEvidence(
        field_name="currency", confidence=i.confidence, page_number=i.page_number,
        box=i.original_box) for i in currencies)
    return lines, currencies[0].text, diagnostics
