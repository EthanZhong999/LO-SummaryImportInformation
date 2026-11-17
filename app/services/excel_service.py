"""Excel COM服务；对象只在调用线程内存在，两表通过显式映射复制本批最终值。"""
import logging
import math
import re
from time import monotonic, sleep
from decimal import Decimal
from pathlib import Path
from shutil import copy2
from uuid import uuid4

from app.core.models import InvoiceRecord, InvoiceLine, ShipmentData
from app.core.validators import validate_invoice_line, validate_shipment_data, validate_currency

logger = logging.getLogger(__name__)


def _remove_successful_backup(backup: Path) -> None:
    """仅在保存与COM清理均成功后删除本次备份，不扫描历史备份。"""
    try:
        backup.unlink(missing_ok=True)
        logger.info('本次成功批次备份已清理 path=%s', backup)
    except OSError:
        # 文件已保存，删除备份失败不能变成可重试写入，以免重复追加。
        logger.warning('文件已保存，但本次备份无法删除，请手动清理 path=%s', backup, exc_info=True)

FIELD_COLUMNS = {
    'inv_no': (1, 'Inv. No.'), 'po_no': (2, 'PO No.'),
    'material_no': (3, 'Material No.'), 'shipper': (6, 'Shipper'),
    'origin': (7, 'Origin'), 'hs_code': (8, 'HS CODE'),
    'value': (12, 'VALUE'), 'currency': (13, 'Currency'), 'qty': (14, 'Qty'),
    'gross_weight': (15, 'G.W'), 'net_weight': (16, 'Net.W'),
    'package': (18, 'Package'), 'port_of_dispatch': (26, 'Port of Dispatch'),
    'hawb_bl': (29, 'HAWB / BL'),
}
TEXT_FIELDS = {name for name in FIELD_COLUMNS if name not in
               {'value', 'qty', 'gross_weight', 'net_weight', 'package'}}

# 前29列同列；Excel2 AD独有，后六列必须逐项明确映射。
EXCEL2_COLUMN_MAPPING = tuple((c, c) for c in range(1, 30)) + (
    (30, 31), (31, 32), (32, 33), (33, 34), (34, 35), (35, 36))
EXCEL2_RESERVED_COLUMNS = (30, 37, 38, 39, 40)


def validate_excel2_path(path: Path) -> Path:
    path = Path(path)
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('请选择存在的报关信息总表（扩展名.xlsx）。')
    return path.resolve()


def validate_inserted_rows(rows):
    if (not isinstance(rows, list) or not rows or
            any(type(r) is not int or r < 2 or r > 65536 for r in rows) or
            any(b != a + 1 for a, b in zip(rows, rows[1:]))):
        raise ValueError('Excel1 inserted_rows必须是非空、连续递增的有效整数行号。')


def plan_excel2_rows(last_business_row, note_row, count, max_rows=1048576):
    if count < 1 or last_business_row < 1 or note_row <= last_business_row:
        raise ValueError('Excel2业务区或批次数无效。')
    start = last_business_row + 1
    extra = max(0, start + count - note_row)
    if note_row + extra + 1 > max_rows:
        raise ExcelWriteError('Excel2没有足够行数安全下移说明区。')
    return list(range(start, start + count)), extra


def _note_row(sheet):
    # 多个已分析标志同时匹配；缺失/重复标志拒绝，不能猜说明边界。
    area = sheet.Range('A:A')
    hit = area.Find(What='随货单据', After=sheet.Cells(1, 1), LookIn=-4163,
                    LookAt=1, SearchOrder=1, SearchDirection=1, MatchCase=False,
                    SearchFormat=False)
    if hit is None:
        raise ExcelWriteError('Material说明区标志缺失，请检查模板。')
    row = int(hit.Row)
    second = area.Find(What='随货单据', After=hit, LookIn=-4163, LookAt=1,
                       SearchOrder=1, SearchDirection=1, MatchCase=False, SearchFormat=False)
    if (second is None or int(second.Row) != row or sheet.Cells(row, 2).Value2 != '随货单据' or
            sheet.Cells(row, 4).Value2 != '基础数据表' or
            sheet.Cells(row + 1, 16).Value2 != '基础数据表'):
        raise ExcelWriteError('Material说明区结构不匹配，已停止同步。')
    return row


def _note_snapshot(sheet, row):
    result = []
    for r in (row, row + 1):
        cells = []
        for c in range(1, 41):
            cell = sheet.Cells(r, c)
            borders = tuple((cell.Borders(i).LineStyle, cell.Borders(i).Weight,
                             cell.Borders(i).Color) for i in (7, 8, 9, 10))
            style = cell.Style
            style_name = style if isinstance(style, str) else style.Name
            cells.append((cell.Formula, cell.NumberFormat, style_name,
                          cell.Interior.Color, cell.Font.Name, cell.Font.Size,
                          cell.Font.Bold, cell.Font.Color, cell.HorizontalAlignment,
                          cell.VerticalAlignment, cell.WrapText, borders))
        result.append((sheet.Rows(r).RowHeight, sheet.Rows(r).Hidden, tuple(cells)))
    return tuple(result)


def _read_excel1_batch(excel, sheet, rows):
    values = []
    for row in rows:
        batch = []
        for column, _ in EXCEL2_COLUMN_MAPPING:
            cell = sheet.Cells(row, column)
            if cell.MergeCells:
                raise ExcelWriteError(f'Excel1第{row}行存在合并源单元格。')
            if cell.HasFormula:
                # 仅无单元格/函数/外链依赖的数值算式可自行确认并计算。
                # 其他公式需要业务确认依赖，不能相信历史缓存或擅自更新外链。
                if not re.fullmatch(r'=[0-9 .+*/()^%-]+', str(cell.Formula)):
                    raise ExcelWriteError(f'Excel1第{row}行第{column}列公式依赖尚未确认，停止同步。')
                cell.Calculate()
                deadline = monotonic() + 30
                # Range.Calculate同步计算本单元格；全局xlPending(2)可能来自历史外链，
                # 不能因此要求重算历史。仅等待实际进行中的xlCalculating(1)。
                while excel.CalculationState == 1:
                    if monotonic() > deadline:
                        raise ExcelWriteError('Excel1本批公式计算未完成，停止同步。')
                    sleep(.05)
            if excel.WorksheetFunction.IsError(cell):
                raise ExcelWriteError(f'Excel1第{row}行第{column}列存在Excel Error，停止同步。')
            value = cell.Value2
            if value is not None and not isinstance(value, (str, int, float, bool)):
                raise ExcelWriteError('Excel1源值类型无法安全复制。')
            if isinstance(value, float) and not math.isfinite(value):
                raise ExcelWriteError('Excel1源值不是有限数值。')
            batch.append(value)
        hs = batch[7]
        if not isinstance(hs, str) or not re.fullmatch(r'[0-9]{10}', hs):
            raise ExcelWriteError(f'Excel1第{row}行HS Code必须是10位文本最终值。')
        if any(not isinstance(batch[c - 1], str) or not batch[c - 1].strip()
               for c in (1, 2, 3, 29)):
            raise ExcelWriteError(f'Excel1第{row}行标识字段缺失或不是文本。')
        values.append(tuple(batch))
    return values


def _sync_value_matches(expected, actual):
    if expected in (None, ''):
        return actual in (None, '')
    if isinstance(expected, bool):
        return type(actual) is bool and actual == expected
    return written_value_matches(expected, actual)


def sync_excel2_batch(excel1_path: Path, excel2_path: Path, inserted_rows: list[int]) -> list[int]:
    source_path = validate_excel1_path(excel1_path)
    path = validate_excel2_path(excel2_path)
    validate_inserted_rows(inserted_rows)
    if source_path == path:
        raise ValueError('两个Excel路径必须不同。')
    import pythoncom
    import win32com.client

    excel = book1 = book2 = sheet1 = sheet2 = source = target = cell = None
    initialized = save_started = saved = False
    stage = 'COM初始化'
    try:
        pythoncom.CoInitialize()
        initialized = True
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
        stage = '只读打开Excel1'
        book1 = excel.Workbooks.Open(str(source_path), UpdateLinks=0, ReadOnly=True,
                                     AddToMru=False)
        excel.Calculation = -4135
        excel.CalculateBeforeSave = False
        if book1.FileFormat != 56 or not book1.ReadOnly:
            raise ExcelWriteError('Excel1必须以只读方式打开真实.xls文件。')
        sheet1 = book1.Worksheets('shipment record')
        for column, title in FIELD_COLUMNS.values():
            if sheet1.Cells(1, column).Value2 != title:
                raise ExcelWriteError('Excel1标题不匹配。')
        stage = '读取本批最终值'
        values = _read_excel1_batch(excel, sheet1, inserted_rows)
        source_date1904 = bool(book1.Date1904)
        stage = '打开Excel2'
        book2 = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                     IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        if book2.ReadOnly:
            raise ExcelWriteError('报关信息总表当前无法写入，请关闭占用文件的Excel窗口后重试。')
        if book2.FileFormat != 51:
            raise ExcelWriteError('Excel2必须是真实.xlsx文件（FileFormat 51）。')
        # ETD是Excel日期数值；两工作簿日期系统不同时转换，不能复制后偏移1462天。
        if source_date1904 != bool(book2.Date1904):
            offset = 1462 if source_date1904 else -1462
            values = [tuple(value + offset if i == 18 and type(value) in (int, float) else value
                            for i, value in enumerate(batch)) for batch in values]
        sheet2 = book2.Worksheets('Material')
        if sheet2.ProtectContents:
            raise ExcelWriteError('Material处于保护状态。')
        stage = '标题和说明区预检'
        for src, dst in EXCEL2_COLUMN_MAPPING:
            if sheet1.Cells(1, src).Value2 != sheet2.Cells(1, dst).Value2:
                raise ExcelWriteError(f'Excel2第{dst}列标题与映射不匹配。')
        if sheet2.Cells(1, 30).Value2 != '破损/索赔/法检商品':
            raise ExcelWriteError('Excel2 AD独有列标题不匹配。')
        note = _note_row(sheet2)
        before_note = _note_snapshot(sheet2, note)
        area = sheet2.Range(f'A2:C{note - 1}')
        hit = area.Find(What='*', After=sheet2.Cells(2, 1), LookIn=-4123,
                        LookAt=2, SearchOrder=1, SearchDirection=2,
                        MatchCase=False, SearchFormat=False)
        last = int(hit.Row) if hit is not None else 1
        hit = area = None
        rows, extra = plan_excel2_rows(last, note, len(values), int(sheet2.Rows.Count))
        # 映射列必须为空；独有列已有默认值/公式按原样保留。
        for row in range(rows[0], note):
            target = sheet2.Range(sheet2.Cells(row, 1), sheet2.Cells(row, 40))
            if target.MergeCells not in (False, 0) or any(
                    target.Formula[0][dst - 1] not in (None, '') for _, dst in EXCEL2_COLUMN_MAPPING):
                raise ExcelWriteError('Excel2候选业务区存在内容、公式或合并，请检查模板。')
        if sheet2.Range(sheet2.Cells(note, 1), sheet2.Cells(note + 1, 40)).MergeCells not in (False, 0):
            raise ExcelWriteError('说明区存在未支持的合并结构，停止同步。')
        stage = '写前备份'
        backup = path.with_name(f'{path.stem}.backup-{uuid4().hex}{path.suffix}')
        copy2(path, backup)
        logger.info('Excel2备份=%s Excel1=%s inserted_rows=%s', backup, source_path, inserted_rows)
        stage = '安全插行和格式'
        if extra:
            sheet2.Rows(f'{note}:{note + extra - 1}').Insert(Shift=-4121, CopyOrigin=0)
        reserved_before = tuple(tuple(sheet2.Cells(row, c).Formula for c in EXCEL2_RESERVED_COLUMNS)
                                for row in rows)
        # 空白行可能只有默认网格、没有业务表边框；视觉样式沿用上一业务行。
        # 数字格式仍保留目标原格式，随后只对实际文本值设置@保护。
        style_row = last if last >= 2 else rows[0]
        source = sheet2.Range(sheet2.Cells(style_row, 1), sheet2.Cells(style_row, 40))
        height = sheet2.Rows(style_row).RowHeight
        for row, batch in zip(rows, values):
            target = sheet2.Range(sheet2.Cells(row, 1), sheet2.Cells(row, 40))
            if target.MergeCells not in (False, 0) or any(
                    target.Formula[0][dst - 1] not in (None, '') for _, dst in EXCEL2_COLUMN_MAPPING):
                raise ExcelWriteError('Excel2目标行不为空，整批未保存。')
            if row != style_row:
                number_formats = tuple(sheet2.Cells(row, c).NumberFormat for c in range(1, 41))
                source.Copy()
                target.PasteSpecial(Paste=-4122)
                for column, number_format in enumerate(number_formats, 1):
                    sheet2.Cells(row, column).NumberFormat = number_format
            sheet2.Rows(row).RowHeight = height
            for (src, dst), value in zip(EXCEL2_COLUMN_MAPPING, batch):
                cell = sheet2.Cells(row, dst)
                if isinstance(value, str):
                    cell.NumberFormat = '@'  # 所有文本按字面值写入，=开头也不能成为公式。
                elif src == 19 and type(value) in (int, float):
                    cell.NumberFormat = 'yyyy/m/d'
                cell.Value2 = value
        excel.CutCopyMode = False
        if rows[-1] == note + extra - 1:
            # Excel相邻单元格共享可见边界；业务行下边框可能改变说明行上边框。
            # 说明区保护优先，恢复其原上边界，避免格式复制越过业务范围。
            for column, note_cell in enumerate(before_note[0][2], 1):
                line_style, weight, color = note_cell[-1][1]  # xlEdgeTop=8
                border = sheet2.Cells(note + extra, column).Borders(8)
                border.Weight = weight
                border.Color = color
                border.LineStyle = line_style
            border = None
        stage = '整批读回和说明区核对'
        for row, batch, reserved in zip(rows, values, reserved_before):
            for (_, dst), value in zip(EXCEL2_COLUMN_MAPPING, batch):
                cell = sheet2.Cells(row, dst)
                if cell.HasFormula or not _sync_value_matches(value, cell.Value2):
                    raise ExcelWriteError(f'Excel2第{row}行第{dst}列读回不一致，整批未保存。')
            if tuple(sheet2.Cells(row, c).Formula for c in EXCEL2_RESERVED_COLUMNS) != reserved:
                raise ExcelWriteError('Excel2独有列发生变化，整批未保存。')
        if _note_row(sheet2) != note + extra or _note_snapshot(sheet2, note + extra) != before_note:
            raise ExcelWriteError('Excel2说明区下移核对失败，整批未保存。')
        stage = '本票行组合'
        # 插行可能继承上一票的大纲级别；仅重置本批行，首行作为两票之间的分隔。
        # 不改整张表的 SummaryRow，也不全表展开，避免影响用户已有组合方向和折叠状态。
        sheet2.Rows(f'{rows[0]}:{rows[-1]}').OutlineLevel = 1
        # 说明区/空白预留行可能已带大纲级别，或在插行时继承了上一票的级别。
        # Excel会将相邻同级行连成一组；必须在本票末行之外建立明确边界。
        footer_end = note + extra + 1
        sheet2.Rows(f'{rows[-1] + 1}:{footer_end}').OutlineLevel = 1
        sheet2.Rows(f'{rows[-1] + 1}:{footer_end}').Hidden = False
        if len(rows) > 1:
            sheet2.Rows(f'{rows[1]}:{rows[-1]}').Group()
        sheet2.Rows(f'{rows[0]}:{rows[-1]}').Hidden = False
        for index, row in enumerate(rows):
            expected_level = 2 if index > 0 else 1
            if int(sheet2.Rows(row).OutlineLevel) != expected_level or sheet2.Rows(row).Hidden:
                raise ExcelWriteError('总表本票行组合核对失败，整批未保存。')
        for row in range(rows[-1] + 1, footer_end + 1):
            if int(sheet2.Rows(row).OutlineLevel) != 1 or sheet2.Rows(row).Hidden:
                raise ExcelWriteError('总表本票组合越过明细边界，整批未保存。')
        stage = '保存Excel2'
        save_started = True
        book2.Save()
        saved = True
        logger.info('Excel2保存成功 excel2_inserted_rows=%s inserted_note_rows=%s', rows, extra)
    except Exception as exc:
        logger.exception('Excel2失败 stage=%s save_started=%s saved=%s', stage, save_started, saved)
        if save_started:
            raise ExcelWriteError('Excel1已成功；Excel2保存结果不确定，请人工检查文件和备份，勿重复同步。', retry_safe=False) from exc
        if isinstance(exc, (ExcelWriteError, ValueError)):
            raise
        raise ExcelWriteError(f'Excel1保持已成功状态；Excel2在“{stage}”失败，请检查占用、模板或Excel环境，详情见日志。') from exc
    finally:
        cell = target = source = sheet1 = sheet2 = None
        cleanup_failed = False
        for book in (book2, book1):
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception:
                    cleanup_failed = True
                    logger.exception('Excel2同步工作簿关闭失败')
        book = book1 = book2 = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                cleanup_failed = True
                logger.exception('Excel2同步实例退出失败')
        excel = None
        if initialized:
            pythoncom.CoUninitialize()
        if cleanup_failed and saved:
            raise ExcelWriteError('Excel2已保存但实例清理失败，请人工核对，勿重复同步。', retry_safe=False)
    _remove_successful_backup(backup)
    return rows


class ExcelWriteError(RuntimeError):
    """可显示给用户的写入失败；保存结果不确定时禁止自动重试。"""
    def __init__(self, message, *, retry_safe=True):
        super().__init__(message)
        self.retry_safe = retry_safe


def validate_excel1_path(path: Path) -> Path:
    path = Path(path)
    if path.suffix.lower() != '.xls' or not path.is_file():
        raise ValueError('请选择存在的Excel1文件（扩展名.xls）。')
    return path.resolve()


def validate_records(records: list[InvoiceRecord]) -> None:
    if not isinstance(records, list) or not records:
        raise ValueError('没有可写入的明细记录。')
    for r in records:
        if not isinstance(r, InvoiceRecord):
            raise ValueError('每条明细必须是InvoiceRecord。')
        validate_invoice_line(InvoiceLine(r.inv_no, r.po_no, r.material_no, r.origin,
                                         r.hs_code, r.value, r.qty, r.net_weight))
        validate_shipment_data(ShipmentData(r.gross_weight, r.package, r.port_of_dispatch,
                                            r.hawb_bl, r.etd, r.carrier))
        validate_currency(r.currency)
        if not isinstance(r.shipper, str) or not r.shipper.strip():
            raise ValueError('shipper必须是非空文本。')
        for name in FIELD_COLUMNS:
            excel_value(getattr(r, name))


def excel_value(value):
    if isinstance(value, Decimal):
        converted = float(value)
        # Excel数值容量有限；不能悄悄丢失输入有效位或整数精度。
        if (not math.isfinite(converted) or Decimal(str(converted)) != value
                or len(value.normalize().as_tuple().digits) > 15):
            raise ValueError('数值超出Excel可安全保存的精度或范围。')
        return converted
    if type(value) is int and (abs(value) > 999999999999999 or int(float(value)) != value):
        raise ValueError('整数超出Excel可安全保存的精度。')
    return value


def record_output_value(record, name, *, first_row):
    """毛重和件数属于整票，输出时仅首行保留，内部记录仍用于完整校验。"""
    if not first_row and name in {'gross_weight', 'package'}:
        return None
    return getattr(record, name)


def written_value_matches(expected, actual) -> bool:
    if isinstance(expected, str):
        return isinstance(actual, str) and actual == expected
    if isinstance(actual, bool) or not isinstance(actual, (int, float)):
        return False
    return Decimal(str(actual)) == Decimal(str(expected))


def _find_append_row(sheet, count: int) -> int:
    # Find查公式而非显示值，空字符串公式也算占用；格式不会使末行变大。
    area = sheet.Range('A:AI')
    last = area.Find(What='*', After=sheet.Cells(1, 1), LookIn=-4123,
                     LookAt=2, SearchOrder=1, SearchDirection=2, MatchCase=False,
                     SearchFormat=False)
    start = max(2, int(last.Row) + 1 if last is not None else 2)
    last = None
    if start + count - 1 > sheet.Rows.Count:
        raise ExcelWriteError('Excel1没有足够的连续追加行。')
    for row in range(start, start + count):
        target = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, 35))
        if target.MergeCells not in (False, 0):
            raise ExcelWriteError('追加区域存在合并单元格，请检查模板。')
        # 空白单元格Formula通常为''；任何公式（包括=""）或常量拒绝。
        if any(v not in (None, '') for v in target.Formula[0]):
            raise ExcelWriteError('追加区域已有内容或公式，已停止写入。')
    return start


def append_records_to_excel1(workbook_path: Path, records: list[InvoiceRecord]) -> list[int]:
    path = validate_excel1_path(workbook_path)
    validate_records(records)
    import pythoncom
    import win32com.client

    excel = workbook = sheet = source = target = cell = None
    initialized = False
    save_started = saved = False
    stage = 'COM初始化'
    rows = []
    try:
        pythoncom.CoInitialize()
        initialized = True
        stage = '创建独立Excel实例'
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3  # 禁用宏，不运行模板内代码。
        excel.AskToUpdateLinks = False
        stage = '打开工作簿'
        logger.info('Excel1开始 path=%s records=%s', path, len(records))
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                       IgnoreReadOnlyRecommended=True, Notify=False,
                                       AddToMru=False)
        excel.Calculation = -4135  # 本批14列无公式，不重算历史外部依赖。
        excel.CalculateBeforeSave = False
        if workbook.ReadOnly:
            raise ExcelWriteError('Excel文件当前无法写入，请关闭正在占用该文件的Excel窗口后重试。')
        if workbook.FileFormat != 56:
            raise ExcelWriteError('Excel1必须是真实.xls工作簿（FileFormat 56）。')
        stage = '验证主表和标题'
        sheet = workbook.Worksheets('shipment record')
        if sheet.ProtectContents:
            raise ExcelWriteError('shipment record处于保护状态，请检查模板。')
        for column, title in FIELD_COLUMNS.values():
            if sheet.Cells(1, column).Value2 != title:
                raise ExcelWriteError(f'Excel1标题不匹配：第{column}列应为“{title}”。')
        stage = '定位安全追加区'
        start = _find_append_row(sheet, len(records))
        rows = list(range(start, start + len(records)))
        logger.info('Excel1 append_start=%s inserted_rows=%s', start, rows)
        stage = '写前备份'
        backup = path.with_name(f'{path.stem}.backup-{uuid4().hex}{path.suffix}')
        copy2(path, backup)
        logger.info('Excel1备份 path=%s', backup)
        stage = '复制格式及写入'
        source = sheet.Range('A5:AI5')
        height = sheet.Rows(5).RowHeight
        for row, record in zip(rows, records):
            target = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, 35))
            # 仅复制格式，不复制旧常量、公式、外部引用，也不插行移动其他区域。
            if row != 5:
                source.Copy()
                target.PasteSpecial(Paste=-4122)
                sheet.Rows(row).RowHeight = height
            for name, (column, _) in FIELD_COLUMNS.items():
                cell = sheet.Cells(row, column)
                if name in TEXT_FIELDS:
                    cell.NumberFormat = '@'
                cell.Value2 = excel_value(record_output_value(record, name, first_row=row == rows[0]))
        excel.CutCopyMode = False
        stage = '读回核对'
        for row, record in zip(rows, records):
            for name, (column, _) in FIELD_COLUMNS.items():
                cell = sheet.Cells(row, column)
                if cell.HasFormula or not _sync_value_matches(record_output_value(record, name, first_row=row == rows[0]), cell.Value2):
                    raise ExcelWriteError(f'第{row}行{name}读回不一致，整批未保存。')
        stage = '保存'
        save_started = True
        workbook.Save()  # 只在全批核对后保存一次。
        saved = True
        logger.info('Excel1保存成功 inserted_rows=%s', rows)
    except Exception as exc:
        logger.exception('Excel1失败 stage=%s save_started=%s saved=%s', stage, save_started, saved)
        if save_started:
            raise ExcelWriteError('Excel保存结果不确定，请人工检查文件和备份，勿重复追加。', retry_safe=False) from exc
        if isinstance(exc, (ExcelWriteError, ValueError)):
            raise
        raise ExcelWriteError(f'Excel1在“{stage}”失败，请检查文件占用、模板和Excel环境；详情见日志。') from exc
    finally:
        # 先放开子对象；只关闭本服务创建的工作簿和实例，不kill Excel进程。
        cell = target = source = sheet = None
        cleanup_failed = False
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                cleanup_failed = True
                logger.exception('Excel1工作簿关闭失败')
            workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                cleanup_failed = True
                logger.exception('Excel独立实例退出失败')
            excel = None
        if initialized:
            pythoncom.CoUninitialize()
        logger.info('COM清理完成 saved=%s cleanup_failed=%s', saved, cleanup_failed)
        if cleanup_failed and saved:
            raise ExcelWriteError('Excel1已保存，但Excel清理失败；请检查残留窗口，勿重复追加。', retry_safe=False)
    _remove_successful_backup(backup)
    return rows
