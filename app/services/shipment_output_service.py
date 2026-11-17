"""本票输出与主数据读取；只使用Excel COM，原件只读，结果与来源可供GUI复核。"""
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
import logging
import math
from pathlib import Path
import re
from shutil import copyfileobj

from app.services.excel_service import (
    ExcelWriteError, FIELD_COLUMNS, excel_value, record_output_value,
    validate_excel1_path, validate_records, written_value_matches,
)

logger = logging.getLogger(__name__)

# 映射使用业务字段与表头，不依赖红/黑字体推断来源。
MASTER_COLUMNS = {
    'english_name': (4, 'English Name', 4, 'English Name'),
    'chinese_name': (5, 'Chinese Name', 5, 'Chinese Name'),
    'duty_rate': (9, 'Duty Rate', 9, 'Duty Rate'),
    'certificate': (10, '证件', 10, '证件'),
    'process': (11, 'Process', 11, 'Process'),
    'term': (26, 'Term', 23, 'Term'),
    'mode': (27, 'Mode', 24, 'Mode'),
    'forwarder': (28, 'Forwarder', 25, 'Forwarder'),
    'port_of_arrival': (30, 'Port of Arrival', 27, 'Port of Arrival'),
    'flight_voyage': (31, 'Flight/Voyage', 28, 'Flight/Voyage'),
}


@dataclass
class MasterDataPreview:
    path: Path
    material_numbers: list[str]
    rows: list[dict]
    warnings: list[str]
    sources: list[str]
    original_rows: list[dict] | None = None
    duplicate_material_warnings: list[str] = field(default_factory=list)


@dataclass
class ShipmentOutput:
    path: Path
    inserted_rows: list[int]


def validate_master_path(path):
    path = Path(path)
    if path.suffix.lower() not in {'.xls', '.xlsx'} or not path.is_file():
        raise ValueError('请选择存在的物料主数据Excel文件（.xls或.xlsx）。')
    return path.resolve()


def validate_output_root(path):
    path = Path(path)
    if not path.is_dir():
        raise ValueError('请选择存在的输出根目录。')
    return path.resolve()


def find_supplier_image(directory, supplier):
    if not directory or not Path(directory).is_dir():
        raise ValueError('请在Excel设置中选择存在的供应商图片文件夹。')
    matches = [
        path for path in Path(directory).iterdir()
        if path.stem.casefold() == supplier.casefold()
        and path.suffix.lower() in {'.png', '.jpg', '.jpeg'} and path.is_file()
    ]
    if len(matches) != 1:
        raise ValueError(f'图片文件夹中必须有且仅有一张以 {supplier} 命名的 PNG 或 JPEG 图片。')
    return matches[0].resolve()


def _signature_cell(sheet, sheet_name):
    used = sheet.UsedRange
    values = used.Value2
    if not isinstance(values, tuple):
        values = ((values,),)
    matches = []
    for row, cells in enumerate(values, used.Row):
        for column, value in enumerate(cells, used.Column):
            if not isinstance(value, str):
                continue
            label = ' '.join(value.split()).casefold()
            if sheet_name == 'Contract':
                match = column == 1 and label.rstrip(' :：') == '卖方'
            else:
                match = re.match(r'signature of authorised\b', label) is not None
            if match:
                matches.append((row, column))
    if not matches:
        raise ValueError(f'{sheet_name} 找不到约定的供应商签名标签。')
    if sheet_name != 'Contract' and len(matches) != 1:
        raise ValueError(f'{sheet_name} 有多个签名标签，无法确定图片位置。')
    return sheet.Cells(*matches[-1])


def _insert_supplier_images(book, image):
    points_per_cm = 72 / 2.54
    width, height = 5.19 * points_per_cm, 1.27 * points_per_cm
    anchors = []
    for name in ('Contract', 'CI', 'PL'):
        sheet = book.Worksheets(name)
        label = _signature_cell(sheet, name)
        area = label.MergeArea
        if name == 'Contract':
            left = sheet.Cells(label.Row, 2).Left
            top = label.Top + label.Height - height
        else:
            anchor = sheet.Cells(area.Row + area.Rows.Count, 4)
            left, top = anchor.Left, anchor.Top + 3
        anchors.append((sheet, left, top))
    for sheet, left, top in anchors:
        # 禁用链接并随工作簿保存，接收者无需原图片文件。
        picture = sheet.Shapes.AddPicture(str(image), 0, -1, left, top, width, height)
        picture.Rotation = 0
        picture.Placement = 2  # 随单元格移动，保持确认的物理尺寸。


def _safe_component(value):
    # 不替换运单号字符以免不同号码映射同名；非法Windows路径要求人工处理。
    if (not isinstance(value, str) or not value or re.search(r'[<>:"/\\|?*\x00-\x1f]', value)
            or value.endswith((' ', '.')) or len(value) > 160
            or re.fullmatch(r'(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?', value)):
        raise ValueError('命名字段含Windows不支持的字符、保留名称或长度，请核对单据。')
    return value


def plan_shipment_path(root, records):
    validate_records(records)
    first = records[0]
    if type(first.etd) is not date or first.carrier != 'DHL':
        raise ValueError('生成子表必须有有效ETD和已识别的DHL承运商。')
    if not re.fullmatch(r'nexperia\s+b\.?v\.?', first.shipper, re.I):
        raise ValueError('发票供应商无法映射到当前支持的NEXPERIA。')
    common = (first.hawb_bl, first.etd, first.carrier, first.shipper)
    if any((r.hawb_bl, r.etd, r.carrier, r.shipper) != common for r in records):
        raise ValueError('同一任务的运单号、ETD、承运商及供应商必须一致。')
    number = _safe_component(first.hawb_bl)
    folder = _safe_component(f'NEXPERIA {first.carrier}-{number} ETD{first.etd:%Y%m%d}')
    filename = _safe_component(f'Contract&Inv&PL-{number}.xls')
    return validate_output_root(root) / folder / filename


def shipment_volume(records):
    volume = records[0].volume_m3
    if (not isinstance(volume, Decimal) or not volume.is_finite() or volume < 0
            or any(record.volume_m3 != volume for record in records)):
        raise ValueError('包装表 Total 体积缺失或不一致，请检查第二份 PDF 并重新识别。')
    return volume


def match_master_rows(records, source_rows, path):
    """按物料号取第一条，保留全部发票明细；重复主数据只提醒，不修改原表。"""
    index = {}
    warnings = []
    duplicate_warnings = []
    for row_number, values in source_rows:
        material = values.get('material_no')
        if material in (None, ''):
            continue
        if not isinstance(material, str):
            raise ValueError(f'HS第{row_number}行物料号必须是文本。')
        material = material.strip()
        if not material:
            continue
        if material in index:
            first_row, first = index[material]
            fields = [name for name in MASTER_COLUMNS if first.get(name) != values.get(name)]
            fields += [name for name in ('shipper', 'origin', 'hs_code')
                       if first.get(name) != values.get(name)]
            detail = '；使用字段有差异：' + '、'.join(fields) if fields else '；使用字段一致'
            warnings.append(f'主数据物料{material}重复：使用HS第{first_row}行，保留第{row_number}行原样{detail}。')
            duplicate_warnings.append(warnings[-1])
        else:
            index[material] = (row_number, values)
    output, sources = [], []
    for record in records:
        if record.material_no not in index:
            raise ValueError(f'主数据HS找不到物料：{record.material_no}。')
        row_number, values = index[record.material_no]
        selected = {}
        for name in MASTER_COLUMNS:
            value = values.get(name)
            if isinstance(value, str):
                value = value.strip()
            if value is None or value == '' or isinstance(value, bool):
                raise ValueError(f'主数据HS第{row_number}行{name}为空或类型不正确。')
            if name == 'duty_rate':
                if type(value) not in (int, float) or not math.isfinite(value):
                    raise ValueError(f'主数据HS第{row_number}行税率必须为有限数值。')
            elif not isinstance(value, str):
                raise ValueError(f'主数据HS第{row_number}行{name}必须为文本。')
            selected[name] = value
        for field in ('shipper', 'origin', 'hs_code'):
            other = values.get(field)
            # Excel可能以数值保存HS；比较时只还原完整整数表示，不补零或改写源数据。
            if field == 'hs_code' and type(other) in (int, float) and math.isfinite(other) and other == int(other):
                other = str(int(other))
            if other not in (None, '') and str(other).strip().casefold() != str(getattr(record, field)).strip().casefold():
                warnings.append(f'物料{record.material_no}的{field}与HS第{row_number}行不同：'
                                f'保留PDF值{getattr(record, field)}；主数据为{other}。')
        output.append(selected)
        sources.append(f'{record.material_no} ← {path} / HS第{row_number}行')
    return MasterDataPreview(Path(path), [r.material_no for r in records], output,
                             list(dict.fromkeys(warnings)), sources,
                             duplicate_material_warnings=list(dict.fromkeys(duplicate_warnings)))


def _new_excel():
    import win32com.client
    excel = win32com.client.DispatchEx('Excel.Application')
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
    except Exception:
        excel.Quit()
        raise
    return excel


def read_master_data(path, records):
    path = validate_master_path(path)
    validate_records(records)
    import pythoncom
    excel = book = sheet = cell = None
    pythoncom.CoInitialize()
    try:
        excel = _new_excel()
        book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        excel.Calculation = -4135
        excel.CalculateBeforeSave = False
        sheet = book.Worksheets('HS')
        headers = {3: 'Material No.', 6: 'Shipper', 7: 'Origin', 8: 'HS CODE'}
        headers.update({src: title for src, title, _, _ in MASTER_COLUMNS.values()})
        for column, title in headers.items():
            if sheet.Cells(1, column).Value2 != title:
                raise ValueError(f'主数据HS第{column}列表头应为“{title}”。')
        hit = sheet.Range('C:C').Find(What='*', After=sheet.Cells(1, 3), LookIn=-4123,
                                     LookAt=2, SearchOrder=1, SearchDirection=2,
                                     MatchCase=False, SearchFormat=False)
        last = int(hit.Row) if hit is not None else 1
        hit = None
        if last > 100000:
            raise ValueError('主数据行数超出本版本读取范围（100000行），请核对HS表。')
        required = {r.material_no for r in records}
        raw = []
        columns = {'material_no': 3, 'shipper': 6, 'origin': 7, 'hs_code': 8}
        columns.update({name: spec[0] for name, spec in MASTER_COLUMNS.items()})
        # 仅检查本票需要的行与授权字段，忽略未使用的净重/pcs错误公式。
        keys = sheet.Range(f'C2:C{last}').Value2 if last >= 2 else ()
        for row_number, (key,) in enumerate(keys, 2):
            if not isinstance(key, str) or key.strip() not in required:
                continue
            values = {}
            for name, column in columns.items():
                cell = sheet.Cells(row_number, column)
                if cell.HasFormula or excel.WorksheetFunction.IsError(cell):
                    raise ValueError(f'主数据HS第{row_number}行{name}含公式或错误，需先确认依赖并提供有效最终值。')
                values[name] = cell.Value2
            raw.append((row_number, values))
        return match_master_rows(records, raw, path)
    except (ValueError, ExcelWriteError):
        raise
    except Exception as exc:
        logger.exception('只读主数据失败 path=%s', path)
        raise ExcelWriteError('读取主数据失败，请检查HS工作表、文件和Excel环境；详情见日志。') from exc
    finally:
        cell = sheet = None
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            book = None
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                excel = None
                pythoncom.CoUninitialize()


def create_shipment_workbook(template, root, records, master, image=None, prepare_documents=False):
    template = validate_excel1_path(template)
    path = plan_shipment_path(root, records)
    volume = shipment_volume(records) if prepare_documents else None
    if image is not None and not image.is_file():
        raise ValueError('供应商图片已不存在，请重新选择图片文件夹。')
    if (not isinstance(master, MasterDataPreview) or master.material_numbers != [r.material_no for r in records]
            or len(master.rows) != len(records)):
        raise ValueError('主数据预览与本票不一致，请重新核对并确认。')
    if any(set(extra) != set(MASTER_COLUMNS) for extra in master.rows):
        raise ValueError('主数据预览字段不完整，请重新核对。')
    if path.parent.exists():
        raise ValueError(f'本票文件夹已存在，禁止覆盖或重复生成：{path.parent}')
    import pythoncom
    excel = book = sheet = source = target = cell = None
    created = saved = False
    rows = list(range(2, 2 + len(records)))
    if rows[-1] > 65536:
        raise ValueError('本票明细超过.xls行数上限。')
    pythoncom.CoInitialize()
    try:
        path.parent.mkdir()  # 不使用exist_ok，防止另一个任务抢占同名目录。
        created = True
        with template.open('rb') as src, path.open('xb') as dst:
            copyfileobj(src, dst)
        excel = _new_excel()
        book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                    IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        excel.Calculation = -4135
        excel.CalculateBeforeSave = False
        if book.ReadOnly or book.FileFormat != 56:
            raise ValueError('子表副本不可写或不是真实.xls文件。')
        if [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)] != [
                'shipment record', 'Contract', 'CI', 'PL', '申报要素']:
            raise ValueError('模板必须包含约定的5张Sheet及其顺序。')
        sheet = book.Worksheets('shipment record')
        if sheet.ProtectContents:
            raise ValueError('模板shipment record处于保护状态。')
        headers = dict(FIELD_COLUMNS.values())
        headers[19] = 'ETD'
        headers.update({dst: title for _, _, dst, title in MASTER_COLUMNS.values()})
        for column, title in headers.items():
            if sheet.Cells(1, column).Value2 != title:
                raise ValueError(f'模板第{column}列表头应为“{title}”。')
        hit = sheet.Range('A:AI').Find(What='*', After=sheet.Cells(1, 1), LookIn=-4123,
                                     LookAt=2, SearchOrder=1, SearchDirection=2,
                                     MatchCase=False, SearchFormat=False)
        last = max(5, rows[-1], int(hit.Row) if hit is not None else 1)
        hit = None
        area = sheet.Range(f'A2:AI{last}')
        if area.MergeCells not in (False, 0):
            raise ValueError('主表数据区有合并单元格，无法安全替换为本票数据。')
        # 副本的主表只容纳本票；ClearContents保留格式，不删行、不移动后四表引用。
        area.ClearContents()
        area = None
        source = sheet.Range('A5:AI5')
        height = sheet.Rows(5).RowHeight
        expected = []
        for row, record, extra in zip(rows, records, master.rows):
            target = sheet.Range(f'A{row}:AI{row}')
            if row != 5:
                source.Copy()
                target.PasteSpecial(Paste=-4122)
            sheet.Rows(row).RowHeight = height
            values = {column: excel_value(record_output_value(record, name, first_row=row == rows[0]))
                      for name, (column, _) in FIELD_COLUMNS.items()}
            values.update({MASTER_COLUMNS[name][2]: value for name, value in extra.items()})
            # 显式日期序列，避免COM把Python日期当字符串或受地区格式影响。
            epoch = date(1904, 1, 1) if book.Date1904 else date(1899, 12, 30)
            values[19] = (record.etd - epoch).days
            for column, value in values.items():
                cell = sheet.Cells(row, column)
                if isinstance(value, str):
                    cell.NumberFormat = '@'
                elif column == 19:
                    cell.NumberFormat = 'yyyy/m/d'
                cell.Value2 = value
            expected.append(values)
        excel.CutCopyMode = False
        for row, values in zip(rows, expected):
            for column in range(1, 36):
                cell = sheet.Cells(row, column)
                value = values.get(column)
                if cell.HasFormula or (cell.Value2 not in (None, '') if value is None
                                       else not written_value_matches(value, cell.Value2)):
                    raise ValueError(f'子表第{row}行第{column}列读回不一致。')
        if prepare_documents:
            from app.services.shipment_pdf_service import prepare_shipment_documents

            prepare_shipment_documents(excel, book, master.path, master.material_numbers, records[0].hawb_bl, volume)
        if image is not None:
            _insert_supplier_images(book, image)
        book.Save()
        saved = True
        logger.info('本票子表保存 path=%s inserted_rows=%s', path, rows)
    except Exception as exc:
        logger.exception('生成本票子表失败 path=%s saved=%s', path, saved)
        if created:
            reason = f'{exc} ' if isinstance(exc, ValueError) else ''
            raise ExcelWriteError(f'{reason}子表生成未完整完成，请核对并保留现场：{path}。'
                                  '文件夹可能含模板副本或已保存结果，勿直接重复生成。', retry_safe=False) from exc
        raise
    finally:
        cell = target = source = sheet = None
        cleanup_failed = False
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                cleanup_failed = True
                logger.exception('本票工作簿关闭失败')
        book = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                cleanup_failed = True
                logger.exception('本票Excel实例退出失败')
        excel = None
        pythoncom.CoUninitialize()
        if cleanup_failed:
            raise ExcelWriteError(f'本票Excel清理失败，保存状态需人工核对：{path}，勿重复生成。', retry_safe=False)
    return ShipmentOutput(path, rows)
