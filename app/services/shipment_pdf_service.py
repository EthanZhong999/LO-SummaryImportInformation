"""在本票子表的临时副本中准备并导出四张单据。"""
from pathlib import Path
from shutil import copyfile, copyfileobj
from tempfile import TemporaryDirectory

from app.services.shipment_output_service import _new_excel, validate_master_path


DOCUMENTS = ('Contract', 'CI', 'PL', '申报要素')


def export_shipment_package(source, waybill, invoice):
    """先暂存本票原件，导出成功后归档；部分成功不可当作整批未执行。"""
    from app.services.excel_service import ExcelWriteError

    source = Path(source).resolve(strict=True)
    originals = [('运单', Path(waybill)), ('发票装箱单', Path(invoice))]
    with TemporaryDirectory(prefix='.originals-', dir=source.parent) as directory:
        pending = []
        reserved_names = {source.with_suffix('.pdf').name.casefold()}
        for label, original in originals:
            if original.suffix.lower() != '.pdf' or not original.is_file():
                raise ValueError(f'{label} PDF 不存在或类型无效，请重新选择：{original}')
            destination = source.parent / original.name
            # Windows 文件名不区分大小写，归档名也不能与生成的 PDF 冲突。
            if destination.name.casefold() in reserved_names:
                raise ValueError(f'PDF 文件名冲突，无法按原文件名归档：{destination.name}。请先调整上传文件名。')
            reserved_names.add(destination.name.casefold())
            if destination.exists():
                raise ValueError(f'原件归档文件已存在，不会覆盖：{destination}')
            staged = Path(directory) / destination.name
            copyfile(original, staged)
            pending.append((staged, destination))
        pdf = export_shipment_pdf(source)
        saved = []
        try:
            for staged, destination in pending:
                # 独占创建，防止预检后出现的同名文件被覆盖。
                with staged.open('rb') as reader, destination.open('xb') as writer:
                    copyfileobj(reader, writer)
                saved.append(destination.name)
        except OSError as exc:
            raise ExcelWriteError(
                f'四页 PDF 已生成：{pdf}；原件归档未完成。'
                f'已归档：{"、".join(saved) or "无"}；'
                f'请人工检查 {destination}（可能不完整），从原上传位置补存原件，勿重复生成 PDF。',
                retry_safe=False,
            ) from exc
    return pdf


def _lookup_rows(sheet, key_column, value_columns, required, *, letters_only=False):
    def normalize(key):
        key = key.strip()
        if letters_only:
            return ''.join(char for char in key if char.isalpha()).casefold()
        return key

    required_by_key = {normalize(key): key for key in required}
    if '' in required_by_key:
        raise ValueError(f'主数据 {sheet.Name} 的匹配键为空。')
    last = sheet.Cells(sheet.Rows.Count, key_column).End(-4162).Row  # xlUp
    result = {}
    for row in range(2, last + 1):
        key = sheet.Cells(row, key_column).Value2
        if not isinstance(key, str):
            continue
        original_key = required_by_key.get(normalize(key))
        if original_key is None or original_key in result:
            continue
        values = []
        for column in value_columns:
            cell = sheet.Cells(row, column)
            value = cell.Value2
            if cell.HasFormula or not isinstance(value, str) or not value.strip():
                raise ValueError(f'主数据 {sheet.Name} 第{row}行第{column}列需要有效文本。')
            values.append(value)
        result[original_key] = values
        if result.keys() == required:
            break
    missing = required - result.keys()
    if missing:
        raise ValueError(f'主数据 {sheet.Name} 缺少：{", ".join(sorted(missing))}。')
    return result


def _extend_document(sheet, count, first, capacity, detail_columns):
    last = first + count - 1
    boundary = first + capacity
    extra = max(count - capacity, 0)
    sample = sheet.Range(f'A{first}:{detail_columns}{first}')
    formulas = sample.FormulaR1C1
    if extra:
        sheet.Rows(f'{boundary}:{boundary + extra - 1}').Insert()
        for row in range(boundary, boundary + extra):
            sample.Copy(Destination=sheet.Range(f'A{row}:{detail_columns}{row}'))
            sheet.Rows(row).RowHeight = sample.RowHeight
    # 按相对引用填充，保留其他列中的整票字段及下移后的签名区。
    sheet.Range(f'A{first}:{detail_columns}{last}').FormulaR1C1 = formulas * count
    if count < capacity:
        sheet.Range(f'A{last + 1}:{detail_columns}{boundary - 1}').ClearContents()
    return last, extra


def _prepare_documents(book, master, materials, hawb):
    if tuple(sheet.Name for sheet in book.Worksheets) != ('shipment record', *DOCUMENTS):
        raise ValueError('子表的工作表名称或顺序与 Phase 13 模板不符。')
    source = book.Worksheets(1)
    count = len(materials)
    for row, material in enumerate(materials, 2):
        if source.Cells(row, 3).Value2 != material or source.Cells(row, 29).Value2 != hawb:
            raise ValueError(f'子表第{row}行与本次确认批次不符。')
    shipper = source.Range('F2').Value2
    if not isinstance(shipper, str) or not shipper.strip():
        raise ValueError('子表缺少 Shipper。')
    shipper = shipper.strip()
    address, country, relationship = _lookup_rows(
        master.Worksheets('basic data'), 2, (3, 4, 5), {shipper}, letters_only=True,
    )[shipper]
    declaration = _lookup_rows(master.Worksheets('HS'), 3, (12, 13, 15), set(materials))
    contract, ci, pl, details = (book.Worksheets(name) for name in DOCUMENTS)
    for sheet, label, address_check in (
        (contract, 'Total', 'A33'), (ci, 'Total', 'A50'), (pl, 'Total', 'A46'),
        (details, '料号/型号', 'C5'),
    ):
        if sheet.Range(address_check).Value2 != label:
            raise ValueError(f'{sheet.Name} 的模板位置 {address_check} 与预期不符。')
    if not details.Range('C6').HasFormula or details.Range('C7').HasFormula:
        raise ValueError('申报要素须使用只有第6行为明细样板的新版模板。')

    # 保留单元格格式及换行，文本前缀防止主数据被解释成公式。
    contract.Range('A7').Value2 = "'" + address
    details.Range('B3').Value2 = "'" + country
    details.Range('E6,H6').ClearContents()  # 两列改写为本票主数据，扩行前先移除旧外链公式。
    print_ranges = []
    for sheet, first, capacity, columns, total, sums, area_col, bottom in (
        (contract, 10, 20, 'I', 33, 'DI', 'K', 53),
        (ci, 12, 38, 'I', 50, 'FI', 'I', 56),
        (pl, 12, 34, 'G', 46, 'FG', 'I', 52),
    ):
        last, extra = _extend_document(sheet, count, first, capacity, columns)
        for column in sums:
            sheet.Range(f'{column}{total + extra}').Formula = f'=SUM({column}{first}:{column}{last})'
        print_ranges.append(f'A1:{area_col}{bottom + extra}')
    _, extra = _extend_document(details, count, 6, 1, 'H')
    # 扩行后按标签定位，特殊关系的值跟随确认区下移。
    relationship_rows = [
        row for row, (label,) in enumerate(details.Range(f'B{6 + count}:B{16 + extra}').Value2, 6 + count)
        if isinstance(label, str) and label.strip().rstrip('：:').strip() == '特殊关系'
    ]
    if len(relationship_rows) != 1:
        raise ValueError('申报要素中无法唯一定位“特殊关系”标签。')
    details.Cells(relationship_rows[0], 3).Value2 = "'" + relationship
    for row, material in enumerate(materials, 6):
        for column, value in zip('AEFH', ('一般贸易', *declaration[material])):
            details.Range(f'{column}{row}').Value2 = "'" + value
    print_ranges.append(f'A1:I{16 + extra}')
    return print_ranges


def _check_print_area(sheet, area):
    if not area or sheet.Evaluate(f'SUMPRODUCT(--ISERROR({area}))') != 0:
        raise ValueError(f'{sheet.Name} 打印范围缺失或存在计算错误，请检查模板和本票数据。')


def prepare_shipment_documents(excel, book, master_path, materials, hawb, volume):
    master = sheet = None
    try:
        master = excel.Workbooks.Open(str(validate_master_path(master_path)), UpdateLinks=0,
                                     ReadOnly=True, AddToMru=False)
        print_ranges = _prepare_documents(book, master, materials, hawb)
        # PL 原合计引用 I12。整票体积只写一次，六位小数保留立方厘米换算精度。
        volume_cell = book.Worksheets('PL').Range('I12')
        volume_cell.NumberFormat = '0.000000'
        volume_cell.Value2 = float(volume)
        volume_cell = None
        book.Activate()
        excel.CalculateFull()
        for name, area in zip(DOCUMENTS, print_ranges):
            sheet = book.Worksheets(name)
            _check_print_area(sheet, area)
            sheet.PageSetup.PrintArea = area
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            sheet.Visible = -1  # xlSheetVisible
    finally:
        sheet = None
        if master is not None:
            master.Close(False)


def _export_copy(excel, copy_path, staged):
    book = sheet = None
    try:
        book = excel.Workbooks.Open(str(copy_path), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        excel.CalculateFull()
        for name in DOCUMENTS:
            sheet = book.Worksheets(name)
            _check_print_area(sheet, sheet.PageSetup.PrintArea)
        # 主表仅作为公式来源。隐藏临时副本中的主表后按工作簿顺序导出，
        # 避免依赖后台 Excel 窗口的工作表分组选取状态。
        book.Worksheets('shipment record').Visible = 0  # xlSheetHidden
        book.ExportAsFixedFormat(Type=0, Filename=str(staged),
                                 IgnorePrintAreas=False, OpenAfterPublish=False)
    finally:
        sheet = None
        if book is not None:
            book.Close(False)


def export_shipment_pdf(source):
    import pypdfium2
    import pythoncom

    source = Path(source).resolve(strict=True)
    destination = source.with_suffix('.pdf')
    if destination.exists():
        raise ValueError(f'本票 PDF 已存在：{destination}。')
    with TemporaryDirectory(prefix='.pdf-', dir=source.parent) as directory:
        copy_path = Path(directory) / source.name
        staged = copy_path.with_suffix('.pdf')
        copyfile(source, copy_path)
        excel = None
        pythoncom.CoInitialize()
        try:
            excel = _new_excel()
            _export_copy(excel, copy_path, staged)
        finally:
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                excel = None
                pythoncom.CoUninitialize()
        with pypdfium2.PdfDocument(staged) as pdf:
            if len(pdf) != len(DOCUMENTS):
                raise ValueError(f'导出得到 {len(pdf)} 页，预期为四页。')
        # Windows 同卷重命名，若其他进程已创建目标则失败，不覆盖。
        staged.rename(destination)
    return destination
