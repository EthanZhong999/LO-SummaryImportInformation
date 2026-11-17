"""正式桌面界面；只有主线程接触Tk，OCR进程仅发送管道消息。"""
from dataclasses import replace
import logging
from pathlib import Path
from time import perf_counter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD
from app.ui.pdf_drop_zone import PdfDropZone

from app.services.settings_service import Settings, save_settings
from app.services.runtime_paths import resource_root
from app.suppliers.registry import list_suppliers
from app.ui.gui_logic import PreviewState, record_row, validate_pdf_path

from app.ui.preview_edit import FIELD_NAMES, LINE_FIELDS, SHIPMENT_FIELDS, MASTER_FIELD_NAMES, edit_master_row, evidence_summary
from app.ui.recognition_task import RecognitionTask
from app.ui.excel_task import ExcelTask
from app.services.excel_service import validate_excel1_path, validate_excel2_path
from app.services.shipment_output_service import (
    find_supplier_image, plan_shipment_path, shipment_volume, validate_master_path, validate_output_root,
)
from app.ui.gui_logic import validate_excel_submission, EXCEL_CONFIDENCE_THRESHOLD, confidence_review_text
from app.core.validators import low_confidence_evidence
from app.version import APP_VERSION

logger = logging.getLogger(__name__)

# 公司标准色RGB(0,153,154)，悬停色仅用于操作反馈。
BRAND_COLOR = '#00999A'
BRAND_HOVER = '#007D7E'
BRAND_TEXT = '#102D2D'
UI_FONT = ('Microsoft YaHei UI', 15)
UI_HEADING = ('Microsoft YaHei UI', 17, 'bold')
UI_TABLE_FONT = ('Microsoft YaHei UI', 11)


class MainWindow(ctk.CTk):
    def __init__(self, root: Path, settings: Settings):
        super().__init__()
        TkinterDnD.require(self)
        self.root_path = root
        # 从项目根目录定位图标，避免启动目录改变后找不到资源；同时替换默认子窗口图标。
        self.iconbitmap(default=str(resource_root() / 'assets/icons/invoice_processor.ico'))
        self.preview_state = PreviewState()
        self.task = RecognitionTask()
        self.excel_task = ExcelTask()
        self.excel_operation = 'excel1'
        self.master_preview = None
        self.close_pending = False
        self.auto_running = False
        self.auto_waiting_for_correction = False
        self.auto_pdf_path = None
        self.last_poll_time = perf_counter()
        self.suppliers = list_suppliers()
        self.supplier_by_name = {s.display_name: s.key for s in self.suppliers}
        selected = next((s for s in self.suppliers if s.key == settings.last_supplier), self.suppliers[0])
        self.supplier = tk.StringVar(value=selected.display_name)
        self.waybill = tk.StringVar()
        self.invoice = tk.StringVar()
        self.excel1 = tk.StringVar(value=settings.excel1_path)
        self.excel2 = tk.StringVar(value=settings.excel2_path)
        self.master_data = tk.StringVar(value=settings.master_data_path)
        self.output_root = tk.StringVar(value=settings.output_root)
        self.supplier_image_dir = tk.StringVar(value=settings.supplier_image_dir)
        self.locked_inputs = []
        self.title(f'Invoice Document Processor v{APP_VERSION}')
        self.geometry('1240x800')
        self.minsize(900, 680)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        tabs = ctk.CTkTabview(self, corner_radius=12, segmented_button_font=UI_FONT,
                              segmented_button_selected_color=BRAND_COLOR,
                              segmented_button_selected_hover_color=BRAND_HOVER, text_color=BRAND_TEXT)
        tabs.grid(row=0, column=0, sticky='nsew', padx=20, pady=12)
        for name in ('单据处理', 'Excel设置'):
            tabs.add(name)
        self._build_processing(tabs.tab('单据处理'))
        self._build_settings(tabs.tab('Excel设置'))
        for variable in (self.supplier, self.waybill, self.invoice):
            variable.trace_add('write', self._input_changed)
        self.excel1.trace_add('write', lambda *_: self._refresh_buttons())
        self.excel2.trace_add('write', lambda *_: self._refresh_buttons())
        self.master_data.trace_add('write', self._master_changed)
        self.output_root.trace_add('write', lambda *_: self._refresh_buttons())
        self.protocol('WM_DELETE_WINDOW', self._close)
        self.after(100, self._poll)
        self._refresh_buttons()

    def _build_processing(self, frame):
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(4, weight=3)
        frame.grid_rowconfigure(6, weight=2)
        inputs = ctk.CTkFrame(frame)
        inputs.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        inputs.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(inputs, font=UI_HEADING, text='供应商').grid(row=0, column=0, padx=12, pady=6)
        menu = ctk.CTkComboBox(inputs, width=220, font=UI_FONT, dropdown_font=UI_FONT, height=34, fg_color=BRAND_COLOR, button_color=BRAND_HOVER, button_hover_color=BRAND_HOVER, text_color=BRAND_TEXT, dropdown_hover_color=BRAND_COLOR, variable=self.supplier, values=list(self.supplier_by_name))
        menu.grid(row=0, column=1, sticky='w', pady=6)
        self.locked_inputs.append(menu)
        self.supplier_menu = menu
        menu.bind('<KeyRelease>', self._filter_suppliers)
        menu.bind('<Return>', self._complete_supplier)
        menu.bind('<FocusOut>', self._complete_supplier)
        self.file_labels = []
        self.pdf_drop_zones = []
        pdf_inputs = ctk.CTkFrame(inputs, fg_color='transparent')
        pdf_inputs.grid(row=1, column=0, columnspan=4, sticky='ew', padx=12, pady=(4, 8))
        for column, (label, variable) in enumerate((('运单PDF', self.waybill), ('发票 / 装箱单PDF', self.invoice))):
            pdf_inputs.grid_columnconfigure(column, weight=1, uniform='pdf')
            ctk.CTkLabel(pdf_inputs, text=label, font=UI_HEADING).grid(row=0, column=column, sticky='w', padx=6)
            zone = PdfDropZone(pdf_inputs, variable=variable, title=label,
                               is_busy=lambda: self.preview_state.busy)
            zone.grid(row=1, column=column, sticky='ew', padx=6)
            info = ctk.CTkLabel(pdf_inputs, text='未选择', height=24, font=UI_FONT)
            info.grid(row=2, column=column, sticky='w', padx=6)
            self.file_labels.append(info)
            self.pdf_drop_zones.append(zone)
            self.locked_inputs.append(zone)
        actions = ctk.CTkFrame(frame, fg_color='transparent')
        actions.grid(row=1, column=0, sticky='ew')
        self.auto_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120,
                                       fg_color=BRAND_COLOR, hover_color=BRAND_HOVER,
                                       text_color=BRAND_TEXT, text='一键处理', command=self._start_auto)
        self.auto_button.pack(side='top', anchor='w', pady=(6, 0))
        self.recognize_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='识别并预览', command=self._start)
        self.recognize_button.pack(side='left', padx=(0, 4), pady=8)
        self.confirm_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='核对主数据', command=self._confirm)
        self.confirm_button.pack(side='left', padx=4)
        self.write_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='生成本票子表', command=self._write_excel)
        self.write_button.pack(side='left', padx=4)
        self.pdf_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER,
                                      text_color=BRAND_TEXT, text='生成 PDF', command=self._generate_pdf)
        self.pdf_button.pack(side='left', padx=4)
        self.sync_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=170, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER,
                                       text_color=BRAND_TEXT, text='同步到报关信息总表', command=self._sync_excel2)
        self.sync_button.pack(side='left', padx=4)
        self.next_button = ctk.CTkButton(actions, font=UI_FONT, height=36, width=120, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='处理下一组', command=self._next)
        self.next_button.pack(side='left', padx=4)
        self.shipment = ctk.CTkLabel(frame, text='整票信息：等待识别', font=UI_FONT, anchor='w', justify='left', wraplength=850)
        self.shipment.grid(row=2, column=0, sticky='ew', pady=6)
        ctk.CTkLabel(frame, text='明细预览 · 选择一行查看字段来源', font=UI_HEADING, anchor='w').grid(row=3, column=0, sticky='ew')
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=4, column=0, sticky='nsew')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = ('状态', 'Inv. No.', 'PO No.', 'Material No.', 'Origin', 'HS Code', 'Value', 'Currency', 'Qty', 'Net Weight', 'G.W', 'Package')
        style = ttk.Style(self)
        style.configure('Brand.Treeview', font=UI_TABLE_FONT, rowheight=31)
        style.configure('Brand.Treeview.Heading', font=('Microsoft YaHei UI', 11, 'bold'), padding=(6, 6))
        style.configure('Brand.TButton', font=UI_TABLE_FONT, padding=(10, 5))
        style.configure('Brand.TCheckbutton', font=UI_TABLE_FONT, padding=(5, 5))
        style.map('Brand.Treeview', background=[('selected', BRAND_COLOR)], foreground=[('selected', BRAND_TEXT)])
        self.tree = ttk.Treeview(table_frame, columns=columns, style='Brand.Treeview', show='headings', height=5, selectmode='browse')
        for name, width in zip(columns, (74, 137, 137, 168, 76, 105, 95, 88, 90, 110, 90, 90)):
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width, minwidth=width, stretch=False)
        self.tree.grid(row=0, column=0, sticky='nsew')
        for orient, row, col, sticky, command in (('vertical', 0, 1, 'ns', self.tree.yview), ('horizontal', 1, 0, 'ew', self.tree.xview)):
            bar = ttk.Scrollbar(table_frame, orient=orient, command=command)
            bar.grid(row=row, column=col, sticky=sticky)
            self.tree.configure(**{('yscrollcommand' if orient == 'vertical' else 'xscrollcommand'): bar.set})
        self.tree.bind('<<TreeviewSelect>>', self._show_line)
        self.tree.bind('<Double-1>', self._edit_clicked_line)
        controls = ttk.Frame(frame)
        controls.grid(row=5, column=0, sticky='w', pady=5)
        ttk.Button(controls, style='Brand.TButton', text='查看识别来源', command=lambda: self._show_evidence(None)).pack(side='left')
        self.source_index = None
        self.technical_source = tk.BooleanVar(self, False)
        ttk.Checkbutton(controls, style='Brand.TCheckbutton', text='展开技术详情', variable=self.technical_source,
                        command=lambda: self._show_evidence(self.source_index)).pack(side='left', padx=6)
        self.edit_shipment_button = ttk.Button(controls, style='Brand.TButton', text='编辑整票信息', command=lambda: self._edit(None))
        self.edit_shipment_button.pack(side='left', padx=6)
        self.edit_line_button = ttk.Button(controls, style='Brand.TButton', text='编辑明细', command=self._edit_selected_line)
        self.edit_line_button.pack(side='left')
        self.edit_master_button = ttk.Button(controls, style='Brand.TButton', text='编辑本票主数据', command=self._edit_master)
        self.edit_master_button.pack(side='left', padx=6)
        self.evidence = ctk.CTkTextbox(frame, height=110, wrap='word', font=UI_FONT, corner_radius=9)
        self.evidence.grid(row=6, column=0, sticky='nsew')
        self.evidence.configure(state='disabled')
        self.status = ctk.CTkLabel(frame, text='请选择两份PDF，识别后核对主数据、确认并生成本票子表。', font=UI_FONT, anchor='w', wraplength=850)
        self.status.grid(row=7, column=0, sticky='ew', pady=5)
        self.progress = ctk.CTkProgressBar(frame, mode='indeterminate', progress_color=BRAND_COLOR)
        self.progress.grid(row=8, column=0, sticky='ew', pady=4)
        self.progress.set(0)

    def _build_settings(self, frame):
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, font=UI_HEADING, text='模板、主数据、汇总主表及输出目录 · 可保存为默认路径').grid(row=0, column=0, columnspan=3, sticky='w', padx=18, pady=(20, 16))
        for row, (name, variable) in enumerate((('Excel1模板（只读）', self.excel1), ('报关信息总表', self.excel2),
                                                ('物料主数据', self.master_data), ('输出根目录', self.output_root),
                                                ('供应商图片文件夹', self.supplier_image_dir)), 1):
            ctk.CTkLabel(frame, font=UI_FONT, text=name).grid(row=row, column=0, sticky='w', padx=18, pady=10)
            entry = ctk.CTkEntry(frame, textvariable=variable, font=UI_FONT, height=36)
            entry.grid(row=row, column=1, sticky='ew', padx=8)
            button = ctk.CTkButton(frame, font=UI_FONT, height=36, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='浏览', width=70,
                                   command=lambda v=variable: self._browse_directory(v)
                                   if v is self.output_root or v is self.supplier_image_dir else self._browse(v, False))
            button.grid(row=row, column=2, padx=8)
            self.locked_inputs.extend((entry, button))
        button = ctk.CTkButton(frame, font=UI_FONT, height=36, fg_color=BRAND_COLOR, hover_color=BRAND_HOVER, text_color=BRAND_TEXT, text='保存为默认路径', command=self._save_settings)
        button.grid(row=6, column=1, sticky='w', padx=8, pady=18)
        self.locked_inputs.append(button)

    def _browse(self, variable, pdf):
        path = filedialog.askopenfilename(parent=self, filetypes=[('PDF文件', '*.pdf')] if pdf else [('Excel文件', '*.xls *.xlsx')])
        if path:
            variable.set(path)

    def _browse_directory(self, variable):
        path = filedialog.askdirectory(parent=self, mustexist=True)
        if path:
            variable.set(path)

    def _master_changed(self, *_):
        self.master_preview = None
        if not self.preview_state.current_excel_written:
            self.preview_state.current_confirmed = False
        self._refresh_buttons()

    def _save_settings(self, notify=True):
        try:
            save_settings(self.root_path / 'config/settings.json', Settings(
                excel1_path=self.excel1.get().strip(), excel2_path=self.excel2.get().strip(),
                last_supplier=self.supplier_by_name[self.supplier.get()],
                master_data_path=self.master_data.get().strip(), output_root=self.output_root.get().strip(),
                supplier_image_dir=self.supplier_image_dir.get().strip()))
        except (OSError, ValueError, KeyError):
            logger.exception('默认设置保存失败')
            messagebox.showerror('设置保存失败', '无法保存默认设置，请检查配置文件和目录权限。', parent=self)
            return False
        if notify:
            messagebox.showinfo('设置', '默认路径和供应商已保存。', parent=self)
        return True

    def _filter_suppliers(self, event=None):
        query = self.supplier.get().strip().casefold()
        matches = [name for name in self.supplier_by_name if query in name.casefold()]
        self.supplier_menu.configure(values=matches)

    def _complete_supplier(self, event=None):
        if self.preview_state.busy:
            return
        query = self.supplier.get().strip().casefold()
        if not query:
            return
        matches = [name for name in self.supplier_by_name if query in name.casefold()]
        exact = next((name for name in matches if name.casefold() == query), None)
        # 只有唯一匹配或完整名称才确认，避免多个候选时误选供应商。
        selected = exact or (matches[0] if len(matches) == 1 else None)
        if selected is not None and selected != self.supplier.get():
            self.supplier.set(selected)
        self._filter_suppliers()

    def _input_changed(self, *_):
        self.master_preview = None
        self.preview_state.invalidate()
        self._clear_preview()
        self.status.configure(text='输入已修改，请重新识别。', text_color='#946200')
        self._refresh_buttons()

    def _clear_preview(self):
        self.auto_waiting_for_correction = False
        self.tree.delete(*self.tree.get_children())
        self.shipment.configure(text='整票信息：等待识别')
        self.source_index = None
        self.technical_source.set(False)
        self._set_evidence('')

    def _check_shipment_destination(self):
        template = validate_excel1_path(Path(self.excel1.get().strip()))
        output = plan_shipment_path(self.output_root.get().strip(), self.preview_state.current_records)
        shipment_volume(self.preview_state.current_records)
        if output.parent.exists():
            raise ValueError(f'本票目录已存在，请选择其他输出根目录：{output.parent}')
        return template

    def _refresh_buttons(self):
        busy = self.preview_state.busy
        self._refresh_edit_buttons()
        valid = self.supplier.get() in self.supplier_by_name
        for variable, label in zip((self.waybill, self.invoice), self.file_labels):
            try:
                path = validate_pdf_path(variable.get())
                label.configure(text=f'已选择：{path.name}')
            except (ValueError, OSError):
                valid = False
                label.configure(text='未选择' if not variable.get().strip() else '路径错误')
        self.recognize_button.configure(state='normal' if valid and not busy else 'disabled')
        self.auto_button.configure(state='normal' if valid and not busy
                                   and not self.preview_state.current_records else 'disabled')
        diagnostics = self.preview_state.current_diagnostics
        can_confirm = (self.preview_state.current_records and not busy
                       and not self.preview_state.current_confirmed
                       and not diagnostics.issues and not diagnostics.blockers)
        self.confirm_button.configure(state='normal' if can_confirm else 'disabled')
        self.confirm_button.configure(text='确认全部结果' if self.master_preview is not None else '核对主数据')
        self.next_button.configure(state='disabled' if busy else 'normal')
        can_write = False
        try:
            if not self.output_root.get().strip() or self.master_preview is None:
                raise ValueError('请先设置输出目录并核对主数据。')
            self._check_shipment_destination()
            validate_excel_submission(self.preview_state.current_records, self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
            can_write = self.preview_state.can_write(True)
        except (ValueError, OSError):
            pass
        self.write_button.configure(state='normal' if can_write and not busy else 'disabled')
        self.pdf_button.configure(state='normal' if self.preview_state.excel1_written_path and not busy else 'disabled')
        can_sync = False
        try:
            validate_excel2_path(Path(self.excel2.get().strip()))
            can_sync = self.preview_state.can_sync(True)
        except (ValueError, OSError):
            pass
        self.sync_button.configure(state='normal' if can_sync and not busy else 'disabled')
        for widget in self.locked_inputs:
            widget.configure(state='disabled' if busy else 'normal')

    def _start_auto(self):
        if self.preview_state.busy or self.preview_state.current_records:
            return
        try:
            # 在耗时识别前检查所有静态路径；本票命名和数据门禁仍由各步骤检查。
            if not self.output_root.get().strip():
                raise ValueError('请先设置输出根目录。')
            validate_excel1_path(Path(self.excel1.get().strip()))
            validate_excel2_path(Path(self.excel2.get().strip()))
            validate_master_path(self.master_data.get().strip())
            validate_output_root(self.output_root.get().strip())
            find_supplier_image(self.supplier_image_dir.get().strip(),
                                self.supplier_by_name[self.supplier.get()])
        except (ValueError, OSError, KeyError) as exc:
            self.status.configure(text=f'无法开始一键处理：{exc}', text_color='#b3261e')
            return
        self.auto_running = True
        self.auto_pdf_path = None
        self._start()
        if not self.preview_state.busy:
            self.auto_running = False

    def _advance_auto(self, operation, kind, payload=None):
        """主线程只调度现有任务；每次成功交付意味着后台资源已回收。"""
        if not self.auto_running:
            return
        if kind != 'success' or self.close_pending:
            self.auto_running = False
            if kind != 'success':
                text = self.status.cget('text')
                saved = self.preview_state.excel1_written_path
                if saved:
                    text += f' 已保存子表：{saved}。'
                if self.auto_pdf_path:
                    text += f' 已生成 PDF：{self.auto_pdf_path}。'
                self.status.configure(text='一键处理已停止：' + text, text_color='#b3261e')
            return
        diagnostics = self.preview_state.current_diagnostics
        if operation == 'recognition' and diagnostics.issues and not diagnostics.blockers:
            self.auto_running = False
            self.auto_waiting_for_correction = True
            self.status.configure(text='一键处理等待修正。请查看下方问题，修正后核对主数据并确认继续。',
                                  text_color='#946200')
            self._refresh_buttons()
            return
        try:
            if operation in ('recognition', 'master'):
                validate_excel_submission(self.preview_state.current_records,
                                          self.preview_state.current_diagnostics,
                                          EXCEL_CONFIDENCE_THRESHOLD)
                warnings = list(self.preview_state.current_diagnostics.warnings)
                if self._review_warning():
                    warnings.append('存在低分字段，请核对识别来源。')
                if operation == 'master':
                    # NEXPERIA 已确认主数据重复取首条，仅展示提醒；其他供应商不沿用此例外。
                    supplier = self.supplier_by_name[self.supplier.get()]
                    duplicates = (self.master_preview.duplicate_material_warnings
                                  if supplier == 'NEXPERIA' else [])
                    warnings.extend(w for w in self.master_preview.warnings if w not in duplicates)
                if warnings:
                    raise ValueError('；'.join(warnings))
            if operation == 'recognition':
                self._confirm()
            elif operation == 'master':
                # 点击一键处理即授权无提醒批次自动确认，仍走既有写入门禁。
                if not self.preview_state.confirm():
                    raise ValueError('当前批次无法确认。')
                self._write_excel()
            elif operation == 'excel1':
                self._generate_pdf()
            elif operation == 'pdf':
                self.auto_pdf_path = payload
                self._sync_excel2()
            elif operation == 'excel2':
                self.auto_running = False
                self.status.configure(text=f'一键处理完成：子表 {self.preview_state.excel1_written_path}；'
                                      f'PDF {self.auto_pdf_path}；总表新增 {len(payload)} 条。',
                                      text_color='#247044')
            # 启动函数已显示具体失败原因；失败时不继续调度或自动重试。
            if not self.preview_state.busy:
                self.auto_running = False
        except (ValueError, OSError) as exc:
            self.auto_running = False
            logger.info('一键处理暂停 operation=%s reason=%s', operation, exc)
            self.status.configure(text=f'一键处理已暂停：{exc} 请复核后使用分步按钮继续。',
                                  text_color='#946200')
        self._refresh_buttons()

    def _start(self):
        if self.preview_state.busy:
            return
        try:
            supplier = self.supplier_by_name[self.supplier.get()]
            waybill, invoice = validate_pdf_path(self.waybill.get()), validate_pdf_path(self.invoice.get())
        except (ValueError, OSError, KeyError):
            messagebox.showerror('输入错误', '请确认供应商及两个PDF路径有效。', parent=self)
            self._refresh_buttons()
            return
        if not self._save_settings(False) or not self.preview_state.begin():
            return
        self._clear_preview()
        self.master_preview = None
        self._refresh_buttons()
        for label in self.file_labels:
            label.configure(text='识别中')
        self.status.configure(text='正在准备本地OCR并识别，请稍候…', text_color=BRAND_HOVER)
        self.progress.start()
        logger.info('识别开始 supplier=%s waybill=%s invoice=%s', supplier, waybill, invoice)
        try:
            self.task.start(self.root_path, supplier, waybill, invoice)
        except Exception:
            logger.exception('OCR进程启动失败')
            self.preview_state.fail()
            self.progress.stop()
            self.progress.set(0)
            self.status.configure(text='无法启动OCR后台进程，请检查日志和运行环境。', text_color='#b3261e')
            self._refresh_buttons()

    def _poll(self):
        now = perf_counter()
        if self.preview_state.running and now - self.last_poll_time > 0.5:
            logger.warning('GUI主线程轮询延迟 seconds=%.2f', now - self.last_poll_time)
        self.last_poll_time = now
        for kind, payload in self.task.poll():
            if kind == 'status':
                self.status.configure(text=payload)
                continue
            self.progress.stop()
            self.progress.set(0)
            if kind == 'success':
                self.preview_state.succeed(*payload)
                low = low_confidence_evidence(self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
                self._refresh_record_display(low)
                diagnostics = self.preview_state.current_diagnostics
                if diagnostics.blockers:
                    text = '本票资料无法完整识别。已显示可定位的结果，请检查下方问题和原 PDF。'
                elif diagnostics.issues:
                    text = f'识别完成，有 {len(diagnostics.issues)} 项问题待修正。'
                else:
                    text = '识别完成，等待确认。'
                if low:
                    text += f' 有{len(low)}项低分，详见下方复核清单。'
                try:
                    validate_excel_submission(self.preview_state.current_records, self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
                except ValueError as exc:
                    text += f' 写入被阻止：{exc}'
                self.status.configure(text=text, text_color='#946200')
                self._show_evidence(None)
            else:
                self.preview_state.fail()
                self._clear_preview()
                self.status.configure(text=payload, text_color='#b3261e')
            self._refresh_buttons()
            for label in self.file_labels:
                label.configure(text='识别成功' if kind == 'success' else '错误')
            self._advance_auto('recognition', kind)
        for kind, payload in self.excel_task.poll():
            self.progress.stop()
            self.progress.set(0)
            if self.excel_operation == 'master':
                self.preview_state.excel_running = False
                if kind == 'success':
                    self.master_preview = payload
                    self.status.configure(text='主数据已读取，请核对下方补齐值、来源及差异，再点击“确认全部结果”。', text_color='#946200')
                    self._show_evidence(None)
                else:
                    self.master_preview = None
                    self.status.configure(text=payload[0], text_color='#b3261e')
                self._refresh_buttons()
                self._advance_auto('master', kind)
                continue
            if self.excel_operation == 'pdf':
                self.preview_state.excel_running = False
                if kind == 'success':
                    self.status.configure(text=f'本票 PDF 已生成（4页），两份上传原件已归档到同一文件夹：{payload}', text_color='#247044')
                else:
                    self.status.configure(text=payload[0], text_color='#b3261e')
                self._refresh_buttons()
                self._advance_auto('pdf', kind, payload)
                continue
            if kind == 'success':
                if self.excel_operation == 'excel2':
                    self.preview_state.finish_sync(payload)
                else:
                    self.preview_state.finish_excel(payload.inserted_rows, payload.path)
                warning = self._review_warning()
                if self.excel_operation == 'excel2':
                    text = (f'本票子表及报关信息总表均已保存：{self.preview_state.excel1_written_path}；'
                            f'总表新增{len(payload)}条，第{payload[0]}至{payload[-1]}行。')
                else:
                    text = f'本票子表已保存：{payload.path}；{len(payload.inserted_rows)}条，行号={payload.inserted_rows}。请同步到报关信息总表。'
                if warning:
                    text += ' 有低分字段，请按下方清单自行复核。'
                self.status.configure(text=text, text_color='#946200' if warning else '#247044')
                self._show_evidence(None)
            else:
                text, retry_safe = payload
                if self.excel_operation == 'excel2':
                    self.preview_state.fail_sync(retry_safe)
                    text = '本票子表保持已保存状态，本次未重新生成。' + text
                else:
                    self.preview_state.fail_excel(retry_safe)
                self.status.configure(text=text, text_color='#b3261e')
            self._refresh_buttons()
            self._advance_auto(self.excel_operation, kind, payload)
        if self.close_pending and not self.preview_state.busy:
            self.destroy()
            return
        self.after(100, self._poll)

    def _refresh_record_display(self, low):
        diagnostics = self.preview_state.current_diagnostics
        for index, record in enumerate(self.preview_state.current_records):
            needs_review = (any(e.line_index is None or e.line_index == index for e in low)
                            or any(issue.line_index in (None, index) for issue in diagnostics.issues)
                            or bool(diagnostics.blockers))
            values = record_row(record, needs_review=needs_review, first_row=index == 0)
            if self.tree.exists(str(index)):
                self.tree.item(str(index), values=values)
            else:
                self.tree.insert('', 'end', iid=str(index), values=values)
        if not self.preview_state.current_records:
            r = diagnostics.shipment_preview
            if r is not None:
                self.shipment.configure(text=f'ETD：{r.etd}    承运商：{r.carrier}    毛重(kg)：{r.gross_weight}    '
                                        f'件数：{r.package}    发运地：{r.port_of_dispatch}    运单号：{r.hawb_bl}')
            return
        r = self.preview_state.current_records[0]
        volume = r.volume_m3 if r.volume_m3 is not None else '未识别'
        self.shipment.configure(text=f'供应商：{self.supplier.get()}    Shipper：{r.shipper}    Currency：{r.currency}    ETD：{r.etd}    承运商：{r.carrier}\n毛重(kg)：{r.gross_weight}    体积(m³)：{volume}    件数：{r.package}    发运地：{r.port_of_dispatch}    HAWB / BL：{r.hawb_bl}    明细：{len(self.preview_state.current_records)}条')

    def _refresh_edit_buttons(self):
        editable = self.preview_state.can_edit
        self.edit_shipment_button.configure(state='normal' if editable else 'disabled')
        line_selected = editable and self.tree.selection()
        self.edit_line_button.configure(state='normal' if line_selected else 'disabled')
        self.edit_master_button.configure(state='normal' if line_selected and self.master_preview is not None else 'disabled')

    def _edit_clicked_line(self, event):
        row = self.tree.identify_row(event.y)
        if row and self.tree.identify_region(event.x, event.y) == 'cell':
            self._edit(int(row))

    def _edit_selected_line(self):
        selected = self.tree.selection()
        if selected:
            self._edit(int(selected[0]))

    def _edit(self, index, focus_field=None):
        state = self.preview_state
        if not state.can_edit:
            return
        record = state.current_records[0 if index is None else index]
        fields = SHIPMENT_FIELDS if index is None else LINE_FIELDS
        if index is None:
            pending = {issue.field_name for issue in state.current_diagnostics.issues}
            fields += tuple(name for name in ('shipper', 'carrier', 'volume_m3') if name in pending)
        title = '编辑整票信息' if index is None else f'编辑第 {index + 1} 条明细'

        def apply(values):
            if state.apply_edit(index, values, review_field=focus_field):
                self.master_preview = None
                self._refresh_record_display(low_confidence_evidence(
                    state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD))
                self._show_evidence(index)
                remaining = len(state.current_diagnostics.issues)
                text = (f'已保存人工修正，还有 {remaining} 项问题待修正。' if remaining
                        else '已保存人工修正，请重新核对主数据并确认。')
                self.status.configure(text=text, text_color='#946200')
                self._refresh_buttons()

        subtitle = '修改将应用到本票所有明细。' if index is None else f'当前物料  {record.material_no}'
        self._edit_dialog(title, subtitle, {name: getattr(record, name) for name in fields},
                          FIELD_NAMES, apply, focus_field=focus_field)

    def _edit_master(self):
        selected = self.tree.selection()
        snapshot = self.master_preview
        if not selected or snapshot is None or not self.preview_state.can_edit:
            return
        index = int(selected[0])

        def apply(values):
            if not self.preview_state.can_edit or self.master_preview is not snapshot:
                raise ValueError('主数据快照已变化，请重新打开表单。')
            row = edit_master_row(snapshot.rows[index], values)
            if row != snapshot.rows[index]:
                rows = list(snapshot.rows)
                rows[index] = row
                self.master_preview = replace(snapshot, rows=rows,
                                              original_rows=snapshot.original_rows or snapshot.rows)
                self.preview_state.current_confirmed = False
                self._show_evidence(index)
                self.status.configure(text='本票主数据已修正，请重新确认全部结果。', text_color='#946200')
                self._refresh_buttons()

        self._edit_dialog(f'编辑第 {index + 1} 条本票主数据',
                          f'{snapshot.material_numbers[index]}  ·  仅应用于本票，源文件保持只读',
                          snapshot.rows[index], MASTER_FIELD_NAMES, apply)

    def _edit_dialog(self, title, subtitle, values, labels, apply, *, focus_field=None):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.minsize(720, 0)
        for column in range(2):
            dialog.columnconfigure(column, weight=1, uniform='field')
        ttk.Label(dialog, text=title, font=('Microsoft YaHei UI', 18, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w', padx=22, pady=(18, 4))
        ttk.Label(dialog, text=subtitle, font=UI_TABLE_FONT).grid(
            row=1, column=0, columnspan=2, sticky='w', padx=22, pady=(0, 12))
        entries = {}
        for i, (name, value) in enumerate(values.items()):
            row, column = divmod(i, 2)
            label = labels[name] + ('  YYYY-MM-DD' if name == 'etd' else '')
            ttk.Label(dialog, text=label, font=('Microsoft YaHei UI', 11, 'bold')).grid(row=2 + 2 * row, column=column, sticky='w', padx=22)
            entry = tk.Text(dialog, height=1, width=26, wrap='none', font=('Microsoft YaHei UI', 12),
                            padx=9, pady=7, relief='solid', borderwidth=1, undo=True)
            entry.grid(row=3 + 2 * row, column=column, sticky='ew', padx=22, pady=(4, 11))
            entry.insert('1.0', '' if value is None else str(value))
            entry.bind('<Tab>', lambda event: (event.widget.tk_focusNext().focus_set(), 'break')[1])
            entries[name] = entry
        footer = 2 + 2 * ((len(values) + 1) // 2)
        error = ttk.Label(dialog, foreground='#a12525', font=UI_TABLE_FONT, wraplength=660)
        error.grid(row=footer, column=0, columnspan=2, sticky='w', padx=22, pady=(0, 4))

        def save():
            try:
                apply({name: entry.get('1.0', 'end-1c') for name, entry in entries.items()})
            except ValueError as exc:
                error.configure(text=str(exc))
                return
            dialog.destroy()

        ttk.Button(dialog, style='Brand.TButton', text='取消', command=dialog.destroy).grid(
            row=footer + 1, column=0, sticky='e', padx=22, pady=(4, 18))
        ctk.CTkButton(dialog, font=UI_FONT, height=36, text='保存', command=save, fg_color=BRAND_COLOR,
                      hover_color=BRAND_HOVER, text_color=BRAND_TEXT).grid(
            row=footer + 1, column=1, sticky='w', padx=22, pady=(4, 18))
        dialog.bind('<Escape>', lambda _: dialog.destroy())
        dialog.wait_visibility()
        dialog.grab_set()
        entries.get(focus_field, next(iter(entries.values()))).focus_set()

    def _set_evidence(self, text, issue_links=()):
        self.evidence.configure(state='normal')
        self.evidence.delete('1.0', 'end')
        self.evidence.insert('1.0', text)
        for line_number, issue in issue_links:
            tag = f'issue_{line_number}'
            self.evidence.tag_add(tag, f'{line_number}.0', f'{line_number}.end')
            self.evidence.tag_config(tag, foreground='#b3261e', underline=True)
            self.evidence.tag_bind(tag, '<Button-1>',
                                   lambda _, target=issue: self._open_issue(target))
        self.evidence.configure(state='disabled')

    def _open_issue(self, issue):
        if not self.preview_state.can_edit:
            return
        if issue.line_index is not None:
            row = str(issue.line_index)
            self.tree.selection_set(row)
            self.tree.see(row)
        self._edit(issue.line_index, issue.field_name)

    def _show_evidence(self, index):
        self.source_index = index
        state = self.preview_state
        diagnostics = state.current_diagnostics
        technical = self.technical_source.get()
        threshold = EXCEL_CONFIDENCE_THRESHOLD
        rows = []
        issue_links = []
        if diagnostics.issues or diagnostics.blockers:
            rows.append('识别问题（点击字段可修正）' if state.can_edit else '识别问题')
        rows.extend(diagnostics.blockers)
        for issue in diagnostics.issues:
            scope = '整票' if issue.line_index is None else f'第{issue.line_index + 1}条明细'
            page = f'｜第{issue.page_number}页' if issue.page_number is not None else ''
            candidates = '｜OCR：' + ' / '.join(issue.candidates) if issue.candidates else ''
            rows.append(f'{scope}｜{FIELD_NAMES[issue.field_name]}{page}｜{issue.message}{candidates}'
                        .replace('\r', ' ').replace('\n', ' '))
            if state.can_edit:
                issue_links.append((len(rows), issue))
        if not state.original_records:
            self._set_evidence('\n'.join(rows))
            return
        rows.append('整票来源' if index is None else f'第{index + 1}条明细来源')
        rows.extend(evidence_summary(state, e, threshold, technical)
                    for e in diagnostics.evidence if e.line_index == index)
        rows.extend(f'{FIELD_NAMES[name]}｜已人工填写：{getattr(state.current_records[index or 0], name)}'
                    for scope, name in sorted(diagnostics.manual_fields, key=lambda key: key[1])
                    if scope == index)
        low = low_confidence_evidence(diagnostics, threshold)
        if low or diagnostics.warnings:
            rows.append('\n需核对')
            for e in low:
                scope = '整票' if e.line_index is None else f'第{e.line_index + 1}条明细'
                rows.append(f'{scope}｜{evidence_summary(state, e, threshold, technical)}')
            rows.extend(diagnostics.warnings)
        text = '\n'.join(rows)
        if self.master_preview is not None:
            text += '\n\n主数据补齐（本次确认使用以下快照）：\n'
            for i, (source, values) in enumerate(zip(self.master_preview.sources, self.master_preview.rows)):
                if index is None or index == i:
                    original = self.master_preview.original_rows
                    fields = []
                    for name, value in values.items():
                        label = MASTER_FIELD_NAMES[name]
                        if original and original[i][name] != value:
                            fields.append(f'{label}：{value}（原主数据：{original[i][name]}，已人工修正）')
                        else:
                            fields.append(f'{label}：{value}')
                    detail = '；'.join(fields)
                    text += f'第{i + 1}条 {source}\n{detail}\n'
            if self.master_preview.warnings:
                text += '\n主数据提醒：\n' + '\n'.join(self.master_preview.warnings)
        self._set_evidence(text, issue_links)

    def _review_warning(self):
        if not self.preview_state.current_records:
            return ''
        return confidence_review_text(self.preview_state.original_records,
                                      self.preview_state.current_diagnostics,
                                      EXCEL_CONFIDENCE_THRESHOLD)

    def _show_line(self, _):
        self._refresh_edit_buttons()
        selected = self.tree.selection()
        if selected:
            self._show_evidence(int(selected[0]))

    def _confirm(self):
        if self.preview_state.busy:
            return
        if self.master_preview is None:
            try:
                path = validate_master_path(self.master_data.get().strip())
                validate_excel_submission(self.preview_state.current_records, self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
                self.preview_state.excel_running = True
                self.excel_operation = 'master'
                self.excel_task.start_master(path, self.preview_state.current_records)
                self.progress.start()
                self.status.configure(text='正在只读核对主数据，请稍候…', text_color=BRAND_HOVER)
            except Exception as exc:
                self.preview_state.excel_running = False
                logger.exception('主数据核对启动失败')
                self.status.configure(text=str(exc), text_color='#b3261e')
            self._refresh_buttons()
            return
        if self.preview_state.confirm():
            logger.info('用户确认 supplier=%s records=%s', self.supplier.get(), len(self.preview_state.current_records))
            text = '当前识别结果已确认，尚未写入Excel。'
            if self._review_warning():
                text += ' 低分允许写入，请自行核对下方复核清单。'
            try:
                validate_excel_submission(self.preview_state.current_records, self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
                self._check_shipment_destination()
            except (ValueError, OSError) as exc:
                text += f' 无法生成本票子表：{exc}'
            self.status.configure(text=text, text_color='#946200')
            self._refresh_buttons()
            if self.auto_waiting_for_correction:
                self.auto_waiting_for_correction = False
                self.auto_running = True
                self._write_excel()
                if not self.preview_state.busy:
                    self.auto_running = False

    def _write_excel(self):
        if self.preview_state.busy:
            return
        try:
            path = self._check_shipment_destination()
            root = validate_output_root(self.output_root.get().strip())
            image = find_supplier_image(self.supplier_image_dir.get().strip(),
                                        self.supplier_by_name[self.supplier.get()])
            if self.master_preview is None:
                raise ValueError('请先核对主数据并确认。')
            validate_excel_submission(self.preview_state.current_records, self.preview_state.current_diagnostics, EXCEL_CONFIDENCE_THRESHOLD)
            if not self.preview_state.begin_excel(True):
                return
            self.excel_operation = 'excel1'
            self.excel_task.start_output(path, root, self.preview_state.current_records,
                                         self.master_preview, image, prepare_documents=True)
        except Exception as exc:
            logger.exception('Excel提交启动失败')
            self.preview_state.fail_excel(True)
            self.status.configure(text=str(exc) if isinstance(exc, ValueError) else '无法启动Excel后台任务，请检查日志。', text_color='#b3261e')
            self._refresh_buttons()
            return
        self.status.configure(text='正在写入信息收集子表并读回核对，请勿打开目标文件…', text_color=BRAND_HOVER)
        self.progress.start()
        self._refresh_buttons()

    def _generate_pdf(self):
        if self.preview_state.busy or not self.preview_state.excel1_written_path:
            return
        try:
            self.preview_state.excel_running = True
            self.excel_operation = 'pdf'
            self.excel_task.start_pdf(self.preview_state.excel1_written_path,
                                      Path(self.waybill.get()), Path(self.invoice.get()))
        except Exception:
            self.preview_state.excel_running = False
            logger.exception('PDF后台任务启动失败')
            self.status.configure(text='无法启动 PDF 后台任务，请联系IT。', text_color='#b3261e')
        else:
            self.progress.start()
            self.status.configure(text='正在生成本票四页 PDF…', text_color=BRAND_HOVER)
        self._refresh_buttons()

    def _sync_excel2(self):
        if self.preview_state.busy:
            return
        try:
            path = validate_excel2_path(Path(self.excel2.get().strip()))
            if not self.preview_state.begin_sync(True):
                return
            self.excel_operation = 'excel2'
            # 使用成功写入时的源路径，不能因用户更改设置而读错另一份Excel1。
            self.excel_task.start_sync(self.preview_state.excel1_written_path, path,
                                       self.preview_state.inserted_rows)
        except Exception as exc:
            logger.exception('报关信息总表同步启动失败，请联系IT')
            self.preview_state.fail_sync(True)
            self.status.configure(text=str(exc) if isinstance(exc, ValueError) else '无法启动Excel2后台任务，请检查日志。', text_color='#b3261e')
            self._refresh_buttons()
            return
        self.status.configure(text='本票子表已经成功写入，正在读取本批最终值并同步报关信息总表，请勿打开目标文件…', text_color=BRAND_HOVER)
        self.progress.start()
        self._refresh_buttons()

    def _next(self):
        if self.preview_state.busy:
            return
        self.waybill.set('')
        self.invoice.set('')
        self.preview_state.invalidate()
        self._clear_preview()
        self.status.configure(text='请选择下一组两份PDF源文件。', text_color=BRAND_HOVER)
        self._refresh_buttons()

    def _close(self):
        if self.preview_state.busy:
            if not messagebox.askyesno('正在处理', '当前正在识别或写入Excel。是否隐藏窗口，等待后台结束后退出？', parent=self):
                return
            # 不强杀推理进程；继续排空管道、回收进程后关闭。
            self.close_pending = True
            self.withdraw()
            logger.info('等待OCR/Excel后台结束后退出')
            return
        logger.info('应用关闭')
        self.destroy()
