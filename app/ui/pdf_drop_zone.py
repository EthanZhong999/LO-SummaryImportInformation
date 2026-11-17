"""PDF本地选择区域；Canvas虚线边框，点击与拖拽共用路径校验。"""
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from tkinterdnd2 import COPY, DND_FILES, REFUSE_DROP

from app.ui.gui_logic import validate_pdf_path
from app.ui.tk_attachment_cache import TkAttachmentCache


class PdfDropZone(tk.Canvas):
    def __init__(self, master, *, variable, title, is_busy):
        super().__init__(master, height=112, background='#F7FAFA',
                         highlightthickness=0, cursor='hand2', takefocus=True)
        self.variable = variable
        self.title = title
        self.is_busy = is_busy
        self.bind('<Configure>', self._draw)
        self.bind('<Button-1>', self._choose)
        self.bind('<Return>', self._choose)
        self.bind('<space>', self._choose)
        # 使用TkDND原有C/Tcl调用链，不注册Python IDropTarget。
        # 包内组合常量并非底层接受的格式名，必须注册实际的Windows名称。
        if sys.platform == 'win32':
            root = self.winfo_toplevel()
            if not hasattr(root, '_pdf_attachment_cache'):
                root._pdf_attachment_cache = TkAttachmentCache(root)
            self.drop_target_register(DND_FILES, 'FileGroupDescriptorW')
        else:
            self.drop_target_register(DND_FILES)
        self.dnd_bind('<<DropEnter>>', self._drag_action)
        self.dnd_bind('<<DropPosition>>', self._drag_action)
        self.dnd_bind('<<Drop>>', self._drop)
        variable.trace_add('write', self._draw)

    def _draw(self, *_):
        self.delete('all')
        width = self.winfo_width()
        disabled = self.cget('state') == 'disabled'
        color = '#9AA6A6' if disabled else '#007D7E'
        self.create_rectangle(3, 3, width - 3, 109, outline=color, width=2, dash=(8, 5))
        self.create_text(width / 2, 24, text='+', fill=color, font=('Microsoft YaHei UI', 19))
        path = self.variable.get().strip()
        name = Path(path).name if path else f'拖入{self.title}，或点击选择'
        if len(name) > 42:
            name = name[:24] + '…' + name[-15:]
        self.create_text(width / 2, 57, text=name, fill=color,
                         width=max(40, width - 28), font=('Microsoft YaHei UI', 12))
        self.create_text(width / 2, 89,
                         text='仅支持PDF，每次一个文件' if not path else '点击或拖入可重新选择 · 仅支持PDF',
                         fill='#567272', font=('Microsoft YaHei UI', 10))

    def configure(self, cnf=None, **kwargs):
        result = super().configure(cnf, **kwargs)
        if 'state' in kwargs and hasattr(self, 'variable'):
            self._draw()
        return result

    def _blocked(self):
        # 禁用外观不能阻止操作系统投递拖放事件，因此回调也检查忙碌状态。
        return self.is_busy() or self.cget('state') == 'disabled'

    def _drag_action(self, event):
        return REFUSE_DROP if self._blocked() else COPY

    def _accept(self, paths):
        if self._blocked():
            return REFUSE_DROP
        try:
            if len(paths) != 1:
                raise ValueError('每个区域每次只能选择一份PDF文件。')
            path = validate_pdf_path(paths[0])
        except (ValueError, OSError, tk.TclError) as exc:
            messagebox.showwarning('请选择PDF', str(exc), parent=self.winfo_toplevel())
            return REFUSE_DROP
        self.variable.set(str(path))
        return COPY

    def _drop(self, event):
        if self._blocked():
            return REFUSE_DROP
        # event.data是Tcl列表，不能按空格分割，否则中文/空格路径会被拆坏。
        return self._accept(self.tk.splitlist(event.data))

    def _choose(self, event=None):
        if self._blocked():
            return
        path = filedialog.askopenfilename(parent=self.winfo_toplevel(),
                                          title=f'选择{self.title}', filetypes=[('PDF文件', '*.pdf')])
        if path:
            self._accept((path,))
