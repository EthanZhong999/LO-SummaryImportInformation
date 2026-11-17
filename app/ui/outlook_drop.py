"""Windows OLE拖放：普通路径、Classic虚拟附件、新版异步文件。

只接收用户拖来的IDataObject，不连接Outlook账号。异步对象经COM封送交给
工作线程，结果通过队列回到Tk线程；不能直接跨线程使用COM代理或Tk控件。
"""
import logging
from pathlib import Path
from queue import Queue
import struct
from tempfile import TemporaryDirectory
from threading import Thread

import pythoncom
import win32clipboard
from win32com.server.util import wrap
from win32com.shell import shell

from app.ui.gui_logic import validate_pdf_path

logger = logging.getLogger(__name__)
COPY = 1
NONE = 0
MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024
E_NOINTERFACE = -2147467262
E_FAIL = -2147467259


def descriptor_name(data: bytes, unicode=True) -> str:
    """FILEGROUPDESCRIPTOR首DWORD为数量，文件名在FILEDESCRIPTOR偏移72。"""
    size = 592 if unicode else 332
    if len(data) < 4 or struct.unpack_from('<I', data)[0] != 1:
        raise ValueError('每个区域每次只能拖入一份PDF附件。')
    if len(data) < 4 + size:
        raise ValueError('附件描述数据不完整，请先另存为PDF再选择。')
    raw = data[76:76 + (520 if unicode else 260)]
    name = raw.decode('utf-16-le' if unicode else 'mbcs').split('\0', 1)[0]
    # Windows文件名不能包含目录或设备路径，防止写到临时目录之外。
    if (not name or name in ('.', '..') or any(c in name for c in '\\/:*?"<>|')
            or any(ord(c) < 32 for c in name) or name.endswith((' ', '.'))
            or Path(name).suffix.lower() != '.pdf'):
        raise ValueError('请拖入文件名有效的PDF附件，不支持整封邮件或文件夹。')
    if name.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *('COM'+str(i) for i in range(1, 10)), *('LPT'+str(i) for i in range(1, 10))}:
        raise ValueError('附件文件名为Windows保留名称，请先另存并重命名。')
    return name


def hdrop_paths(data: bytes) -> tuple[str, ...]:
    if len(data) < 20:
        raise ValueError('拖入的文件路径数据不完整。')
    offset, _, _, _, wide = struct.unpack_from('<IiiII', data)
    if offset < 20 or offset >= len(data):
        raise ValueError('拖入的文件路径数据无效。')
    raw = data[offset:]
    step = 2 if wide else 1
    terminator = b'\0' * (2 * step)
    # HGLOBAL容量可大于有效载荷（也可能多出奇数个填充字节），只解码
    # 双NUL结束之前的有效路径；UTF-16的结束符必须位于代码单元边界。
    for end in range(0, len(raw) - len(terminator) + 1, step):
        if raw[end:end + len(terminator)] == terminator:
            text = raw[:end].decode('utf-16-le' if wide else 'mbcs')
            return tuple(p for p in text.split('\0') if p)
    raise ValueError('拖入的文件路径列表缺少结束符。')


class AttachmentFiles:
    """仅清理本次进程创建的临时副本，不删除用户原文件或其他会话文件。"""
    def __init__(self):
        self.directory = None

    def save(self, name, chunks):
        if self.directory is None:
            self.directory = TemporaryDirectory(prefix='invoice-attachment-')
        # 每次一层独立目录，重复文件名也不会覆盖之前仍在使用的附件。
        from uuid import uuid4
        folder = Path(self.directory.name) / uuid4().hex
        folder.mkdir()
        target = folder / name
        try:
            size = 0
            with target.open('xb') as stream:
                for chunk in chunks:
                    size += len(chunk)
                    if size > MAX_ATTACHMENT_BYTES:
                        raise ValueError('PDF附件超过50MB，请先另存后通过点击选择。')
                    stream.write(chunk)
            with target.open('rb') as stream:
                if b'%PDF-' not in stream.read(1024):
                    raise ValueError('附件内容不是有效的PDF文件头，请检查原附件。')
            return target
        except Exception:
            target.unlink(missing_ok=True)
            folder.rmdir()
            raise

    def close(self):
        if self.directory is not None:
            self.directory.cleanup()
            self.directory = None


def format_etc(cf, tymed=pythoncom.TYMED_HGLOBAL, index=-1):
    return cf, None, pythoncom.DVASPECT_CONTENT, index, tymed


class OutlookDropTarget:
    _com_interfaces_ = [pythoncom.IID_IDropTarget]
    _public_methods_ = ['DragEnter', 'DragOver', 'DragLeave', 'Drop']

    def __init__(self, widget, blocked, accept, notify, changed):
        self.widget = widget
        self.blocked = blocked
        self.accept = accept
        self.notify = notify
        self.changed = changed
        self.files = AttachmentFiles()
        self.queue = Queue()
        self.receiving = False
        self.operation = None
        self.supported = False
        self.closed = False
        self.descriptor_w = win32clipboard.RegisterClipboardFormat('FileGroupDescriptorW')
        self.descriptor_a = win32clipboard.RegisterClipboardFormat('FileGroupDescriptor')
        self.contents = win32clipboard.RegisterClipboardFormat('FileContents')
        self.hwnd = widget.winfo_id()
        self.gateway = wrap(self, pythoncom.IID_IDropTarget)
        pythoncom.OleInitialize()
        try:
            pythoncom.RegisterDragDrop(self.hwnd, self.gateway)
        except Exception:
            pythoncom.CoUninitialize()
            raise
        self.after_id = widget.after(80, self._poll)

    def _formats(self, data):
        return {entry[0] for entry in data.EnumFormatEtc(pythoncom.DATADIR_GET)}

    def DragEnter(self, data, keys, point, effect):
        try:
            formats = self._formats(data)
            self.supported = bool(formats & {15, self.descriptor_w, self.descriptor_a})
            logger.debug('拖放格式ID：%s', sorted(formats))
            return self.DragOver(keys, point, effect)
        except Exception:
            logger.exception('检查拖放数据格式失败')
            self.supported = False
            return NONE

    def DragOver(self, keys, point, effect):
        return COPY if self.supported and not self.blocked() and not self.receiving and effect & COPY else NONE

    def DragLeave(self):
        self.supported = False

    def _extract(self, data, copy_paths):
        formats = self._formats(data)
        if self.descriptor_w in formats or self.descriptor_a in formats:
            wide = self.descriptor_w in formats
            descriptor = data.GetData(format_etc(self.descriptor_w if wide else self.descriptor_a))
            name = descriptor_name(descriptor.data, wide)
            medium = data.GetData(format_etc(self.contents,
                                 pythoncom.TYMED_ISTREAM | pythoncom.TYMED_HGLOBAL, 0))
            if medium.tymed == pythoncom.TYMED_ISTREAM:
                stream = medium.data
                def chunks():
                    while chunk := stream.Read(64 * 1024):
                        yield chunk
                return self.files.save(name, chunks())
            if medium.tymed == pythoncom.TYMED_HGLOBAL:
                return self.files.save(name, (medium.data,))
            raise ValueError('附件内容格式暂不支持，请先另存为PDF。')
        if 15 in formats:
            medium = data.GetData(format_etc(15))
            paths = hdrop_paths(medium.data)
            if len(paths) != 1:
                raise ValueError('每个区域每次只能拖入一份PDF文件。')
            path = validate_pdf_path(paths[0])
            if not copy_paths:
                return path
            # 新版Outlook可能提供短暂存在的下载路径，在EndOperation前保留副本。
            with path.open('rb') as stream:
                return self.files.save(path.name, iter(lambda: stream.read(64 * 1024), b''))
        raise ValueError('此拖放来源未提供PDF文件，请先另存后选择。')

    def Drop(self, data, keys, point, effect):
        if self.blocked() or self.receiving or not effect & COPY:
            return NONE
        operation = None
        try:
            try:
                operation = data.QueryInterface(shell.IID_IAsyncOperation)
            except pythoncom.com_error as exc:
                # 不支持可选异步接口是正常协商结果，不是COM环境故障。
                if exc.hresult != E_NOINTERFACE:
                    raise
            if operation is not None and operation.GetAsyncMode():
                operation.StartOperation(None)
                try:
                    marshaled = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDataObject, data)
                    self.receiving = True
                    self.operation = operation
                    self.changed()
                    Thread(target=self._receive, args=(marshaled,), daemon=False).start()
                except Exception:
                    self.receiving = False
                    self.operation = None
                    operation.EndOperation(E_FAIL, None, NONE)
                    self.changed()
                    raise
                return COPY
            # 同步源的内容必须在Drop返回前读完，否则Outlook可立即释放附件。
            path = self._extract(data, copy_paths=False)
            return COPY if self.accept((str(path),)) == 'copy' else NONE
        except Exception:
            logger.exception('接收Outlook/文件拖放失败')
            self.queue.put((None, '无法接收拖入的PDF附件，请先另存后选择；详细原因见日志。'))
            return NONE
        finally:
            operation = None

    def _receive(self, marshaled):
        path = None
        error = None
        data = None
        initialized = False
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            initialized = True
            data = pythoncom.CoGetInterfaceAndReleaseStream(marshaled, pythoncom.IID_IDataObject)
            marshaled = None
            path = self._extract(data, copy_paths=True)
        except Exception:
            logger.exception('异步接收Outlook附件失败')
            error = 'Outlook附件接收失败，请先另存为PDF后选择；详细原因见日志。'
        finally:
            data = marshaled = None
            if initialized:
                pythoncom.CoUninitialize()
            self.queue.put((path, error))

    def _poll(self):
        if self.closed:
            return
        while not self.queue.empty():
            path, error = self.queue.get_nowait()
            if self.operation is not None:
                try:
                    # 在StartOperation所属线程收尾；即使工作线程解封送失败也须配对结束。
                    self.operation.EndOperation(0 if error is None else E_FAIL, None, COPY if error is None else NONE)
                except Exception:
                    logger.exception('结束Outlook附件传输失败')
                    error = 'Outlook附件传输结束状态异常，请先另存为PDF后选择。'
                finally:
                    self.operation = None
            self.receiving = False
            self.changed()
            if error:
                self.notify(error)
            elif path:
                self.accept((str(path),))
        self.after_id = self.widget.after(80, self._poll)

    def close(self):
        if self.closed:
            return
        self.closed = True
        self.widget.after_cancel(self.after_id)
        pythoncom.RevokeDragDrop(self.hwnd)
        self.gateway = None
        self.files.close()
        pythoncom.CoUninitialize()
