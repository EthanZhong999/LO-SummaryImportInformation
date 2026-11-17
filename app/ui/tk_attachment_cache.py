"""为TkDND内置的Outlook附件转换提供本进程独占目录。"""
import logging
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

logger = logging.getLogger(__name__)


def enable_classic_file_type(root):
    # 仅注册窗口类型还不够：TkDND supported_types会先按此表过滤来源。
    # 默认Windows映射未启用FileGroupDescriptorW，漏掉此项会始终禁止拖入。
    # 只扩展当前应用解释器，不改site-packages或磁盘上的第三方库。
    root.tk.call('dict', 'set', '::tkdnd::generic::_platform2tkdnd',
                 'FileGroupDescriptorW', 'DND_Files')


class TkAttachmentCache:
    def __init__(self, root):
        self.root = root
        enable_classic_file_type(root)
        self.directory = TemporaryDirectory(prefix='invoice-tkdnd-')
        self.closed = False
        command = root.register(self._new_directory)
        # TkDND在读取虚拟文件内容前调用此过程。每次返回新目录，避免其
        # 内置保存逻辑按附件原名覆盖仍被预览/OCR使用的上一份同名附件。
        root.tk.call('proc', '::tkdnd::GetDropFileTempDirectory', '', command)
        root.bind('<Destroy>', self._destroyed, add='+')

    def _new_directory(self):
        folder = Path(self.directory.name) / uuid4().hex
        folder.mkdir()
        logger.info('TkDND开始接收Outlook附件，临时目录=%s', folder)
        return str(folder)

    def _destroyed(self, event):
        if event.widget is self.root and not self.closed:
            self.closed = True
            try:
                self.directory.cleanup()
            except OSError:
                logger.exception('本次Outlook附件临时目录清理失败，保留现场')
