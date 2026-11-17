"""Excel后台线程：只传Path/记录/消息，COM对象不跨线程。"""
import logging
from queue import Queue, Empty
from threading import Thread

from app.services.excel_service import append_records_to_excel1, sync_excel2_batch, ExcelWriteError
from app.services.shipment_output_service import read_master_data, create_shipment_workbook
from app.services.shipment_pdf_service import export_shipment_pdf, export_shipment_package

logger = logging.getLogger(__name__)


class ExcelTask:
    def __init__(self):
        self.messages = Queue()
        self.thread = None

    def start(self, path, records):
        if self.thread is not None:
            raise RuntimeError('Excel任务尚未回收。')
        self.thread = Thread(target=self._worker, args=(path, list(records)), daemon=False)
        self.thread.start()

    def start_master(self, path, records):
        self._start_service(read_master_data, path, list(records))

    def start_output(self, template, root, records, master, image=None, prepare_documents=False):
        args = (template, root, list(records), master)
        if prepare_documents:
            args += (image, True)
        elif image is not None:
            args += (image,)
        self._start_service(create_shipment_workbook, *args)

    def start_pdf(self, source, waybill=None, invoice=None):
        if waybill is None and invoice is None:
            self._start_service(export_shipment_pdf, source)
        else:
            self._start_service(export_shipment_package, source, waybill, invoice)

    def _start_service(self, service, *args):
        if self.thread is not None:
            raise RuntimeError('Excel任务尚未回收。')
        self.thread = Thread(target=self._service_worker, args=(service, args), daemon=False)
        self.thread.start()

    def _service_worker(self, service, args):
        try:
            self.messages.put(('success', service(*args)))
        except Exception as exc:
            logger.exception('后台任务失败 service=%s', service.__name__)
            text = str(exc) if isinstance(exc, (ExcelWriteError, ValueError)) else 'Excel任务失败，请查看日志。'
            self.messages.put(('error', (text, getattr(exc, 'retry_safe', True))))

    def _worker(self, path, records):
        try:
            rows = append_records_to_excel1(path, records)
            self.messages.put(('success', rows))
        except Exception as exc:
            logger.exception('Excel后台写入失败')
            text = str(exc) if isinstance(exc, (ExcelWriteError, ValueError)) else 'Excel写入失败，请查看日志。'
            self.messages.put(('error', (text, getattr(exc, 'retry_safe', True))))

    def start_sync(self, excel1_path, excel2_path, inserted_rows):
        if self.thread is not None:
            raise RuntimeError('Excel任务尚未回收。')
        self.thread = Thread(target=self._sync_worker,
                             args=(excel1_path, excel2_path, list(inserted_rows)), daemon=False)
        self.thread.start()

    def _sync_worker(self, excel1_path, excel2_path, inserted_rows):
        try:
            rows = sync_excel2_batch(excel1_path, excel2_path, inserted_rows)
            self.messages.put(('success', rows))
        except Exception as exc:
            logger.exception('Excel2后台同步失败')
            text = str(exc) if isinstance(exc, (ExcelWriteError, ValueError)) else 'Excel2同步失败，请查看日志。'
            self.messages.put(('error', (text, getattr(exc, 'retry_safe', True))))

    def poll(self):
        # 任务结束且线程已退出后才交付结果，避免COM未释放就允许再次提交。
        if self.thread is None or self.thread.is_alive():
            return []
        self.thread.join()
        self.thread = None
        try:
            return [self.messages.get_nowait()]
        except Empty:
            return [('error', ('Excel任务未返回结果，请人工检查文件，勿重复追加。', False))]
