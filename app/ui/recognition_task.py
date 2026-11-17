"""独立OCR进程；不导入Tk，不传递窗口对象。"""
import logging
import multiprocessing as mp
from pathlib import Path
from time import perf_counter

from app.core.validators import ValidationError
from app.core.workflow import recognize_documents
from app.services.ocr_service import OCRService, DETECTION_MODEL, RECOGNITION_MODEL
from app.services.runtime_paths import model_root

logger = logging.getLogger(__name__)


def recognition_worker(connection, root, supplier, waybill, invoice):
    logging.basicConfig(filename=Path(root) / 'logs/app.log', encoding='utf-8',
                        level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(name)s %(message)s')
    stage = '模型准备'
    started = perf_counter()
    try:
        models = model_root()
        service = OCRService(models / DETECTION_MODEL, models / RECOGNITION_MODEL, cpu_threads=2, recognition_mkldnn=True)
        ready = perf_counter()
        logger.info('OCR模型准备耗时 seconds=%.2f', ready - started)
        stage = '识别'
        connection.send(('status', '正在识别并解析校验两份PDF，请稍候…'))
        records, diagnostics = recognize_documents(supplier, waybill, invoice, service, preview=True)
        if not records and not diagnostics.blockers:
            raise ValidationError('没有可预览的明细。')
        logger.info('识别成功 supplier=%s records=%s recognition_seconds=%.2f total_seconds=%.2f',
                    supplier, len(records), perf_counter() - ready, perf_counter() - started)
        connection.send(('success', (records, diagnostics)))
    except Exception as exc:
        logger.exception('识别失败 stage=%s supplier=%s waybill=%s invoice=%s', stage, supplier, waybill, invoice)
        if stage == '模型准备':
            message = 'OCR模型未准备好，请联系IT或检查本地模型配置。'
        elif isinstance(exc, ValidationError):
            message = f'数据校验失败：{str(exc)[:200]}'
        elif type(exc).__name__ == 'ParseError':
            message = f'识别失败：请确认PDF角色和内容是否正确。{str(exc)[:200]}'
        else:
            message = 'OCR识别失败，请检查PDF和本地OCR模型环境。详细原因见日志。'
        connection.send(('error', message))
    finally:
        connection.close()


class RecognitionTask:
    def __init__(self):
        self.process = None
        self.connection = None
        self.terminal = None

    def start(self, root, supplier, waybill, invoice):
        if self.process is not None:
            raise RuntimeError('已有识别任务。')
        context = mp.get_context('spawn')
        receiver, sender = context.Pipe(duplex=False)
        process = context.Process(target=recognition_worker,
                                  args=(sender, root, supplier, waybill, invoice))
        try:
            process.start()
        except Exception:
            receiver.close()
            sender.close()
            process.close()
            raise
        sender.close()
        self.connection = receiver
        self.process = process
        self.terminal = None

    def poll(self):
        messages = []
        if self.process is None:
            return messages
        # 持续读取管道，避免较大结果阻塞子进程退出；不等待推理完成。
        def drain():
            try:
                while self.connection.poll():
                    message = self.connection.recv()
                    if message[0] == 'status':
                        messages.append(message)
                    else:
                        self.terminal = message
            except (EOFError, BrokenPipeError):
                pass
        drain()
        if not self.process.is_alive():
            self.process.join(timeout=0)
            drain()  # 进程在第一次poll后退出时，读完最后一条消息。
            exitcode = self.process.exitcode
            terminal = self.terminal
            if terminal is None or exitcode != 0:
                logger.error('OCR进程异常退出 exitcode=%s', exitcode)
                terminal = ('error', 'OCR后台进程异常退出，请检查日志及本地OCR环境。')
            self.connection.close()
            self.process.close()
            self.process = None
            self.connection = None
            self.terminal = None
            messages.append(terminal)
        return messages
