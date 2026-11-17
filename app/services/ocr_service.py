"""页面像素坐标是唯一OCR坐标系；本层不解释供应商业务字段。"""
import logging
import math
from pathlib import Path
from copy import deepcopy
from threading import Lock
from time import perf_counter

from app.core.models import OCRItem

logger = logging.getLogger(__name__)
_pdfium_lock = Lock()  # PDFium不允许不同线程同时调用，即使文档不同。
DETECTION_MODEL = "PP-OCRv5_mobile_det"
RECOGNITION_MODEL = "en_PP-OCRv5_mobile_rec"


class DocumentOCRError(ValueError):
    """单据无法完整读取，保留出错前已识别的页面。"""

    def __init__(self, message: str, items: list[OCRItem]):
        super().__init__(message)
        self.items = items


def render_pdf_pages(path: Path, dpi: int = 250):
    """逐页返回(1起始页码, BGR图像)，每次仅保留当前页。"""
    if not isinstance(dpi, int) or isinstance(dpi, bool) or not 72 <= dpi <= 600:
        raise ValueError("渲染DPI须为72至600之间的整数。")
    import pypdfium2 as pdfium

    opened = perf_counter()
    with _pdfium_lock:
        document = pdfium.PdfDocument(str(path))
        count = len(document)
    logger.info("PDF打开耗时 file=%s seconds=%.4f", path, perf_counter() - opened)
    try:
        if count == 0:
            raise DocumentOCRError("PDF没有页面。", [])
        for index in range(count):
            rendered = perf_counter()
            with _pdfium_lock:
                page = document[index]
                bitmap = None
                try:
                    bitmap = page.render(scale=dpi / 72, force_bitmap_format=pdfium.raw.FPDFBitmap_BGR)
                    # to_numpy共享原位图内存，必须在关闭位图之前复制。
                    image = bitmap.to_numpy().copy()
                finally:
                    if bitmap is not None:
                        bitmap.close()
                    page.close()
            logger.info("PDF页面渲染耗时 file=%s page=%s seconds=%.4f width=%s height=%s dpi=%s",
                        path, index + 1, perf_counter() - rendered, image.shape[1], image.shape[0], dpi)
            yield index + 1, image
    finally:
        with _pdfium_lock:
            document.close()


def adapt_ocr_result(result, *, page_number: int, width: int, height: int) -> list[OCRItem]:
    """适配PaddleOCR 3.x结果，不过滤低分或改写原文。"""
    texts, scores, boxes = (result[key] for key in ("rec_texts", "rec_scores", "rec_polys"))
    if not len(texts) == len(scores) == len(boxes):
        raise ValueError("OCR文字、分数与坐标数量不一致。")
    items = []
    for text, score, box in zip(texts, scores, boxes):
        points = tuple(tuple(float(value) for value in point) for point in box)
        confidence = float(score)
        if (not isinstance(text, str) or len(points) != 4
                or any(len(point) != 2 for point in points)
                or not math.isfinite(confidence) or not 0 <= confidence <= 1
                or any(not math.isfinite(value) for point in points for value in point)):
            raise ValueError("OCR返回非法文字、分数或四点框。")
        xs, ys = zip(*points)
        if min(xs) < 0 or min(ys) < 0 or max(xs) > width or max(ys) > height:
            raise ValueError("OCR坐标超出渲染页面，须检查缩放或旋转。")
        items.append(OCRItem(text, confidence, page_number, min(xs), min(ys),
                             max(xs), max(ys), points, width, height))
    return items


def recognition_mkldnn_config(config, cpu_threads):
    """已真实验证：检测关闭oneDNN，识别启用；保留现有管线布局参数。"""
    config = deepcopy(config)
    for name, mode in (('TextDetection', 'paddle'), ('TextRecognition', 'mkldnn')):
        options = {'run_mode': mode, 'mkldnn_cache_capacity': 10,
                   'enable_new_ir': True, 'enable_cinn': False}
        if cpu_threads is not None:
            options['cpu_threads'] = cpu_threads
        config['SubModules'][name]['engine'] = 'paddle_static'
        config['SubModules'][name]['engine_config'] = options
    return config


class OCRService:
    def __init__(self, detection_dir: Path, recognition_dir: Path, *, dpi: int = 250, cpu_threads: int | None = None, recognition_mkldnn: bool = False):
        if type(recognition_mkldnn) is not bool:
            raise ValueError('识别MKL-DNN开关必须为bool。')
        if cpu_threads is not None and (type(cpu_threads) is not int or cpu_threads < 1):
            raise ValueError('CPU线程数必须为正整数。')
        # None保留旧阶段引擎默认值；GUI试验显式限制线程，为界面留出CPU资源。
        thread_options = {} if cpu_threads is None else {'cpu_threads': cpu_threads}
        logger.info('OCR初始化 cpu_threads=%s', cpu_threads if cpu_threads is not None else 'engine_default')
        # 先检查本地目录，禁止把缺失模型交给引擎自动下载。
        for directory in (detection_dir, recognition_dir):
            for filename in ("inference.json", "inference.pdiparams", "inference.yml"):
                if not (directory / filename).is_file():
                    raise FileNotFoundError(f"本地OCR模型缺少文件：{directory / filename}")
        from paddleocr import PaddleOCR

        engine_options = {}
        if recognition_mkldnn:
            from paddlex.inference import load_pipeline_config
            engine_options = {
                'engine': 'paddle_static',
                'engine_config': {'run_mode': 'paddle', **thread_options},
                'paddlex_config': recognition_mkldnn_config(load_pipeline_config('OCR'), cpu_threads),
            }
        logger.info('OCR配置 detection_mkldnn=False recognition_mkldnn=%s', recognition_mkldnn)
        self.dpi = dpi
        self.engine = PaddleOCR(
            text_detection_model_name=DETECTION_MODEL,
            text_detection_model_dir=str(detection_dir.resolve()),
            text_recognition_model_name=RECOGNITION_MODEL,
            text_recognition_model_dir=str(recognition_dir.resolve()),
            # 当前Windows环境的oneDNN推理报属性转换错误，使用普通CPU路径验证。
            device="cpu", enable_mkldnn=False, use_doc_orientation_classify=False,
            use_doc_unwarping=False, use_textline_orientation=False,
            text_rec_score_thresh=0.0, **thread_options, **engine_options,
        )  # 禁用几何变换，确保框仍对应输入页面；保留低分证据。

    def recognize_pdf(self, path: Path) -> list[OCRItem]:
        from pypdfium2 import PdfiumError

        started = perf_counter()
        items = []
        empty_pages = []
        pages = render_pdf_pages(path, self.dpi)
        try:
            for page_number, image in pages:
                predicted = perf_counter()
                results = self.engine.predict(image)
                logger.info("OCR整页推理耗时 file=%s page=%s seconds=%.4f",
                            path, page_number, perf_counter() - predicted)
                if len(results) != 1:
                    raise ValueError("单页OCR返回结果数量异常。")
                adapted = perf_counter()
                page_items = adapt_ocr_result(results[0], page_number=page_number,
                                              width=image.shape[1], height=image.shape[0])
                logger.info("OCR结果适配耗时 file=%s page=%s seconds=%.4f items=%s",
                            path, page_number, perf_counter() - adapted, len(page_items))
                if not page_items:
                    empty_pages.append(page_number)
                    continue
                logger.info("OCR页面完成：page=%s items=%s", page_number, len(page_items))
                items.extend(page_items)
        except PdfiumError as exc:
            raise DocumentOCRError(f"PDF 无法完整读取：{exc}", items) from exc
        finally:
            pages.close()
        if empty_pages:
            page_numbers = "、".join(map(str, empty_pages))
            raise DocumentOCRError(f"第{page_numbers}页没有识别到文字。", items)
        logger.info("OCR文件总耗时 file=%s seconds=%.4f items=%s", path, perf_counter() - started, len(items))
        return items
