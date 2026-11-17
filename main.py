"""正式GUI入口；导入不创建窗口或加载OCR模型。"""
import logging
from pathlib import Path
from app.services.settings_service import load_settings
from app.services.runtime_paths import data_root

ROOT = data_root()


def main() -> None:
    log_dir = ROOT / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=log_dir / 'app.log', encoding='utf-8',
                        level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(name)s %(message)s')
    from app.services.single_instance import SingleInstance
    instance = SingleInstance()
    try:
        if not instance.acquire():
            logging.info('阻止重复启动：程序已运行')
            from tkinter import Tk, messagebox
            notice = Tk()
            notice.withdraw()
            try:
                messagebox.showinfo('程序已运行', '程序已运行，请使用现有窗口。识别期间请等待当前任务完成。', parent=notice)
            finally:
                notice.destroy()
            return
        settings = load_settings(ROOT / 'config/settings.json')
        import customtkinter as ctk
        from app.ui.main_window import MainWindow
        ctk.set_appearance_mode('light')
        logging.info('应用启动，供应商=%s', settings.last_supplier)
        window = MainWindow(ROOT, settings)
        window.mainloop()
    except Exception:
        logging.exception('应用启动或运行失败')
        raise
    finally:
        instance.close()


if __name__ == '__main__':
    from multiprocessing import freeze_support
    freeze_support()
    main()
