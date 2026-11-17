"""冻结版资源只读；用户配置和日志写入LocalAppData，避免安装目录权限问题。"""
import os
import sys
from pathlib import Path


def resource_root() -> Path:
    return Path(__file__).resolve().parents[2]


def data_root() -> Path:
    if getattr(sys, 'frozen', False):
        if os.environ.get('INVOICE_CLASSIC_TEST_PACKAGE') == '1':
            return Path(os.environ['LOCALAPPDATA']) / 'LO_Invoice_Auto_ClassicTest'
        return Path(os.environ['LOCALAPPDATA']) / 'LO_Invoice_Auto'
    return resource_root()


def model_root() -> Path:
    if getattr(sys, 'frozen', False):
        return resource_root() / 'models'
    return Path.home() / '.paddlex/official_models'
