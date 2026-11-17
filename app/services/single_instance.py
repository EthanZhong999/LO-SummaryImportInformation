"""Windows单实例保护；命名互斥量由系统管理，无锁文件残留。"""
import ctypes
from ctypes import wintypes

MUTEX_NAME = r'Local\LO_Invoice_Auto_Desktop_SingleInstance'
ERROR_ALREADY_EXISTS = 183


class SingleInstance:
    def __init__(self):
        self.handle = None
        self.kernel = None

    def acquire(self) -> bool:
        if self.handle is not None:
            raise RuntimeError('单实例保护已获取。')
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
        kernel.CreateMutexW.restype = wintypes.HANDLE
        kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
        kernel.CloseHandle.restype = wintypes.BOOL
        # 不持有mutex所有权，只保留命名对象句柄；创建和存在检查由Windows原子执行。
        ctypes.set_last_error(0)
        handle = kernel.CreateMutexW(None, False, MUTEX_NAME)
        error = ctypes.get_last_error()
        if not handle:
            raise ctypes.WinError(error)
        if error == ERROR_ALREADY_EXISTS:
            kernel.CloseHandle(handle)
            return False
        self.kernel = kernel
        self.handle = handle
        return True

    def close(self) -> None:
        if self.handle is not None:
            self.kernel.CloseHandle(self.handle)
            self.handle = None
