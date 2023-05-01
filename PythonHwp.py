from .pythonhwp._hwpObject import hwpObject
from .pythonhwp.keyBinding import hwpKeyBinding
from .pythonhwp.readWrite import hwpReadWrite
from .pythonhwp.utility import hwpUtility
from .pythonhwp.userMethod import hwpUserMethod

__all__ = ["open"]

class _PythonHwp(hwpObject):
    """
    한/글 객체를 생성
    """
    def __init__(self, path: str = None, gencache: bool = True) -> None:
        super().__init__(path, gencache)
        self.keyBinding = hwpKeyBinding(self)
        self.readWrite = hwpReadWrite(self)
        self.utility = hwpUtility(self)
        self.userMethod = hwpUserMethod(self)


def openhwp(path: str = None, gencache: bool = True) -> _PythonHwp:
    """
    한/글 파일을 여는 function
    """
    return _PythonHwp(path, gencache)