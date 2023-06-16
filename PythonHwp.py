from .pythonhwp.HwpUserMethods import HwpUserMethods

__all__ = ["open"]

class _PythonHwp(HwpUserMethods):
    """
    한/글 객체를 생성
    """
    def __init__(self, path: str = None, gencache: bool = True) -> None:
        super().__init__(path, gencache)
        self.HWPPATH = path
        """ 한글 파일 경로 저장 """


def openhwp(path: str = None, gencache: bool = True) -> _PythonHwp:
    """
    한/글 파일을 여는 function
    """
    return _PythonHwp(path, gencache)