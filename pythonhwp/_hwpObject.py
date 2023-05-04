import win32com.client as win32
import shutil
import os
from ._decorator import clearReadState

class hwpObject():
    """
    최하위 Class

    한/글 경로를 받아서 객체를 생성
    """

    # 참고 - Dispatch로 파일 연 후 Gencache 호출하면 기존 라이브러리를 지우게 되어 있음. 반드시 Gencache를 먼저 호출할 것
    # 각 한/글 파일마다 새로운 hwp 객체를 만들게 되므로 hwp.XHwpWindows.Item의 인덱스가 언제나 0번임
    # 만약 NewFile 등의 방식을 이용하여 하나의 객체에서 새로운 한/글을 열었다면 만든 순서대로 인덱스가 부여됨
    # 따라서 해당 item 변수는 신경쓰지 않아도 됨
    item = 0

    def __init__(self, path: str = None, gencache: bool = True) -> None:
        """
        ## 한/글 문서 경로를 인자로 받아 한/글 객체 생성\n
        :param path: 한글 문서 경로
        :param gencache: Early Binding할건지 여부 (기본 True)


        * win32가 한글 파일을 열지 못할 경우 -> gen_py 라이브러리를 삭제 후 시도\n
        gencache.EnsureDispatch의 문제\n
        

        * gencache.EnsureDispatch / Dispatch의 차이\n
        gencache.EnsureDispatch는 Early Binding\n
        Dispatch는 Late Binding\n
        Early Binding은 먼저 로드를 전부 다 하는 방식\n
        gencache.EnsureDispatch가 Early Binding 한 것들을 나중에 못가져오는 문제 존재\n
        Dispatch를 사용 권장\n


        * 중요 - Dispatch 사용하면 InitScan 인자 오류 발생함\n
        ~> 기본 인자(format="HWP", args="") 없는 문제

        한글을 읽어들이는 작업에는 gencache.EnsureDispatch\n
        한글에 출력하는 작업에는 Dispatch 사용 권장\n
        한글 경로 없으면 FileNotFoundError
        """
        self._readState = 0  # InitScan 상태면 1, ReleaseScan 상태면 0
        if not (os.path.isfile(path) or path.endswith(".hwp")):  # 파일이 없으면
            raise FileNotFoundError("한/글 파일이 아닙니다. 파일을 확인해 주세요.")

        if gencache and hwpObject.item == 0:
            try:
                shutil.rmtree(os.path.join(os.path.expanduser('~'), r'AppData\Local\Temp\gen_py\\'))  # 삭제
            except FileNotFoundError:
                pass

            try:
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한글 객체
                self.item = hwpObject.item
                hwpObject.item += 1
            except FileNotFoundError:
                raise Exception("한/글 객체를 생성할 수 없습니다. 프로그램을 다시 실행해 주세요.")

        else:
            self.hwp = win32.Dispatch("HWPFrame.HwpObject")  # 한글 객체
            self.item = hwpObject.item
            hwpObject.item += 1

        self.hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")  # 자동화 보안 모듈
        self.hwp.XHwpWindows.Item(0).Visible = True  # 한글 백그라운드 실행 -> False
        self.hwp.Open(path, "HWP", "versionwarning:false")   # 파일 열기
        self.hwp.HAction.Run("MoveTopLevelBegin")  # 맨 위 페이지로 이동

        # 편집 모드가 아니라면 self.hwp.EditMode = 1로 만들어서 강제 수정 가능
        if self.editMode != 1:
            self.editMode = 1

    def close(self) -> None:
        """
        문서 저장하지 않고 종료
        """
        # 참고 - Class를 이용한 해당 방식은 다른 객체이므로 언제나 index가 0번임
        # 하나의 hwpframe에서 탭을 여러 개 열 때는 각 한/글마다 인덱스가 붙여짐
        # self.hwp.XHwpDocuments.Item(self.item).SetActive_XHwpDocument() # 닫고자 하는 파일로 이동(여러 탭으로 열었을 때)
        # self.hwp.XHwpDocuments.Item(self.item).Close(isDirty=False)
        self.hwp.XHwpDocuments.Item(0).Close(isDirty=False)
        self.hwp.Quit()

    def saveFile(self, path: str, pdfopt: int = 0, quitopt: int = 1) -> None:
        """
        현재 한글 파일을 다른이름으로 저장하고 종료함\n

        path -> 확장자(.hwp)를 포함한 절대 경로여야 함\n
        파일을 열었을 때와 다른 경로로 입력시 원본과 다른 파일로 저장됨\n
        Format을 PDF로 바꾸고 확장자를 .pdf로 주면 PDF로도 저장 가능\n
        한글 종료 시 파일에 변동이 있었다면 저장하겠냐는 창이 나오므로 주의\n

        * pdf로 저장 시 일부 잘림 문제 발생 시\n
        한/글 -> 도구 -> 환경 설정 -> 기타 -> PDF 드라이버 바꿔보기\n
        ezPDF Builder Supreme 사용 시 일부 잘리는 문제 있음\n
        페이지 여백 설정 문제라는데,, 다른 드라이버 사용하는 게 편함

        현재 위치에 그대로 저장하려고 하면 오류 발생, 이미 존재하는 파일을 그대로 저장하면 saveFile_e를 사용할 것
        :param path: 파일 저장 경로
        :param pdfopt: 1 -> pdf로 출력(기본 0)
        :param quitopt: 1 -> 한글 파일 종료(기본 1)
        :return: 0

        >>> hwp.saveFile(path=path, pdfopt=0, quitopt=1)
        """
        if pdfopt:
            self.hwp.HAction.GetDefault("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.Attributes = 0
            self.hwp.HParameterSet.HFileOpenSave.filename = path
            self.hwp.HParameterSet.HFileOpenSave.Format = "PDF"
            self.hwp.HAction.Execute("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
        else:
            self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)  # 파일 저장 액션의 파라미터
            self.hwp.HParameterSet.HFileOpenSave.filename = path
            self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            self.hwp.HParameterSet.HFileOpenSave.Attributes = 0
            self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        if quitopt:
            self.hwp.Quit()

        # hwp.XHwpDocuments.Item(0).Close(isDirty=False)
        # 이 옵션을 수행하면 저장하지 않고 종료함

    def saveFile_e(self, path:str) -> None:
        """
        현재 한글 파일을 저장하고 종료함(이미 존재하는 파일을 그 자리에 저장)\n

        path -> 확장자(.hwp)를 포함한 절대 경로여야 함\n
        파일을 열었을 때와 다른 경로로 입력시 원본과 다른 파일로 저장됨\n

        :param path: 파일 저장 경로        
        """
        self.hwp.HAction.GetDefault("FileSave_S", self.hwp.HParameterSet.HFileOpenSave.HSet)   # 파일 저장 액션의 파라미터를
        self.hwp.HParameterSet.HFileOpenSave.filename = path
        self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        self.hwp.HParameterSet.HFileOpenSave.Attributes = 0
        self.hwp.HAction.Execute("FileSave_S", self.hwp.HParameterSet.HFileOpenSave.HSet)


    @property
    def readState(self) -> int:
        return self._readState

    @readState.setter
    def readState(self, state: int):
        self._readState = state
    
    @property
    def keyIndicator(self) -> tuple:
        """
        현재 포인터 위치의 keyindicator를 반환
        튜플로 반환

        (Boolean, 총 구역, 현재 구역, 쪽, 단, 줄, 칸, 삽입/수정, 컨트롤 이름)
        seccnt : 총 구역
        secno : 현재 구역
        prnpageno : 쪽
        colno : 단
        line : 줄
        pos : 칸
        over : (true:수정, false:삽입)

        페이지 가져오기 -> keyindicator[3]
        :return: (Boolean, 총 구역, 현재 구역, 쪽, 단, 줄, 칸, 삽입/수정, 컨트롤 이름)

        >>> hwp.KeyIndicator[3]
        # 현재 위치 페이지 반환
        """
        return self.hwp.KeyIndicator()

    @property
    def pos(self) -> tuple:
        """
        - 현재 위치를 반환
        - (list, para, pos)의 위치로 이동

        :param curpos: Pos로 가져온 (list,para,pos)
        >>> hwp.Pos = Position
        """
        return self.hwp.GetPos()

    @pos.setter
    @clearReadState
    def pos(self, position: tuple):
        self.hwp.SetPos(*position)
        return

    @property
    def editMode(self) -> int:
        """
        문서의 현재 편집 모드를 반환\n
        0 : 읽기 전용\n
        1 : 일반 편집모드\n
        2 : 양식 모드(양식 사용자 모드) : Cell과 누름틀 중 양식 모드에서 편집 가능 속성을 가진 것만 편집 가능하다.\n
        16 : 배포용 문서 (SetEditMode로 지정 불가능)\n
        1로 지정하여 편집모드로 강제 전환 가능\n
        SetEditMode(0)이라는 기능이 있는 것 같은데 미구현 ?

        :return: 현재 편집 모드
        """
        return self.hwp.EditMode

    @editMode.setter
    def editMode(self, value):
        self.hwp.EditMode = value
