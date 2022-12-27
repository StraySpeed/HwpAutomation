import win32com.client as win32
import shutil
import os
import warnings

"""
pointNotMoved 같은걸 쓰지 말고 아까 읽어들인 위치 + 값을 저장했다가
이를 비교하는 방식이 낫지 않은가 ?
훨씬 더 정확하고 읽는데 문제도 없는 방식이긴 한데 고민하기
>>> @_clearReadState 사용하기
"""
class PythonHwp():
    """
    자주 사용한 한글 기능들을 클래스로 저장\n
    이 클래스를 사용하거나, 혹은 이를 참고하여 원하는대로 사용하기
    """

    # 참고 - Dispatch로 파일 연 후 Gencache 호출하면 기존 라이브러리를 지우게 되어 있음. 반드시 Gencache를 먼저 호출할 것
    # 각 한/글 파일마다 새로운 PythonHwp 객체를 만들게 되므로 hwp.XHwpWindows.Item의 인덱스가 언제나 0번임
    # 만약 NewFile 등의 방식을 이용하여 하나의 객체에서 새로운 한/글을 열었다면 만든 순서대로 인덱스가 부여됨
    # 따라서 해당 item 변수는 신경쓰지 않아도 됨
    item = 0

    def __init__(self, path: str = None, gencache: bool = True) -> None:
        """
        한/글 문서 경로를 인자로 받아 한/글 객체 생성\n
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
         한글을 읽어들이는 작업에는 gencache.EnsureDispatch\n
         한글에 출력하는 작업에는 Dispatch 사용 권장\n
         한글 경로 없으면 FileNotFoundError
        """
        if path is None:
            raise FileNotFoundError

        if gencache and PythonHwp.item == 0:
            try:
                shutil.rmtree(os.path.join(os.path.expanduser('~'), r'AppData\Local\Temp\gen_py\\'))  # 삭제
            except FileNotFoundError:
                pass

            try:
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한글 객체
                self.item = PythonHwp.item
                PythonHwp.item += 1
            except FileNotFoundError:
                raise Exception("한/글 객체를 생성할 수 없습니다. 프로그램을 다시 실행해 주세요.")

        else:
            self.hwp = win32.Dispatch("HWPFrame.HwpObject")  # 한글 객체
            self.item = PythonHwp.item
            PythonHwp.item += 1

        if os.path.isfile(path) and path.endswith(".hwp"):    # 파일이 존재할 경우
            self.hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")  # 자동화 보안 모듈
            self.hwp.XHwpWindows.Item(0).Visible = True  # 한글 백그라운드 실행 -> False
            self.hwp.Open(path, "HWP", "versionwarning:false")   # 파일 열기
            self.hwp.HAction.Run("MoveTopLevelBegin")  # 맨 위 페이지로 이동
        else:   # 파일이 존재하지 않을 경우
            raise FileNotFoundError

        self.readState = 0  # InitScan 상태면 1, ReleaseScan 상태면 0

        # 파일이 편집 모드가 아니면 나오는 경고. 근데 한글 파일 열리는 시간이 있어서 그런지 항상 해당 경고가 나옴
        # 편집 모드가 아니라면 self.hwp.EditMode = 1로 만들어서 강제 수정 가능
        if self._editMode != 1:
            warnings.warn("File is not Editmode.")

    # decorator
    def _clearReadState(func):
        """ 
        # ReleaseScan decorator
        강제로 ReleaseScan을 시행
        이렇게 구현하고 싶지 않았는데,,,
        어쩔 수 없는 듯 함
        
        한/글 캐럿이 움직이는 method에 붙여 놓을 것
        """
        def inner_function(*args, **kwargs):
            self = args[0]  # class 함수의 첫 인자는 언제나 self
            self.hwp.ReleaseScan()
            self.readState = 0
            func(*args, **kwargs)
        return inner_function

    def close(self) -> int:
        """
        문서 저장하지 않고 종료

        :return: 0
        """
        # 참고 - Class를 이용한 해당 방식은 다른 객체이므로 언제나 index가 0번임
        # 하나의 hwpframe에서 탭을 여러 개 열 때는 각 한/글마다 인덱스가 붙여짐
        # self.hwp.XHwpDocuments.Item(self.item).SetActive_XHwpDocument() # 닫고자 하는 파일로 이동(여러 탭으로 열었을 때)
        # self.hwp.XHwpDocuments.Item(self.item).Close(isDirty=False)
        self.hwp.XHwpDocuments.Item(0).Close(isDirty=False)
        self.hwp.Quit()
        return 0

    def _readtest(self) -> None:
        """
        한글 파일 출력 TEST용

        >>> hwp._readtest()
        """
        self.hwp.InitScan(option=None, Range=0x0077)

        while True:
            result = self.hwp.GetText()  # 문단별로 텍스트와 상태코드 얻기
            if result[0] == 1:  # 상태코드1 == 문서 끝에 도달하면
                break  # while문 종료
            if result[0] == 0:  # 텍스트 정보 없음
                break
            result1 = result[1].strip()  # 텍스트만 추출
            print(result1)
        self.hwp.ReleaseScan()
        self.hwp.Quit()
        return

    @_clearReadState
    def insertLine(self, text: str) -> int:
        """
        한글에 텍스트를 입력

        입력할 위치로 포인터를 옮긴 후 실행
        text -> 문자열 형식
        hwp.HAction.Run("BreakPara") 추가

        :return: 0
        
        >>> hwp.insertLine("텍스트")
        """
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", str(text))
        act.Execute(set)
        self.hwp.HAction.Run("BreakPara")
        return 0

    @_clearReadState
    def insertLinebyField(self, text: str) -> int:
        """
        한글에 텍스트를 입력(누름틀 버전)

        누름틀을 이용할 경우 포인터가 입력한 문자 끝으로 넘어가지 않음

        :return: 0

        >>> hwp.insertLinebyField("텍스트")
        """
        self.hwp.CreateField(Direction="입력칸", memo="텍스트 입력", name="textarea")
        self.hwp.PutFieldText("textarea", str(text))
        self.hwp.Run("DeleteField")
        # self.hwp.Run("MoveNextParaBegin")
        return 0

    @_clearReadState
    def insertPicture(self, picturepath: str) -> int:
        """
        한글에 이미지를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행\n

        :param picturepath: 이미지 파일의 전체 경로
        :return: 0
        """
        self.hwp.InsertPicture(picturepath, True, 0, 0, 0, 0)   # 원래 크기로, 반전 X, 워터마크 X, 실제 이미지 그대로
        return 0
        
#        === 참고용 ===
#        이미지 객체 속성을 변경할 경우
#        
#        hwp.FindCtrl()  # 현재 포인터에 인접한 개체 선택 (양쪽에 존재하면 우측 우선)
#        
#        # 이미지 속성 변경
#        hwp.HAction.GetDefault("FormObjectPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
#        hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 액션 초기화
#        hwp.HParameterSet.HShapeObject.TextWrap = hwp.TextWrapType("BehindText")  # 글 뒤로 배치
#        hwp.HParameterSet.HShapeObject.TreatAsChar = 0  # 글자처럼 취급 해제
#        hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 실행    
    

    def readLine_tuple(self, opt :int=None, ran :int =0x0077, statesave :int = 1) -> tuple:
        """         한글에서 한 줄을 읽어오기\n

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임\n
        읽어들인 text를 반환\n
        텍스트를 읽어들이고 읽어들인 위치로 이동\n
        공백을 포함해서 읽어들임, 튜플 형태를 반환\n
        중간에 텍스트에 변동이 있으면 _clearReadState를 줄 것

        :param opt: 읽어들이는 옵션(기본 None)
        :param ran: 범위(기본 문서 시작-끝까지)
        :param saveprev: InitScan 초기화 여부 (초기화하고자 하면 0)
        :return: 읽어들인 튜플(없으면 (-1, None))

        >>> hwp.readLine_tuple(opt=0, ran=0x0033) """

        
        # 문서의 맨 마지막이라면 None
        #if self._getPos() == self._findLastPos():
        #    return (-1, None)

        # readState가 아니라면 InitScan() 호출
        if not self.readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self.readState = 1

        texttuple = self.hwp.GetText()

        if texttuple[0] == 1 or texttuple[0] == 0:
            return (-1, None)
        
        self.hwp.MovePos(201)   # 읽어들인 위치로 이동

        # readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self.readState and not statesave:
            self.hwp.ReleaseScan()
            self.readState = 0

        return texttuple


    def readLine(self, opt :int=None, ran :int =0x0077, statesave :int = 1) -> None | str:
        """
        한글에서 한 줄을 읽어오기\n

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임\n
        읽어들인 text를 반환\n
        텍스트를 읽어들이고 읽어들인 위치로 이동\n
        공백없이 읽어들임

        :param opt: 읽어들이는 옵션(기본 None)
        :param ran: 범위(기본 문서 시작-끝까지)
        :param statesave: 상태 저장 여부
        :return: 읽어들인 텍스트(없으면 None)

        >>> hwp.readLine(opt=0, ran=0x0033)
        """

        # readState가 아니라면 InitScan() 호출
        if not self.readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self.readState = 1

        text = (-1, '')
        while text[1].strip() == '':
            text = self.hwp.GetText()

            if text[0] == 1 or text[0] == 0:
                return None
        
        self.hwp.MovePos(201)   # 읽어들인 위치로 이동

        # readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self.readState and not statesave:
            self.hwp.ReleaseScan()
            self.readState = 0

        return text[1]

    def setNewNumber(self, num: int) -> None:
        """
        미주 번호를 바꾸는 함수\n

        한글 -> 새 번호로 시작 -> 미주 번호\n
        * 내부에서 num을 정수로 바꾸도록 되어 있음

        :param num: 바꿀 번호
        """
        self.hwp.HAction.GetDefault("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        self.hwp.HParameterSet.HAutoNum.NumType = self.hwp.AutoNumType("Endnote")
        self.hwp.HParameterSet.HAutoNum.NewNumber = int(num)
        self.hwp.HAction.Execute("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        return

    @_clearReadState
    def findNumber(self) -> tuple:
        """
        미주로 이동하고 미주의 앞 위치를 반환\n
        # GetPos가 반환이 안 되는 문제????

        한글 내 포인터 미주 앞으로 이동함\n
        문서의 맨 마지막 미주 너머에서 이를 실행 시 처음부터 재탐색하겠냐는 메시지 창이 뜸 -> 루프 시 탈출 조건 필요\n
        -> hwp.SetMessageBoxMode(0x10000)\n
        미주의 앞 위치를 반환하므로 루프 내에서 미주 탐색 시 동일 위치를 무한루프 할 수 있음에 주의\n
        hwp.HAction.Run("MoveNextParaBegin")으로 현재 문장을 넘어간 후 실행할 것\n
        또는 다른 방식으로 포인터 위치를 옮긴 후 실행할 것\n
        :return: (list, para, pos)
        """
        # self.hwp.SetMessageBoxMode(0x10000)  # 예/아니오 창에서 "예"를 누르는 method
        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
        self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
        self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)

        return self.hwp.GetPos()  # 미주로 이동 후 현재 위치를 반환

    @_clearReadState
    def insertEndnote(self, text: str) -> int:
        """
        미주를 삽입하는 함수\n

        :param text: 삽입할 내용
        :return: 0
        """
        self.hwp.HAction.Run("InsertEndnote")  # 미주 삽입
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", str(text))
        act.Execute(set)
        self.hwp.HAction.Run("CloseEx") # 원래 위치로 돌아감
        return 0

    def saveFile(self, path: str, pdfopt: int = 0, quitopt: int = 1) -> int:
        """
        현재 한글 파일을 다른이름으로 저장하고 종료함\n

        path -> 확장자(.hwp)를 포함한 절대 경로여야 함\n
        파일을 열었을 때와 다른 경로로 입력시 원본과 다른 파일로 저장됨\n
        Format을 PDF로 바꾸고 확장자를 .pdf로 주면 PDF로도 저장 가능\n
        한글 종료 시 파일에 변동이 있었다면 저장하겠냐는 창이 나오므로 주의\n

        현재 위치에 그대로 저장하려고 하면 오류 발생, saveFile_e를 사용할 것
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
        return 0

    def saveFile_e(self, path:str) -> int:
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
        return 0

    def allreplace(self, findtext: str, changetext: str, regex: int = 0) -> None:
        """
        한글 -> 찾아 바꾸기를 실행\n

        findtext에 이를 넣고 changetext에 바꿀 머리말을 넣어서 머리말 변경 가능\n
        바꾸기 = RepeatFind, 모두 바꾸기 = AllReplace\n

        :param findtext: 찾을 문자열
        :param changetext: 바꿀 문자열
        :param regex: 정규표현식 사용 여부(기본 0)

        >>> hwp.allreplace(findtext="\[(.+)\]", changetext="", regex=1)
        """
        # 모두 바꾸기
        self.hwp.HAction.GetDefault("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)

        self.hwp.HParameterSet.HFindReplace.MatchCase = 0
        self.hwp.HParameterSet.HFindReplace.AllWordForms = 0
        self.hwp.HParameterSet.HFindReplace.SeveralWords = 0
        self.hwp.HParameterSet.HFindReplace.UseWildCards = 0
        self.hwp.HParameterSet.HFindReplace.WholeWordOnly = 0
        self.hwp.HParameterSet.HFindReplace.AutoSpell = 1
        # Forward -> 위에서 아래로, Backward -> 아래에서 위로
        self.hwp.HParameterSet.HFindReplace.Direction = self.hwp.FindDir("Forward")
        self.hwp.HParameterSet.HFindReplace.IgnoreFindString = 0
        self.hwp.HParameterSet.HFindReplace.IgnoreReplaceString = 0
        self.hwp.HParameterSet.HFindReplace.FindString = findtext
        self.hwp.HParameterSet.HFindReplace.ReplaceString = changetext

#        # 밑줄 긋는 옵션
#        self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineType = self.hwp.HwpUnderlineType("Bottom")
#        self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineColor = 0
#        self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineShape = self.hwp.HwpUnderlineShape("Solid")
#
#        # 굵은 글씨 옵션
#        self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.Bold = 1
#
#        # 찾아 바꾸기를 쓰려면?
#        # ReplaceCharShape -> FindCharShape

        self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1
        self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        self.hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0
        self.hwp.HParameterSet.HFindReplace.FindJaso = 0
        # 정규 표현식으로 찾는 옵션 -> 1
        self.hwp.HParameterSet.HFindReplace.FindRegExp = regex
        self.hwp.HParameterSet.HFindReplace.FindStyle = ""
        self.hwp.HParameterSet.HFindReplace.ReplaceStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindType = 1
        # 모두 바꾸기
        self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)

        return

    @_clearReadState
    def find(self, text: str, regex: int = 0) -> int:
        """
        한글 -> 찾기를 실행\n

        정규표현식을 이용하려면 1, 아니면 0을 대입\n
        한 바퀴를 돌면 hwp.SelectionMode == 0이 됨
        한 바퀴를 돌면 0, 아니면 1을 반환

        :param text: 찾을 문자열
        :param regex: 정규표현식 사용 여부(기본 0)
        :return: 문서의 끝에 도달하면 0, 아니면 1
        >>> flag = hwp.find(text="텍스트", regex=0)
        """
        # 반복 찾기 방법
        self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)

        self.hwp.HParameterSet.HFindReplace.ReplaceString = ""
        self.hwp.HParameterSet.HFindReplace.FindString = text
        self.hwp.HParameterSet.HFindReplace.IgnoreReplaceString = 0
        self.hwp.HParameterSet.HFindReplace.IgnoreFindString = 0
        self.hwp.HParameterSet.HFindReplace.Direction = self.hwp.FindDir("Forward")
        self.hwp.HParameterSet.HFindReplace.WholeWordOnly = 0
        self.hwp.HParameterSet.HFindReplace.UseWildCards = 0
        self.hwp.HParameterSet.HFindReplace.SeveralWords = 0
        self.hwp.HParameterSet.HFindReplace.AllWordForms = 0
        self.hwp.HParameterSet.HFindReplace.MatchCase = 0
        self.hwp.HParameterSet.HFindReplace.ReplaceMode = 0
        self.hwp.HParameterSet.HFindReplace.ReplaceStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindRegExp = regex   # 정규표현식으로 찾을 경우
        self.hwp.HParameterSet.HFindReplace.FindJaso = 0
        self.hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0
        self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        self.hwp.HParameterSet.HFindReplace.FindType = 1

        self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
        return self.hwp.SelectionMode

    def textStyle1(self, bold: int = 0, italic: int = 0, underline: int = 0, strikeline: int = 0) -> None:
        """
        글씨 스타일 굵게(bold), 기울임(italic), 밑줄(underline), 취소선(strikeline)을 조정\n

        1 -> 적용함, 0 -> 적용안함\n
        *기본은 적용하지 않음*\n
        :param bold: 굵게
        :param italic: 기울임
        :param underline: 밑줄
        :param strikeline: 취소선
        >>> hwp.textStyle1(bold=1, italic=0)
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        if Set.Item("Bold") ^ bold:     # XOR 연산
            self.hwp.HAction.Run("CharShapeBold")
        if Set.Item("Italic") ^ italic:     # XOR 연산
            self.hwp.HAction.Run("CharShapeItalic")
        if Set.Item("UnderlineType") ^ underline:
            self.hwp.HAction.Run("CharShapeUnderline")
        if Set.Item("StrikeOutType") ^ strikeline:
            self.hwp.HAction.Run("CharShapeStrikeout")

        return

#        
#        if underline == 1:
#            self.hwp.HAction.Run("CharShapeUnderline")
#        if strikeline == 1:
#            self.hwp.HAction.Run("CharShapeStrikeout")
#        
#
#        # 사용을 추천하지는 않으나, 혹시 글자 스타일 또는 크기로 조건을 판단하는 경우의 예시
#
#        Act = hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
#        Set = Act.CreateSet()  # 세트 생성
#        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
#        # 켜짐 -> 1, 꺼짐 -> 0
#        if Set.Item("Bold") == 1:   # 굵은 글씨 옵션이 켜져 있는지 여부
#            pass
#        if Set.Item("Italic") == 1:   # 기울임 옵션이 켜져 있는지 여부
#            pass
#        if Set.Item("Height") > 900:    # 글씨 크기가 9.00pt 초과하는 경우
#            pass
#


    def textStyle2(self, color: tuple = (255, 255, 255), font: str = "바탕", size: float = 9.5) -> None:
        """
        글씨 스타일 font, size, color(RGB) 지정
        또는 선택 후 텍스트 스타일 지정

        :param color: 색 지정(RGB)
        :param font: 글씨체(기본 "바탕")
        :param size: 폰트 크기(기본 9.5)
        >>> hwp.textStyle2(color=(255, 0, 0), font="나눔스퀘어라운드 Regular", size=9.5)
        """
        # 글자 모양 - 글꼴종류
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.FaceNameUser = font
        self.hwp.HParameterSet.HCharShape.FaceNameSymbol = font
        self.hwp.HParameterSet.HCharShape.FaceNameOther = font
        self.hwp.HParameterSet.HCharShape.FaceNameJapanese = font
        self.hwp.HParameterSet.HCharShape.FaceNameHanja = font
        self.hwp.HParameterSet.HCharShape.FaceNameLatin = font
        self.hwp.HParameterSet.HCharShape.FaceNameHangul = font

        # 글자 모양 - 폰트 타입
        self.hwp.HParameterSet.HCharShape.FontTypeUser = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeSymbol = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeOther = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeJapanese = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeHanja = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeLatin = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.FontTypeHangul = self.hwp.FontType("TTF")

        # 글자 모양 - 상대크기 %
        self.hwp.HParameterSet.HCharShape.SizeUser = 100
        self.hwp.HParameterSet.HCharShape.SizeSymbol = 100
        self.hwp.HParameterSet.HCharShape.SizeOther = 100
        self.hwp.HParameterSet.HCharShape.SizeJapanese = 100
        self.hwp.HParameterSet.HCharShape.SizeHanja = 100
        self.hwp.HParameterSet.HCharShape.SizeLatin = 100
        self.hwp.HParameterSet.HCharShape.SizeHangul = 100

        # 글자 모양 - 장평 %
        self.hwp.HParameterSet.HCharShape.RatioUser = 100
        self.hwp.HParameterSet.HCharShape.RatioSymbol = 100
        self.hwp.HParameterSet.HCharShape.RatioOther = 100
        self.hwp.HParameterSet.HCharShape.RatioJapanese = 100
        self.hwp.HParameterSet.HCharShape.RatioHanja = 100
        self.hwp.HParameterSet.HCharShape.RatioLatin = 100
        self.hwp.HParameterSet.HCharShape.RatioHangul = 100

        # 글자 모양 - 자간 %
        self.hwp.HParameterSet.HCharShape.SpacingUser = 0
        self.hwp.HParameterSet.HCharShape.SpacingSymbol = 0
        self.hwp.HParameterSet.HCharShape.SpacingOther = 0
        self.hwp.HParameterSet.HCharShape.SpacingJapanese = 0
        self.hwp.HParameterSet.HCharShape.SpacingHanja = 0
        self.hwp.HParameterSet.HCharShape.SpacingLatin = 0
        self.hwp.HParameterSet.HCharShape.SpacingHangul = 0

        # 글자 모양 - 글자위치 %
        self.hwp.HParameterSet.HCharShape.OffsetUser = 0
        self.hwp.HParameterSet.HCharShape.OffsetSymbol = 0
        self.hwp.HParameterSet.HCharShape.OffsetOther = 0
        self.hwp.HParameterSet.HCharShape.OffsetJapanese = 0
        self.hwp.HParameterSet.HCharShape.OffsetHanja = 0
        self.hwp.HParameterSet.HCharShape.OffsetLatin = 0
        self.hwp.HParameterSet.HCharShape.OffsetHangul = 0

        self.hwp.HParameterSet.HCharShape.Height = self.hwp.PointToHwpUnit(size)
        self.hwp.HParameterSet.HCharShape.TextColor = self.hwp.RGBColor(*color)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
         
    @_clearReadState
    def deleteLine(self) -> None:
        """
        현재 줄(엔터 전까지를) 삭제\n
        """
        self.hwp.HAction.Run("MoveSelNextParaBegin")     # 다음 문단 (Ctrl + Shift + 아래방향키) 선택
        self.hwp.HAction.Run("Delete")
        return

    @_clearReadState
    def deleteWord(self) -> None:
        """
        현재 단어(공백 전까지를) 삭제\n
        """
        self.hwp.HAction.Run("MoveSelNextWord")    # 다음 단어 (Ctrl + Shift + 오른쪽방향키) 선택
        self.hwp.HAction.Run("Delete")
        return
    
    def insertFile(self, path: str) -> int:
        """
        현재 한글 문서의 맨 마지막에 다른 한글 문서를 끼워넣을 경우\n

        :param path: 끼워넣을 문서 경로
        """

        act = self.hwp.CreateAction("InsertFile")    # 한글 파일 끼워넣기
        pset = act.CreateSet()
        act.GetDefault(pset)  # 파리미터 초기화
        pset.SetItem("FileName", path)  # 파일 불러오기
        pset.SetItem("KeepSection", 1)  # 끼워 넣을 문서를 구역으로 나누어 쪽 모양을 유지할지 여부 on / off
        pset.SetItem("KeepCharshape", 1)     # 끼워 넣을 문서의 글자 모양을 유지할지 여부 on / off
        pset.SetItem("KeepParashape", 1)     # 끼워 넣을 문서의 문단 모양을 유지할지 여부 on / off
        pset.SetItem("KeepStyle", 0)    # 끼워 넣을 문서의 스타일을 유지할지 여부 on / off
        act.Execute(pset)

        return 0

    def _findLastPos(self) -> tuple:
        """
        문서의 마지막 위치 list para pos 반환
        :return: (list, para, pos)
        """
        nowpos = self._getPos()  # tuple
        self.MoveTopLevelEnd()# 맨 아래 위치 기록하고 돌아옴
        last = self._getPos()
        self._setPos(nowpos)
        return last


    def _findFirstPos(self) -> tuple:
        """
        문서의 처음 위치 list para pos 반환
        :return: (list, para, pos)
        """
        nowpos = self._getPos()  # tuple
        self.MoveTopLevelBegin()  # 맨 위 위치 기록하고 돌아옴
        first = self._getPos()
        self._setPos(nowpos)
        return first

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

        >>> hwp.KeyIndicator()[3]
        # 현재 위치 페이지 반환
        """
        return self.hwp.KeyIndicator()

    def linetoLFCR(self) -> None:
        """
        # 더이상 사용하지 않음

        명시적으로 표시한 개행문자(\\r\\n)개행하기
        >>> self.allreplace("\\\\r\\\\n", "\\r")
        # 이를 이용하기
        """

        self.hwp.HAction.Run("MoveTopLevelBegin")  # Ctrl + PGUP(맨 위 페이지로 이동)
        flag = 0

        while flag:
            flag = self.find("\\r\\n", 1)

            # 삭제
            self.hwp.HAction.Run("BreakPara")
            self.hwp.HAction.Run("MovePrevParaBegin")
        with warnings.catch_warnings():
            warnings.simplefilter('always')
            warnings.warn("Use allreplace(\"\\r\\n\", \"\\r\")")
        return None

    def multiColumn(self, num: int) -> None:
        """
        단을 n개로 분할

        선 양식 등을 조절 가능
        :param num: 분할할 단 개수
        >>> hwp.multiColumn(2)
        """
        self.hwp.HAction.GetDefault("MultiColumn", self.hwp.HParameterSet.HColDef.HSet)
        self.hwp.HParameterSet.HColDef.Count = num
        self.hwp.HParameterSet.HColDef.SameGap = self.hwp.MiliToHwpUnit(8.0)
        self.hwp.HParameterSet.HColDef.LineType = self.hwp.HwpLineType("Solid")
        self.hwp.HParameterSet.HColDef.LineWidth = self.hwp.HwpLineWidth("0.4mm")
        self.hwp.HParameterSet.HColDef.HSet.SetItem("ApplyClass", 832)
        self.hwp.HParameterSet.HColDef.HSet.SetItem("ApplyTo", 6)

        self.hwp.HAction.Execute("MultiColumn", self.hwp.HParameterSet.HColDef.HSet)
        return
    
    def _getPos(self) -> tuple:
        """
        현재 위치를 반환
        >>> a, b, c = hwp._getPos()
        """
        return self.hwp.GetPos()


    @_clearReadState
    def _setPos(self, curpos: tuple) -> None:
        """
        (list, para, pos)의 위치로 이동
        :param curpos: GetPos로 가져온 (list,para,pos)
        >>> hwp._setPos(curPos)
        """
        a, b, c = curpos
        self.hwp.SetPos(a, b, c)
        return

    def _deleteCtrl(self) -> None:
        """
        누름틀 제거용
        >>> hwp._deleteCtrl()
        """
        self.hwp.HAction.GetDefault("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
        self.hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
        self.hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 17)
        self.hwp.HAction.Execute("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)

    def _isbold(self) -> int:
        """
        포인터 위치가 굵은 글씨인지 판단
        
        >>> if hwp._isbold():
                pass

        :return: 굵은 글씨면 1 
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        # 켜짐 -> 1, 꺼짐 -> 0
        if Set.Item("Bold") == 1:   # 굵은 글씨 옵션이 켜져 있는지 여부
            return 1
        return 0

    def _editMode(self) -> int:
        """
        문서의 현재 편집 모드를 반환
        0 : 읽기 전용
        1 : 일반 편집모드
        2 : 양식 모드(양식 사용자 모드) : Cell과 누름틀 중 양식 모드에서 편집 가능 속성을 가진 것만 편집 가능하다.
        16 : 배포용 문서 (SetEditMode로 지정 불가능)
        self.hwp.EditMode = 1로 편집모드 강제 전환 가능
        SetEditMode(0)이라는 기능이 있는 것 같은데 미구현 ?

        :return: 현재 편집 모드
        """
        return self.hwp.EditMode

    def isPageOverbyEndnote(self) -> tuple:
        """
        # 수정 필요
        미주 위치를 근거로 확인
        페이지가 넘어갔는가?

        :return: 페이지가 다르면 1, 같으면 0, 한바퀴 돌았으면 -1과 함께 시작 위치 반환
        """
        # findNumber의 버그?
        # 찾은 번호의 위치를 반환하지 않음
        srcPos = self.findNumber()
        startpage = self.keyIndicator()[3]  # 시작 위치의 페이지

        self.MoveRight()
        dstPos = self.findNumber()
        lastpage = self.keyIndicator()[3]   # 마지막 위치의 페이지


        if dstPos < srcPos: # 한 바퀴 돌았으면 None
            return -1, srcPos

        if startpage != lastpage:   # 시작 페이지와 마지막 페이지가 다르다면 True
            return 1, srcPos

        return 0, srcPos
     
    def isPageOver(self, startpos: tuple, lastpos: tuple = None) -> int:
        """
        startpos와 lastpos의 페이지 위치를 확인
        페이지가 넘어갔는가?

        :param startpos: 시작 (list, para, pos)
        :param lastpos: 끝 (list, para, pos) 지정하지 않으면 현재 위치
        :return: 페이지가 다르면 1
        """
        nowpos = self.hwp.GetPos()  # 현재 위치

        if lastpos is None:     # 마지막 위치가 없으면
            lastpos = self.hwp.GetPos() # 현재 위치를 마지막 위치로
        
        self.hwp.MovePos(*startpos)
        startpage = self.keyIndicator()[3]  # 시작 위치의 페이지
        self.hwp.MovePos(*lastpos)
        lastpage = self.keyIndicator()[3]   # 마지막 위치의 페이지

        self.hwp.MovePos(*nowpos)   # 현재 위치로 다시 돌아옴

        if startpage != lastpage:   # 시작 페이지와 마지막 페이지가 다르다면 True
            return 1
        return 0

    def lineSpaceDecrease(self, line: int) -> None:
        """
        # (수정 필요)
        한칸 위 페이지 시작 ~ 현재 위치까지를 드래그하고 줄 간격 10% 줄임\n
        (2단 기준) 6줄을 한 페이지에 더 넣을 수 있음\n
        45, 45 -> 48, 48까지 가능
        """
        self.MovePageBegin()
        self.MoveSelPageDown()
        for _ in range(line):
            self.MoveSelNextParaBegin()
        self.ParagraphShapeDecreaseLineSpacing()
        self.hwp.HAction.Run("Cancel")
        self.BreakPara()
        self.ParagraphShapeIncreaseLineSpacing()

    @_clearReadState
    def BreakPara(self):
        """ Enter """
        self.hwp.HAction.Run("BreakPara")

    @_clearReadState
    def BreakPage(self):
        """ Ctrl + Enter """
        self.hwp.HAction.Run("BreakPage")

    @_clearReadState
    def BreakColumn(self):
        """ Ctrl + Shift + Enter """
        self.hwp.HAction.Run("BreakColumn")

    @_clearReadState
    def DeleteBack(self):
        """ BackSpace """
        self.hwp.HAction.Run("DeleteBack")

    @_clearReadState
    def MoveTopLevelBegin(self):
        """ Ctrl + PGUP(맨 위 페이지로 이동) """
        self.hwp.HAction.Run("MoveTopLevelBegin")

    @_clearReadState
    def MoveTopLevelEnd(self):
        """ Ctrl + PGDN(맨 아래 페이지로 이동) """
        self.hwp.HAction.Run("MoveTopLevelEnd")

    @_clearReadState
    def MoveRight(self):
        """ MoveRight(우방향키) """
        self.hwp.HAction.Run("MoveRight")

    @_clearReadState
    def MoveLeft(self):
        """ MoveLeft(좌방향키) """
        self.hwp.HAction.Run("MoveLeft")
    
    @_clearReadState
    def MoveUp(self):
        """ MoveUp(위방향키) """
        self.hwp.HAction.Run("MoveUP")

    @_clearReadState
    def MoveDown(self):
        """ MoveDown(아래방향키) """
        self.hwp.HAction.Run("MoveDown")

    @_clearReadState
    def MoveNextParaBegin(self):
        """ Ctrl + MoveDown """
        self.hwp.HAction.Run("MoveNextParaBegin")

    @_clearReadState
    def MovePrevParaBegin(self):
        """ Ctrl + MoveUp """
        self.hwp.HAction.Run("MovePrevParaBegin")

    @_clearReadState
    def MoveLineBegin(self):
        self.hwp.HAction.Run("MoveLineBegin")

    @_clearReadState
    def MoveSelNextParaBegin(self):
        """ Ctrl + Shift + MoveDown """
        self.hwp.HAction.Run("MoveSelNextParaBegin")

    @_clearReadState
    def MoveSelTopLevelEnd(self):
        self.hwp.HAction.Run("MoveSelTopLevelEnd")

    @_clearReadState
    def Delete(self):
        """ Delete """
        self.hwp.HAction.Run("Delete")

    @_clearReadState
    def CloseEx(self):
        """ Shift + ESC """
        self.hwp.HAction.Run("CloseEx")

    def CharShapeUnderline(self):
        """ # textStyle1로 통합 완료
        더이상 사용 X, 나중에 제거할 것
        >>> hwp.textStyle(underline=1)
         """
        self.hwp.HAction.Run("CharShapeUnderline")

    @_clearReadState
    def MoveSelNextWord(self):
        """ Ctrl + Shift + MoveRight """
        self.hwp.HAction.Run("MoveSelNextWord")
    
    @_clearReadState
    def MoveSelPrevWord(self):
        """ Ctrl + Shift + MoveLeft """
        self.hwp.HAction.Run("MoveSelPrevWord")

    @_clearReadState
    def MoveSelLeft(self):
        """ Shift + MoveLeft """
        self.hwp.HAction.Run("MoveSelLeft")

    @_clearReadState
    def MoveSelRight(self):
        """ Shift + MoveRight """
        self.hwp.HAction.Run("MoveSelRight")

    @_clearReadState
    def MoveLineEnd(self):
        self.hwp.HAction.Run("MoveLineEnd")

    def Undo(self):
        """ Ctrl + Z """
        self.hwp.Run("Undo")

    @_clearReadState
    def MovePageBegin(self):
        self.hwp.HAction.Run("MovePageBegin")

    @_clearReadState
    def MoveSelPageDown(self):
        self.hwp.HAction.Run("MoveSelPageDown")

    def ParagraphShapeDecreaseLineSpacing(self):
        """ 줄 간격 점점 줄임 (10%) -> 글자크기 9.5pt 기준 6줄을 줄일 수 있음 """
        self.hwp.HAction.Run("ParagraphShapeDecreaseLineSpacing")

    def ParagraphShapeIncreaseLineSpacing(self):
        """ 줄 간격 점점 늘림 (10%) -> 글자크기 9.5pt 기준 6줄을 늘릴 수 있음 """
        self.hwp.HAction.Run("ParagraphShapeIncreaseLineSpacing")
