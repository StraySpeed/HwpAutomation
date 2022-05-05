import win32com.client as win32
import shutil
import base64
from PIL import Image
from io import BytesIO
import os

'''
# 자주쓰는 키 모음
hwp.HAction.Run("DeleteBack")   # BackSpace
hwp.HAction.Run("Delete")   # Delete
hwp.HAction.Run("BreakPara")    # Enter
hwp.HAction.Run("BreakPage")    # Ctrl + Enter
hwp.HAction.Run("BreakColumn")  # Ctrl + Shift + Enter
hwp.Run("Undo") # Ctrl + Z
hwp.HAction.Run("MoveTopLevelBegin")  # Ctrl + PGUP(맨 위 페이지로 이동)
hwp.HAction.Run("MoveTopLevelEnd")  # Ctrl + PGDN(맨 아래 페이지로 이동)
'''


def BasicForm():
    """
    기본 양식을 만드는 함수\n
    :return: gen_py_object(한글 객체 반환)
    """
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한글 객체 생성
    hwp.XHwpWindows.Item(0).Visible = True  # 한글 백그라운드 실행 -> False

    # 페이지 여백 설정
    hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
    hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(13.0)
    hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(13.0)
    hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(13.0)
    hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(13.0)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)

    hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)

    # 머리말 설정
    hwp.HAction.GetDefault("HeaderFooter", hwp.HParameterSet.HHeaderFooter.HSet)
    hwp.HParameterSet.HHeaderFooter.HSet.SetItem("HeaderFooterStyle", 0)
    hwp.HParameterSet.HHeaderFooter.HSet.SetItem("HeaderFooterCtrlType", 0)

    hwp.HAction.Execute("HeaderFooter", hwp.HParameterSet.HHeaderFooter.HSet)

    # 이미지 삽입
    # base64 형태로는 입력할 수 없는가?
    # 이미지 디코딩 -> 임시로 저장 -> 한글에 출력 -> 임시 저장 이미지 삭제
    # byte-like object
    img = None   # base64 형태의 이미지 입력
    if img is not None:
        dt = os.path.join(os.path.expanduser('~'), 'Desktop\\')     # 바탕화면 경로 지정
        img = Image.open(BytesIO(base64.b64decode(img)))    # 임시로 이미지 저장
        img.save(dt + 'temp.png', 'png')

        hwp.InsertPicture(dt + 'temp.png', True, 0, 0, 0, 0)
        os.remove(dt + 'temp.png')  # 이미지 삭제

    # 이미지 객체 속성 조절하기
    hwp.FindCtrl()  # 인접한 개체 선택 (양쪽에 존재하면 우측 우선)
    hwp.HAction.GetDefault("FormObjectPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HAction.GetDefault("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 액션 초기화
    hwp.HParameterSet.HShapeObject.TextWrap = hwp.TextWrapType("BehindText")  # 글 뒤로 배치
    hwp.HParameterSet.HShapeObject.TreatAsChar = 0  # 글자처럼 취급 해제
    hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)  # 실행
    hwp.HAction.Run("MoveLineEnd")

    # 머리말 입력
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = " [머리말]"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("MoveSelPrevParaBegin")
    hwp.HAction.Run("MoveSelPrevParaBegin")

    # 글씨체 조절
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameUser = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameSymbol = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameOther = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameJapanese = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameHanja = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameLatin = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameHangul = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(12.0)

    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HAction.Run("ParagraphShapeAlignRight")  # 오른쪽 정렬
    hwp.HAction.Run("CloseEx")  # 머리말 종료

    # 쪽 번호 지정
    hwp.HAction.GetDefault("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)
    hwp.HParameterSet.HPageNumPos.DrawPos = hwp.PageNumPosition("BottomCenter")
    hwp.HAction.Execute("PageNumPos", hwp.HParameterSet.HPageNumPos.HSet)

    # 쪽 테두리 지정
    hwp.HAction.GetDefault("PageBorder", hwp.HParameterSet.HSecDef.HSet)
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderTypeLeft = hwp.HwpLineType("DoubleSlim")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderTypeRight = hwp.HwpLineType("DoubleSlim")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderTypeTop = hwp.HwpLineType("DoubleSlim")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderWidthLeft = hwp.HwpLineWidth("0.5mm")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderWidthRight = hwp.HwpLineWidth("0.5mm")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderWidthTop = hwp.HwpLineWidth("0.5mm")
    hwp.HParameterSet.HSecDef.PageBorderFillBoth.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyToPageBorderFill", 3)

    hwp.HAction.Execute("PageBorder", hwp.HParameterSet.HSecDef.HSet)

    # 단 나누기
    hwp.HAction.GetDefault("MultiColumn", hwp.HParameterSet.HColDef.HSet)
    hwp.HParameterSet.HColDef.Count = 2
    hwp.HParameterSet.HColDef.SameGap = hwp.MiliToHwpUnit(8.0)
    hwp.HParameterSet.HColDef.LineType = hwp.HwpLineType("Solid")
    hwp.HParameterSet.HColDef.LineWidth = hwp.HwpLineWidth("0.4mm")
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyClass", 832)
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyTo", 6)

    hwp.HAction.Execute("MultiColumn", hwp.HParameterSet.HColDef.HSet)

    hwp.HAction.Run("MoveTopLevelBegin")

    # 글씨체 조절
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameUser = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeUser = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameSymbol = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeSymbol = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameOther = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeOther = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameJapanese = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeJapanese = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameHanja = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeHanja = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameLatin = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeLatin = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.FaceNameHangul = "나눔스퀘어라운드 Regular"
    hwp.HParameterSet.HCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(9.5)

    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HAction.Run("MoveTopLevelBegin")
    return hwp


# 자주 사용한 한글 기능들을 클래스로 저장
# 이 클래스를 사용하거나, 혹은 이를 참고하여 원하는대로 사용하기
class Pythonhwp():
    def __init__(self, path: str = None):
        """
        __init__\n
        한글 파일을 열고 한글 객체를 생성\n
        해당 위치에 한글 파일이 존재하지 않을 경우, 기본 양식을 생성 시도\n
        :param path: 오픈할 한글 파일 경로
        """
        # win32가 한글 파일을 열지 못할 경우 -> gen_py 라이브러리를 삭제 후 시도
        try:
            shutil.rmtree(os.path.join(os.path.expanduser('~'), r'AppData\Local\Temp\gen_py\\'))  # 삭제
        except FileNotFoundError:
            pass

        # 한글 파일을 열고 한글 객체를 반환
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한글 객체
        self.hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")  # 자동화 보안 모듈
        self.hwp.XHwpWindows.Item(0).Visible = True  # 한글 백그라운드 실행 -> False
        if os.path.isfile(path) and path.split('\\')[-1] == '.hwp':    # 파일이 존재할 경우
            self.hwp.Open(path, "HWP", "template:true, versionwarning:false")   # 파일 열기
            self.hwp.HAction.Run("MoveTopLevelBegin")  # 맨 위 페이지로 이동
        else:   # 파일이 존재하지 않을 경우
            self.hwp = BasicForm()

    def readtest(self):
        """
        한글 파일 출력 TEST용\n
        """
        self.hwp.InitScan()
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
        return 0

    def insertline(self, text: str):
        """
        한글에 텍스트를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행할 것\n
        :param text: 입력할 문자열
        """
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", text)
        act.Execute(set)
        self.hwp.HAction.Run("BreakPara")
        return 0

    def insertpicture(self, picturepath: str):
        """
        **수정필요**\n
        한글에 이미지를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행\n
        :param picturepath: 이미지 파일 경로
        """
        self.hwp.InsertPicture(picturepath, True, 0, 0, 0, 0)   # 원래 크기로, 반전 X, 워터마크 X, 실제 이미지 그대로
        return 0

    def readline(self, moveopt: int = 1):
        """
        한글에서 한 줄을 읽어오기\n
        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들인 text를 반환\n
        InitScan()으로 시작, ReleaseScan()으로 종료\n
        반드시 사용 후 ReleaseScan()을 줄 것\n
        InitScan 중 반복해서 GetText() 실행시 다음 줄을 읽어들임\n
        **그러나, 실제로 한글에서 포인터가 이동하는 것은 아님**\n
        읽어들인 위치로 이동하려면 hwp.MovePos(201)을 줄 것\n
        :param moveopt: 읽어들인 위치로 이동 시 1 (기본 0)
        :return: string(읽어들인 text)
        """
        self.hwp.InitScan()
        result = self.hwp.GetText()
        # 상태코드1 == 문서 끝에 도달하면 False 반환
        if result[0] == 1:
            return False
        text = result[1].strip()
        if moveopt:
            self.hwp.MovePos(201)   # 읽어들인 위치로 한글 포인터 이동
        self.hwp.ReleaseScan()
        return text

    def setnewnumber(self, num: int):
        """
        미주 번호를 바꾸는 함수\n
        한글 -> 새 번호로 시작 -> 미주 번호\n
        :param num: 입력 번호
        """
        self.hwp.HAction.GetDefault("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        self.hwp.HParameterSet.HAutoNum.NumType = self.hwp.AutoNumType("Endnote")
        self.hwp.HParameterSet.HAutoNum.NewNumber = int(num)
        self.hwp.HAction.Execute("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        return 0

    def findnumber(self):
        """
        미주로 이동하고 미주의 앞 위치를 반환\n
        한글 내 포인터 미주 앞으로 이동함\n
        문서의 맨 마지막 미주 너머에서 이를 실행 시 처음부터 재탐색하겠냐는 메시지 창이 뜸 -> 루프 시 탈출 조건 필요\n
        미주의 앞 위치를 반환하므로 루프 내에서 미주 탐색 시 동일 위치를 무한루프 할 수 있음에 주의\n
        hwp.HAction.Run("MoveNextParaBegin")으로 현재 문장을 넘어간 후 실행할 것\n
        또는 다른 방식으로 포인터 위치를 옮긴 후 실행할 것\n
        :return: tuple(list, para, pos) (현 위치)
        """
        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
        self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
        self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        return self.hwp.GetPos()  # 미주로 이동 후 현재 위치를 반환

    def insertendnote(self, text: str):
        """
        미주를 삽입하는 함수\n
        :param text: 삽입할 내용
        """
        self.hwp.HAction.Run("InsertEndnote")  # 미주 삽입
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", text)
        act.Execute(set)
        self.hwp.HAction.Run("CloseEx")     # 원래 위치로 돌아감
        return 0

    def save_file(self, path: str, savepdf: int = 0, quitopt: int = 0):
        """
        현재 한글 파일을 저장하고 종료함\n
        path -> 확장자(.hwp/.pdf)를 포함한 절대 경로여야 함\n
        파일을 열었을 때와 다른 경로로 입력시 원본과 다른 파일로 저장됨\n
        Format을 PDF로 바꾸고 확장자를 .pdf로 주면 PDF로도 저장 가능\n
        **pdf로 저장 시 결과물 한쪽 쏠림 현상 있음**\n
        한글 종료 시 파일에 변동이 있었다면 저장하겠냐는 창이 나오므로 주의\n
        :param path: 파일의 저장 경로
        :param savepdf: pdf일 경우 1, 아닐 경우 0 (기본 0)
        :param quitopt: 한글 창 닫기 시 1, 아닐 경우 0 (기본 0)
        """
        self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)  # 파일 저장 액션의 파라미터
        self.hwp.HParameterSet.HFileOpenSave.filename = path
        self.hwp.HParameterSet.HFileOpenSave.Attributes = 0
        if savepdf:
            self.hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        else:
            self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        if quitopt:
            self.hwp.Quit()
        return 0

    def change(self, findtext: str, changetext: str, regex: int = 0):
        """
        한글 -> 찾아 바꾸기를 실행\n
        기본양식의 경우, 머리말이 [머리말]로 되어 있음\n
        findtext에 이를 넣고 changetext에 바꿀 머리말을 넣어서 머리말 변경 가능\n
        모두 바꾸기 = AllReplace\n
        :param findtext: 찾을 text
        :param changetext: 바꿀 text
        :param regex: 정규표현식 사용 시 1 (기본 0)
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
        return 0

    def find(self, text: str, regex: int = 0):
        """
        한글 -> 찾기를 실행\n
        정규표현식을 이용하려면 1, 아니면 0을 대입\n
        한 바퀴를 돌면 hwp.SelectionMode == 0이 됨\n
        한 바퀴를 돌면 0, 아니면 1을 반환\n
        :param text: 찾을 문자열
        :param regex: 정규표현식 사용 시 1 (기본 0)
        :returns: tuple(list, para, pos) (찾은 위치), SelectionMode
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
        self.hwp.HAction.Run("MoveLeft")
        self.hwp.HAction.Run("MoveRight")
        return self.hwp.GetPos(), self.hwp.SelectionMode

    def textstyle(self, bold: int = 1, italic: int = 1):
        """
        글씨 스타일 굵게(bold), 기울임(italic)을 조정\n
        1 -> 적용, 0 -> 적용안함\n
        *밑줄, 취소선은 구현 필요*\n
        :param bold: 굵게 (기본 0)
        :param italic: 기울임 (기본 0)
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        if Set.Item("Bold") ^ bold:     # XOR 연산
            self.hwp.HAction.Run("CharShapeBold")
        if Set.Item("Italic") ^ italic:     # XOR 연산
            self.hwp.HAction.Run("CharShapeItalic")
        '''
        if underline == 1:
            self.hwp.HAction.Run("CharShapeUnderline")
        if strikeline == 1:
            self.hwp.HAction.Run("CharShapeStrikeout")
            
        # 사용을 추천하지는 않으나, 혹시 글자 스타일 또는 크기로 조건을 판단하는 경우의 예시
        Act = hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        # 켜짐 -> 1, 꺼짐 -> 0
        if Set.Item("Bold") == 1:   # 굵은 글씨 옵션이 켜져 있는지 여부
            pass
        if Set.Item("Italic") == 1:   # 기울임 옵션이 켜져 있는지 여부
            pass
        if Set.Item("Height") > 900:    # 글씨 크기가 9.00pt 초과하는 경우
            pass
        '''
        return 0

    def deleteline(self):
        """
        현재 줄(엔터 전까지를) 삭제\n
        **InitScan 도중에 시행하지 말 것\n
        만약 도중에 삭제해야 한다면 ReleaseScan으로 종료 후 다시 InitScan을 시행할 것\n
        """
        self.hwp.HAction.Run("MoveSelNextParaBegin")     # 다음 문단 (Ctrl + Shift + 아래방향키) 선택
        # self.hwp.HAction.Run("MoveSelNextWord")    # 다음 단어 (Ctrl + Shift + 오른쪽방향키) 선택
        self.hwp.HAction.Run("Delete")
        return 0
    
    def multicolumn(self, colnum: int = 2):
        """
        단을 나누는 함수\n
        :param colnum: 나눌 단의 개수 (기본 2)
        """
        self.hwp.HAction.GetDefault("MultiColumn", self.hwp.HParameterSet.HColDef.HSet)
        self.hwp.HParameterSet.HColDef.Count = colnum
        self.hwp.HParameterSet.HColDef.SameGap = self.hwp.MiliToHwpUnit(8.0)
        self.hwp.HParameterSet.HColDef.LineType = self.hwp.HwpLineType("Solid")
        self.hwp.HParameterSet.HColDef.LineWidth = self.hwp.HwpLineWidth("0.4mm")
        self.hwp.HParameterSet.HColDef.HSet.SetItem("ApplyClass", 832)
        self.hwp.HParameterSet.HColDef.HSet.SetItem("ApplyTo", 6)

        self.hwp.HAction.Execute("MultiColumn", self.hwp.HParameterSet.HColDef.HSet)
        return 0

    def insertfile(self, path: str, kepsec: int = 1, kepchrshp: int = 1, kepparashp: int = 1, kepstyle: int = 0):
        """
        현재 한글 문서의 맨 마지막에 다른 한글 문서를 끼워넣을 경우
        """
        self.hwp.HAction.Run("MoveTopLevelEnd")  # 가장 마지막 페이지로 이동
        self.hwp.HAction.Run("BreakPage")  # 쪽 나누기

        act = self.hwp.CreateAction("InsertFile")    # 한글 파일 끼워넣기
        pset = act.CreateSet()
        act.GetDefault(pset)  # 파리미터 초기화
        pset.SetItem("FileName", path)  # 파일 불러오기
        pset.SetItem("KeepSection", kepsec)  # 끼워 넣을 문서를 구역으로 나누어 쪽 모양을 유지할지 여부 on / off
        pset.SetItem("KeepCharshape", kepchrshp)     # 끼워 넣을 문서의 글자 모양을 유지할지 여부 on / off
        pset.SetItem("KeepParashape", kepparashp)     # 끼워 넣을 문서의 문단 모양을 유지할지 여부 on / off
        pset.SetItem("KeepStyle", kepstyle)    # 끼워 넣을 문서의 스타일을 유지할지 여부 on / off
        act.Execute(pset)
        return 0
    
    # Enter
    def BreakPara(self):
        self.hwp.HAction.Run("BreakPara")

    # Ctrl + Enter
    def BreakPage(self):
        self.hwp.HAction.Run("BreakPage")

    # Ctrl + Shift + Enter
    def BreakColumn(self):
        self.hwp.HAction.Run("BreakColumn")

    # BackSpace
    def DeleteBack(self):
        self.hwp.HAction.Run("DeleteBack")

    # Ctrl + PGUP(맨 위 페이지로 이동)
    def MoveTopLevelBegin(self):
        self.hwp.HAction.Run("MoveTopLevelBegin")

    # Ctrl + PGDN(맨 아래 페이지로 이동)
    def MoveTopLevelEnd(self):
        self.hwp.HAction.Run("MoveTopLevelEnd")

    # MoveRight
    def MoveRight(self):
        self.hwp.HAction.Run("MoveRight")

    # MoveLeft
    def MoveLeft(self):
        self.hwp.HAction.Run("MoveLeft")

    # MoveUp
    def MoveUp(self):
        self.hwp.HAction.Run("MoveUP")

    # MoveDown
    def MoveDown(self):
        self.hwp.HAction.Run("MoveDown")
