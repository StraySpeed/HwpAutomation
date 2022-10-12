import win32com.client as win32
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


# 자주 사용한 한글 기능들을 클래스로 저장
# 이 클래스를 사용하거나, 혹은 이를 참고하여 원하는대로 사용하기
class Pythonhwp():
    def __init__(self, path: str) -> None:
        """
        win32가 한글 파일을 열지 못할 경우 -> gen_py 라이브러리를 삭제 후 시도
        gencache.EnsureDispatch의 문제

        한글 파일을 열고 한글 객체를 반환
        
        gencache.EnsureDispatch / Dispatch의 차이
        gencache.EnsureDispatch는 Early Binding
        Dispatch는 Late Binding
        Early Binding은 먼저 로드를 전부 다 하는 방식
        gencache.EnsureDispatch가 Early Binding 한 것들을 나중에 못가져오는 문제 존재
        Dispatch를 사용 권장

        중요 - Dispatch 사용하면 InitScan 인자 오류 발생함
        한글을 읽어들이는 작업에는 gencache.EnsureDispatch
        한글에 출력하는 작업에는 Dispatch 사용 권장
        """

#        try:
#            shutil.rmtree(os.path.join(os.path.expanduser('~'), r'AppData\Local\Temp\gen_py\\'))  # 삭제
#        except FileNotFoundError:
#            pass

        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한글 객체
#        self.hwp = win32.Dispatch("HWPFrame.HwpObject")  # 한글 객체

        self.hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")  # 자동화 보안 모듈
        self.hwp.XHwpWindows.Item(0).Visible = True  # 한글 백그라운드 실행 -> False
        if os.path.isfile(path) and path.endswith(".hwp"):    # 파일이 존재할 경우
            self.hwp.Open(path, "HWP", "template:true, versionwarning:false")   # 파일 열기
            self.hwp.HAction.Run("MoveTopLevelBegin")  # 맨 위 페이지로 이동
        else:   # 파일이 존재하지 않을 경우
            raise FileNotFoundError

    def close(self) -> int:
        """
        문서 저장하지 않고 종료
        """
        self.hwp.XHwpDocuments.Item(0).Close(isDirty=False)
        self.hwp.Quit()

        return 0

    def _readtest(self) -> None:
        """
        한글 파일 출력 TEST용\n
        option, Range의 인자를 가짐\r
        생략하면 모든 컨트롤 대상, 문서 시작부터 끝까지 읽음\n
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

    def insertLine(self, text: str) -> int:
        """
        한글에 텍스트를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행\n
        text -> 문자열 형식\n
        엔터입력은 따로 해줘야 함\n
        hwp.HAction.Run("BreakPara") 또는 개행문자 입력\n
        """
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", str(text))
        act.Execute(set)
        return 0

    def insertPicture(self, picturepath: str) -> int:
        """
        한글에 이미지를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행\n
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
        

    def readLine(self, opt :int=None, ran :int =0x0033) -> str:
        """
        한글에서 한 줄을 읽어오기\n
        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임\n
        읽어들인 text 튜플을 반환\n
        InitScan()으로 시작, ReleaseScan()으로 종료\n
        반드시 사용 후 ReleaseScan()을 줄 것\n
        InitScan 중 반복해서 GetText() 실행시 다음 줄을 읽어들임\n
        그러나, 실제로 한글에서 포인터가 이동하는 것은 아님\n
        읽어들인 위치로 이동하려면 hwp.MovePos(201)을 줄 것\n
        """
        
        if opt is not None:
            self.hwp.InitScan(Range=ran)
        else:
            self.hwp.InitScan(option=opt, Range=ran)

        text = ''
        while text[1].strip() == '':
            text = self.hwp.GetText()

            if text[0] == 1 or text[0] == 0:
                return None
        
        self.hwp.MovePos(201)   # 읽어들인 위치로 이동
        self.hwp.HAction.Run("MoveNextParaBegin")
        self.hwp.ReleaseScan()
        return text

    def setNewNumber(self, num: int) -> None:
        """
        미주 번호를 바꾸는 함수\n
        한글 -> 새 번호로 시작 -> 미주 번호\n
        * 내부에서 num을 정수로 바꾸도록 되어 있음
        """
        self.hwp.HAction.GetDefault("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        self.hwp.HParameterSet.HAutoNum.NumType = self.hwp.AutoNumType("Endnote")
        self.hwp.HParameterSet.HAutoNum.NewNumber = int(num)
        self.hwp.HAction.Execute("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        return

    def findNumber(self) -> tuple:
        """
        미주로 이동하고 미주의 앞 위치를 반환\n
        한글 내 포인터 미주 앞으로 이동함\n
        문서의 맨 마지막 미주 너머에서 이를 실행 시 처음부터 재탐색하겠냐는 메시지 창이 뜸 -> 루프 시 탈출 조건 필요\n
        미주의 앞 위치를 반환하므로 루프 내에서 미주 탐색 시 동일 위치를 무한루프 할 수 있음에 주의\n
        hwp.HAction.Run("MoveNextParaBegin")으로 현재 문장을 넘어간 후 실행할 것\n
        또는 다른 방식으로 포인터 위치를 옮긴 후 실행할 것\n
        :return: (list, para, pos)
        """
        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
        self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
        self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        return self.hwp.GetPos()  # 미주로 이동 후 현재 위치를 반환

    def insertEndnote(self, text: str) -> None:
        """
        미주를 삽입하는 함수\n
        :param text: 삽입할 내용
        :return: None
        """
        self.hwp.HAction.Run("InsertEndnote")  # 미주 삽입
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", str(text))
        act.Execute(set)
        self.hwp.HAction.Run("CloseEx") # 원래 위치로 돌아감
        return

    def saveFile(self, path: str, pdfopt: int = 0, quitopt: int = 1) -> int:
        """
        현재 한글 파일을 저장하고 종료함\n
        path -> 확장자(.hwp)를 포함한 절대 경로여야 함\n
        파일을 열었을 때와 다른 경로로 입력시 원본과 다른 파일로 저장됨\n
        Format을 PDF로 바꾸고 확장자를 .pdf로 주면 PDF로도 저장 가능\n
        한글 종료 시 파일에 변동이 있었다면 저장하겠냐는 창이 나오므로 주의\n
        pdfopt 1 -> pdf로 출력(기본 0)
        quitopt 1 -> 한글 파일 종료(기본 1)
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

    def allReplace(self, findtext: str, changetext: str) -> None:
        """
        한글 -> 찾아 바꾸기를 실행\n
        바꾸기 = RepeatFind, 모두 바꾸기 = AllReplace\n
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
        self.hwp.HParameterSet.HFindReplace.FindRegExp = 0
        self.hwp.HParameterSet.HFindReplace.FindStyle = ""
        self.hwp.HParameterSet.HFindReplace.ReplaceStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindType = 1
        # 모두 바꾸기
        self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)

        return


    def find(self, text: str, regex: int = 0) -> int:
        """
        한글 -> 찾기를 실행\n
        정규표현식을 이용하려면 1, 아니면 0을 대입\n
        한 바퀴를 돌면 hwp.SelectionMode == 0이 됨
        한 바퀴를 돌면 0, 아니면 1을 반환
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

    def textStyle1(self, bold: int, italic: int) -> None:
        """
        글씨 스타일 굵게(bold), 기울임(italic), 밑줄(underline), 취소선(strikeline)을 조정\n
        1 -> 적용함, 0 -> 적용안함\n
        *밑줄, 취소선 수정 필요*\n
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        if Set.Item("Bold") ^ bold:     # XOR 연산
            self.hwp.HAction.Run("CharShapeBold")
        if Set.Item("Italic") ^ italic:     # XOR 연산
            self.hwp.HAction.Run("CharShapeItalic")

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


    def textStyle2(self, color: tuple, font: str = "바탕", size: int = 9.5) -> None:
        """
        현 위치 텍스트 font, size, color(RGB) 지정
        또는 선택 후 텍스트 스타일 지정
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
         

    def deleteLine(self) -> None:
        """
        현재 줄(엔터 전까지를) 삭제\n
        InitScan 도중에 시행하지 말 것\n
        만약 도중에 삭제해야 한다면 ReleaseScan으로 종료 후 다시 InitScan을 시행할 것\n
        """
        self.hwp.HAction.Run("MoveSelNextParaBegin")     # 다음 문단 (Ctrl + Shift + 아래방향키) 선택
        # self.hwp.HAction.Run("MoveSelNextWord")    # 다음 단어 (Ctrl + Shift + 오른쪽방향키) 선택
        self.hwp.HAction.Run("Delete")

        return

    def insertFile(self, path: str) -> int:
        """
        현재 한글 문서의 맨 마지막에 다른 한글 문서를 끼워넣을 경우\n
        """
        self.hwp.HAction.Run("MoveTopLevelEnd")  # 가장 마지막 페이지로 이동
        self.hwp.HAction.Run("BreakPage")  # 쪽 나누기

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
        nowpos = self.hwp.GetPos()  # tuple
        self.hwp.HAction.Run("MoveTopLevelEnd") # 맨 아래 위치 기록하고 돌아옴
        last = self.hwp.GetPos()
        self.hwp.MovePos(nowpos[0], nowpos[1], nowpos[2])
        return last

    def _findFirstPos(self) -> tuple:
        """
        문서의 처음 위치 list para pos 반환
        :return: (list, para, pos)
        """
        nowpos = self.hwp.GetPos()  # tuple
        self.hwp.HAction.Run("MoveTopLevelBegin") # 맨 아래 위치 기록하고 돌아옴
        last = self.hwp.GetPos()
        self.hwp.MovePos(nowpos[0], nowpos[1], nowpos[2])
        return last

    def keyIndicator(self) -> tuple:
        """
        현재 포인터 위치의 keyindicator를 반환
        튜플로 반환
        (Boolean, 총 구역, 현재 구역, 쪽, 단, 줄, 칸, 삽입/수정, 컨트롤 이름)
        페이지 가져오기 -> keyindicator[3]
        """
        return self.hwp.KeyIndicator()

    def linetoLFCR(self) -> None:
        """
        명시적으로 표시한 개행문자(\\r\\n)개행하기
        """
        self.hwp.HAction.Run("MoveTopLevelBegin")  # Ctrl + PGUP(맨 위 페이지로 이동)
        flag = 0

        while flag:
            flag = self.find("\\r\\n", 1)

            # 삭제
            self.hwp.HAction.Run("BreakPara")
            self.hwp.HAction.Run("MovePrevParaBegin")

        return

    def multiColumn(self, num: int) -> None:
        """
        단을 n개로 분할
        선 양식 등을 조절 가능
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
    
    def _setPos(self, curpos: tuple) -> None:
        """
        (list, para, pos)의 위치로 이동
        """
        self.hwp.SetPos(*curpos)
        return

    def _deleteCtrl(self) -> None:
        """
        누름틀 제거용
        """
        self.hwp.HAction.GetDefault("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
        self.hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
        self.hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 17)
        self.hwp.HAction.Execute("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)

    def _isbold(self) -> int:
        """
        포인터 위치가 굵은 글씨인지 판단        
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        # 켜짐 -> 1, 꺼짐 -> 0
        if Set.Item("Bold") == 1:   # 굵은 글씨 옵션이 켜져 있는지 여부
            return 1
        return 0