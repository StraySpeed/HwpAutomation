from ._decorator import clearReadState
from .HwpKeyBinding import HwpKeyBinding

class HwpMethods(HwpKeyBinding):
    """
    한/글 method들을 정의
    """
    def __init__(self, path: str = None, gencache: bool = True):
        super().__init__(path, gencache)


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
    

    def readLine_tuple(self, opt :int=None, ran :int =0x0077, statesave :int = 1) -> tuple:
        """         ## 한글에서 한 줄을 읽어오기

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임

        읽어들인 text를 반환

        텍스트를 읽어들이고 읽어들인 위치로 이동하지 않음

        공백을 포함해서 읽어들임, 튜플 형태를 반환

        중간에 텍스트에 변동이 있으면 _clearReadState를 줄 것

        ## 참고 - 미주가 여러 줄인 경우 이를 읽어들이지 못함

        ## MovePos의 문제인 듯함

        :param opt: 읽어들이는 옵션(기본 None)
        :param ran: 범위(기본 문서 시작-끝까지)
        :param saveprev: InitScan 초기화 여부 (초기화하고자 하면 0)
        :return: 읽어들인 튜플(없으면 (-1, None))

        >>> hwp.readLine_tuple(opt=0, ran=0x0033) """

        
        # 문서의 맨 마지막이라면 None
        #if self._getPos() == self._findLastPos():
        #    return (-1, None)

        # readState가 아니라면 InitScan() 호출
        if not self._readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self._readState = 1

        texttuple = self.hwp.GetText()

        if texttuple[0] == 1 or texttuple[0] == 0:
            return (-1, None)
        
        # self.hwp.MovePos(201)   # 읽어들인 위치로 이동
        # _readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self._readState and not statesave:
            self.hwp.ReleaseScan()
            self._readState = 0

        return texttuple


    def readLine(self, opt :int=None, ran :int =0x0077, statesave :int = 1) -> None | str:
        """
        ## 한글에서 한 줄을 읽어오기

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임

        읽어들인 text를 반환

        텍스트를 읽어들이고 읽어들인 위치로 이동하지 않음

        공백없이 읽어들임

        ## 참고 - 미주가 여러 줄인 경우 이를 읽어들이지 못함

        ## MovePos의 문제인 듯함

        :param opt: 읽어들이는 옵션(기본 None)
        :param ran: 범위(기본 문서 시작-끝까지)
        :param statesave: 상태 저장 여부
        :return: 읽어들인 텍스트(없으면 None)

        >>> hwp.readLine(opt=0, ran=0x0033)
        """

        # readState가 아니라면 InitScan() 호출
        if not self._readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self._readState = 1

        text = (-1, '')
        while text[1].strip() == '':
            text = self.hwp.GetText()

            if text[0] == 1 or text[0] == 0:
                return None
        
        # self.hwp.MovePos(201)   # 읽어들인 위치로 이동
        # _readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self._readState and not statesave:
            self.hwp.ReleaseScan()
            self._readState = 0

        return text[1]
    

    @clearReadState
    def insertLine(self, text: str) -> None:
        """
        ## 한글에 텍스트를 입력

        입력할 위치로 포인터를 옮긴 후 실행

        hwp.HAction.Run("BreakPara") 추가
        >>> hwp.insertLine("텍스트")
        """
        
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", self.unicodetoAscii(str(text)))
        act.Execute(set)
        self.hwp.HAction.Run("BreakPara")

    @clearReadState
    def insertLinebyField(self, text: str) -> None:
        """
        ## 한글에 텍스트를 입력(누름틀 버전)

        누름틀을 이용할 경우 포인터가 입력한 문자 끝으로 넘어가지 않음

        >>> hwp.insertLinebyField("텍스트")
        """
        self.hwp.CreateField(Direction="입력칸", memo="텍스트 입력", name="textarea")
        self.hwp.PutFieldText("textarea", self.unicodetoAscii(str(text)))
        self.hwp.Run("DeleteField")
        # self.hwp.Run("MoveNextParaBegin")

    @clearReadState
    def insertPicture(self, picturepath: str, Embedded=True, sizeoption=0, Reverse=False, watermark=False, Effect=0, Width=0, Height=0):
        """
        ## 한글에 이미지를 입력

        입력할 위치로 포인터를 옮긴 후 실행

        (Path: 파일경로, 
        Embedded: 문서에포함여부,
        sizeoption=사이즈옵션[0: 이미지원래크기,
                            1: Width와 Height로 크기지정,
                            2: 셀 안에 있을 때 셀을 채움(그림비율 무시),
                            3: 셀에 맞추되 그림비율 유지(그림크기 변경)],
        Reverse=반전여부,
        watermark=워터마크여부,
        Effect=그림효과[0: 원래이미지,
                        1: 그레이스케일,
                        2: 흑백효과],
        Width=이미지너비mm,
        Height=이미지높이mm)

        :param picturepath: 이미지 파일의 전체 경로
        :return: 이미지 컨트롤 객체
        """
        ctrl = self.hwp.InsertPicture(Path=picturepath, Embedded=Embedded, sizeoption=sizeoption, Reverse=Reverse, watermark=watermark, Effect=Effect, Width=Width, Height=Height)   # 원래 크기로, 반전 X, 워터마크 X, 실제 이미지 그대로
        self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))  # 그림 앞으로 커서 이동
        return ctrl
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

    @clearReadState
    def insertEndnote(self, text: str) -> None:
        """
        ## 미주를 삽입하는 함수

        :param text: 삽입할 내용
        :return: 0
        """
        self.hwp.HAction.Run("InsertEndnote")  # 미주 삽입
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", self.unicodetoAscii(str(text)))
        act.Execute(set)
        self.hwp.HAction.Run("CloseEx") # 원래 위치로 돌아감

    @clearReadState
    def deleteLine(self) -> None:
        """
        ### 현재 줄(엔터 전까지를) 삭제
        """
        self.hwp.HAction.Run("MoveSelNextParaBegin")     # 다음 문단 (Ctrl + Shift + 아래방향키) 선택
        self.hwp.HAction.Run("Delete")
        return

    @clearReadState
    def deleteWord(self) -> None:
        """
        ### 현재 단어(공백 전까지를) 삭제
        """
        self.hwp.HAction.Run("MoveSelNextWord")    # 다음 단어 (Ctrl + Shift + 오른쪽방향키) 선택
        self.hwp.HAction.Run("Delete")
        return
    

    def setNewNumber(self, num: int) -> None:
        """
        ## 미주 번호를 바꾸는 함수

        한글 -> 새 번호로 시작 -> 미주 번호

        * 참고 - 내부에서 num을 정수로 바꾸도록 되어 있음

        :param num: 바꿀 번호
        """
        self.hwp.HAction.GetDefault("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        self.hwp.HParameterSet.HAutoNum.NumType = self.hwp.AutoNumType("Endnote")
        self.hwp.HParameterSet.HAutoNum.NewNumber = int(num)
        self.hwp.HAction.Execute("NewNumber", self.hwp.HParameterSet.HAutoNum.HSet)
        return
    
    @clearReadState
    def findNumber(self) -> tuple:
        """
        ## 미주로 이동하고 미주의 앞 위치를 반환

        한글 내 포인터 미주 앞으로 이동함

        ---
        ## 수정
        이 MessageBox는 제어되지 않는 박스임

        CtrlID를 체크하는 방식으로 해야 할 듯함

        ```
        ctrl = hwp.HeadCtrl
        hwp.MoveTopLevelBegin()
        while ctrl: # 모든 컨트롤 확인
            if ctrl.CtrlID == "tbl":  # 표 컨트롤이 있다면
                pass
            ctrl = ctrl.Next    # 다음 컨트롤로
        ```

        ---
        문서의 맨 마지막 미주 너머에서 이를 실행 시 처음부터 재탐색하겠냐는 메시지 창이 뜸 -> 루프 시 탈출 조건 필요

        -> hwp.SetMessageBoxMode(0x10000)

        미주의 앞 위치를 반환하므로 루프 내에서 미주 탐색 시 동일 위치를 무한루프 할 수 있음에 주의

        hwp.HAction.Run("MoveNextParaBegin")으로 현재 문장을 넘어간 후 실행할 것

        또는 다른 방식으로 포인터 위치를 옮긴 후 실행할 것
        :return: (list, para, pos)
        """
        # self.hwp.SetMessageBoxMode(0x10000)  # 예/아니오 창에서 "예"를 누르는 method
        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
        self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
        self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
        self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)

        return self.pos  # 미주로 이동 후 현재 위치를 반환

    def allreplace(self, findtext: str, changetext: str, regex: int = 0) -> None:
        """
        ## 한글 -> 찾아 바꾸기를 실행

        프라임에듀 기본 양식.hwp의 경우, 머리말이 [프라임에듀 머리말]로 되어 있음

        findtext에 이를 넣고 changetext에 바꿀 머리말을 넣어서 머리말 변경 가능

        바꾸기 = RepeatFind, 모두 바꾸기 = AllReplace
        
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

    @clearReadState
    def find(self, text: str, regex: int = 0) -> int:
        """
        ## 한글 -> 찾기를 실행

        정규표현식을 이용하려면 1, 아니면 0을 대입

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
        ## 글씨 스타일 굵게(bold), 기울임(italic), 밑줄(underline), 취소선(strikeline)을 조정

        1 -> 적용함, 0 -> 적용안함

        *기본은 적용하지 않음*
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
        # Act.execute(Set)
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

    def textStyle2(self, color: tuple = (0, 0, 0), font: str = r"바탕", size: float = 9.5) -> None:
        """
        ## 글씨 스타일 font, size, color(RGB) 지정

        또는 선택 후 텍스트 스타일 지정

        :param color: 색 지정(RGB)
        :param font: 글씨체(기본 "바탕")
        :param size: 폰트 크기(기본 9.5)
        >>> hwp.textStyle2(color=(255, 0, 0), font=r"나눔스퀘어라운드 Regular", size=9.5)
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        Set.SetItem("FaceNameUser", font)
        Set.SetItem("FaceNameSymbol", font)
        Set.SetItem("FaceNameOther", font)
        Set.SetItem("FaceNameJapanese", font)
        Set.SetItem("FaceNameHanja", font)
        Set.SetItem("FaceNameLatin", font)
        Set.SetItem("FaceNameHangul", font)
        
        Set.SetItem("Height", self.hwp.PointToHwpUnit(size))
        Set.SetItem("TextColor", self.hwp.RGBColor(*color))
        Act.Execute(Set)
        return

    def multiColumn(self, num: int) -> None:
        """
        ## 단을 n개로 분할

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
    
    def _deleteCtrl(self) -> None:
        """
        ### 누름틀 제거용
        >>> hwp._deleteCtrl()
        """
        self.hwp.HAction.GetDefault("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
        self.hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
        self.hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 17)
        self.hwp.HAction.Execute("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)

    def tableToString(self) -> None:
        """
        ## 표를 문자열로 전환
        
        표 안에 커서가 있어야 함
        
        Ctrl 선택하면 될 듯?함

        * 공백 3칸으로 분리
        """
        self.hwp.HAction.GetDefault("TableTableToString", self.hwp.HParameterSet.HTableTblToStr.HSet)
        self.hwp.HParameterSet.HTableTblToStr.HSet.SetItem("DelimiterType", 3)  # UserDefine으로 사용하도록 설정(0 Tab/ 1 쉼표/ 2 공백/ 3 사용자설정)
        self.hwp.HParameterSet.HTableTblToStr.UserDefine = "   "
        self.hwp.HAction.Execute("TableTableToString", self.hwp.HParameterSet.HTableTblToStr.HSet)

    @staticmethod
    def unicodetoAscii(text: str) -> str:
        """
        ## text를 ascii로 바꿔주는 함수(staticmethod)
        
        preserve 값 수정을 통해 보존할 유니코드 지정 가능(현재는 U+fff0이 들어가 있음)

        :param text: 바꿀 문자열
        :return: 아스키 값으로 바뀐 문자열
        """
        from unidecode import unidecode
        import re

        ptext = text
        preserve = ["￰"]    # 보존할 유니코드 값
        for p in preserve:  # 보존할 값을 미리 치환함
            ptext = ptext.replace(p, str(p.encode("utf-8")))

        textlist = re.compile(r'[^ㄱ-ㅣ가-힣①②③④⑤ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝ■]+').findall(text) # 한글/보기문자가 아닌 것들과 매칭
        for letter in textlist:   
            ptext = ptext.replace(letter, unidecode(letter))    # 유니코드를 아스키로 통일함

        for p in preserve:  # 보존한 값을 다시 되돌림
            ptext = ptext.replace(str(p.encode("utf-8")), p)
        return ptext
