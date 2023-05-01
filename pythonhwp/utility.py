from ._decorator import clearReadState
from ._hwpObject import hwpObject

class hwpUtility():
    """
    한/글의 여러 기능들(HAction)
    """
    def __init__(self, hwpObject: hwpObject):
        self.hwp = hwpObject.hwp
        self.hwpObject = hwpObject

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
    
    @clearReadState
    def findNumber(self) -> tuple:
        """
        미주로 이동하고 미주의 앞 위치를 반환\n

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

        return self.Pos  # 미주로 이동 후 현재 위치를 반환

    def allreplace(self, findtext: str, changetext: str, regex: int = 0) -> None:
        """
        한글 -> 찾아 바꾸기를 실행\n

        프라임에듀 기본 양식.hwp의 경우, 머리말이 [프라임에듀 머리말]로 되어 있음\n
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

    @clearReadState
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
        글씨 스타일 font, size, color(RGB) 지정
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
    
    def insertFile(self, path: str) -> None:
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
    
    def _deleteCtrl(self) -> None:
        """
        누름틀 제거용
        >>> hwp._deleteCtrl()
        """
        self.hwp.HAction.GetDefault("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
        self.hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
        self.hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 17)
        self.hwp.HAction.Execute("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
    

