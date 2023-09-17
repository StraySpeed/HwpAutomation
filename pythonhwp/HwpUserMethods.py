from .HwpMethods import HwpMethods
from ._decorator import clearReadState

class HwpUserMethods(HwpMethods):
    def __init__(self, path: str = None, gencache: bool = True):
        super().__init__(path, gencache)

    def _findLastPos(self) -> tuple:
        """
        ### 문서의 마지막 위치 list para pos 반환

        :return: (list, para, pos)
        """
        nowpos = self.pos  # tuple
        self.MoveTopLevelEnd()# 맨 아래 위치 기록하고 돌아옴
        last = self.pos
        self.pos = nowpos
        return last


    def _findFirstPos(self) -> tuple:
        """
        ### 문서의 처음 위치 list para pos 반환

        :return: (list, para, pos)
        """
        nowpos = self.pos  # tuple
        self.MoveTopLevelBegin()  # 맨 위 위치 기록하고 돌아옴
        first = self.pos
        self.pos = nowpos
        return first
    
    def _isbold(self) -> bool:
        """
        ## 포인터 위치가 굵은 글씨인지 판단
        
        >>> if hwp._isbold():
                pass

        :return: 굵은 글씨면 1 
        """
        Act = self.hwp.CreateAction("CharShape")  # 액션테이블에서 "글자 모양" 검색, 액션아이디에서 "CharShape" 찾음
        Set = Act.CreateSet()  # 세트 생성
        Act.GetDefault(Set)  # 세트 초기화(Set의 파라미터에 현재 문서의 값을 적용)
        # 켜짐 -> 1, 꺼짐 -> 0
        if Set.Item("Bold") == 1:   # 굵은 글씨 옵션이 켜져 있는지 여부
            return True
        return False
    
    def isPageOver(self, startpos: tuple, lastpos: tuple = None) -> int:
        """
        ## startpos와 lastpos의 페이지 위치를 확인
        
        페이지가 넘어갔는가?

        :param startpos: 시작 (list, para, pos)
        :param lastpos: 끝 (list, para, pos) 지정하지 않으면 현재 위치
        :return: 페이지가 다르면 1, 줄간격 줄일거면 2, 넘어가지 않았으면 0
        """
        nowpos = self.pos  # 현재 위치

        if lastpos is None:     # 마지막 위치가 없으면
            lastpos = self.pos # 현재 위치를 마지막 위치로
        
        self.pos = startpos
        startpage = self.keyIndicator[3]  # 시작 위치의 페이지
        self.pos = lastpos
        lastpage = self.keyIndicator[3]   # 마지막 위치의 페이지
        lastline = self.keyIndicator[5]   # 마지막 위치의 줄

        self.pos = nowpos   # 현재 위치로 다시 돌아옴

        if startpage != lastpage:   # 시작 페이지와 마지막 페이지가 다르다면 True
            if lastline <= 6:
                return 2    # 줄 간격을 줄일 것
            return 1    # 문제를 다음 페이지로 넘길 것
        return 0    # 페이지가 넘어가지 않음
    
    def lineSpaceDecrease(self) -> None:
        """
        # (수정 필요)
        한칸 위 페이지 시작 ~ 현재 위치까지를 드래그하고 줄 간격 10% 줄임

        (2단 기준) 6줄을 한 페이지에 더 넣을 수 있음

        Alt + Page Up / Alt + Page Down

        45, 45 -> 48, 48까지 가능
        """
        self.MovePageBegin()
        self.MoveSelTopLevelEnd()
        self.ParagraphShapeDecreaseLineSpacing()
        self.Cancel()
        self.BreakColumn()
        self.ParagraphShapeIncreaseLineSpacing()
        return
    
    @clearReadState
    def __findUnderline(self) -> tuple:
        """
        ### 내부에서만 사용할 함수

        밑줄을 찾는 method
        """
        # 반복 찾기 방법
        self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)

        self.hwp.HParameterSet.HFindReplace.ReplaceString = ""
        self.hwp.HParameterSet.HFindReplace.FindString = ""
        self.hwp.HParameterSet.HFindReplace.IgnoreReplaceString = 0
        self.hwp.HParameterSet.HFindReplace.IgnoreFindString = 0
        self.hwp.HParameterSet.HFindReplace.Direction = self.hwp.FindDir("Forward")
        self.hwp.HParameterSet.HFindReplace.WholeWordOnly = 0
        self.hwp.HParameterSet.HFindReplace.UseWildCards = 0
        self.hwp.HParameterSet.HFindReplace.SeveralWords = 0
        self.hwp.HParameterSet.HFindReplace.AllWordForms = 0
        self.hwp.HParameterSet.HFindReplace.MatchCase = 0

        self.hwp.HParameterSet.HFindReplace.FindCharShape.UnderlineType = self.hwp.HwpUnderlineType("Bottom")
        self.hwp.HParameterSet.HFindReplace.FindCharShape.UnderlineColor = 0
        self.hwp.HParameterSet.HFindReplace.FindCharShape.UnderlineShape = self.hwp.HwpUnderlineShape("Solid")

        self.hwp.HParameterSet.HFindReplace.ReplaceMode = 0
        self.hwp.HParameterSet.HFindReplace.ReplaceStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindStyle = ""
        self.hwp.HParameterSet.HFindReplace.FindRegExp = 0   # 정규표현식으로 찾을 경우
        self.hwp.HParameterSet.HFindReplace.FindJaso = 0
        self.hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0
        self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        self.hwp.HParameterSet.HFindReplace.FindType = 1

        self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)

        return self.hwp.GetPos()

    def underLineUnicode(self, opt:int = 1) -> None:
        """
        ## 한/글에서 밑줄 양식을 치환하는 method

        "￰"(U+FFF0) 사용

        위치(self.Pos)를 기준으로 밑줄 연산을 고려했는데

        도형 안에 글자가 있으면 이게 꼬임

        위에서 아래로 순서대로 읽을 수는 있으나 도형은 (list, para, pos)에서 list가 3(일반문단 0)

        연산 구조를 바꿔야 함

        >>> if last < flag:
        의 연산이 문제

        :param opt: 옵션(1 = 밑줄->유니코드로 치환, 0 = 유니코드->밑줄 되돌리기)
        """
        if opt:
            self.MoveTopLevelBegin()
            prev = self._findFirstPos()
            last = self._findLastPos()
            first = None
            while True:
                flag = self.__findUnderline()

                if first is None:
                    first = flag
                else:
                    if flag[0] == first[0] and flag[1] == first[1] and flag[2] == first[2] + 2:
                        break
                
                # 해당 연산이 도형 안에 있는 케이스를 고려하지 못함
                # 일단 보류해놓을 것
                # if last < flag: # 미주에 있는 밑줄은 무시하기
                #    continue

                if prev == flag or (prev[0] == flag[0] and prev[1] == flag[1] and prev[2] + 2 == flag[2]):    # 밑줄이 없거나 한개이면
                    break

                prev = flag

                self.MoveLeft()
                self.MoveRight()
                self.insertLinebyField("￰")    # 밑줄의 앞쪽에 U+fff0을 입력

                self.__findUnderline()

                self.MoveRight()
                self.MoveLeft()
                self.insertLinebyField("￰")    # 밑줄의 뒤쪽에 U+fff0을 입력
                self.MoveRight()
                
            self.MoveTopLevelBegin()
        else:
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
            self.hwp.HParameterSet.HFindReplace.FindString = "￰[^￰]+￰"
            self.hwp.HParameterSet.HFindReplace.ReplaceString = ""

            # 밑줄 긋는 옵션
            self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineType = self.hwp.HwpUnderlineType("Bottom")
            self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineColor = 0
            self.hwp.HParameterSet.HFindReplace.ReplaceCharShape.UnderlineShape = self.hwp.HwpUnderlineShape("Solid")

            self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1
            self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            self.hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0
            self.hwp.HParameterSet.HFindReplace.FindJaso = 0
            # 정규 표현식으로 찾는 옵션 -> 1
            self.hwp.HParameterSet.HFindReplace.FindRegExp = 1
            self.hwp.HParameterSet.HFindReplace.FindStyle = ""
            self.hwp.HParameterSet.HFindReplace.ReplaceStyle = ""
            self.hwp.HParameterSet.HFindReplace.FindType = 1
            # 모두 바꾸기
            self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
            self.allreplace("￰", "", 0)
        return