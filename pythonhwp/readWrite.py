from ._decorator import clearReadState
from ._hwpObject import hwpObject

class hwpReadWrite():
    """
    한/글 읽기 / 쓰기에 관련된 method
    """
    def __init__(self, hwpObject: hwpObject):
        self.hwp = hwpObject.hwp
        self.hwpObject = hwpObject

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
        """         한글에서 한 줄을 읽어오기\n

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임\n
        읽어들인 text를 반환\n
        텍스트를 읽어들이고 읽어들인 위치로 이동\n
        공백을 포함해서 읽어들임, 튜플 형태를 반환\n
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
        if not self.hwpObject._readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self.hwpObject._readState = 1

        texttuple = self.hwp.GetText()

        if texttuple[0] == 1 or texttuple[0] == 0:
            return (-1, None)
        
        # self.hwp.MovePos(201)   # 읽어들인 위치로 이동
        # hwpObject._readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self.hwpObject._readState and not statesave:
            self.hwp.ReleaseScan()
            self.hwpObject._readState = 0

        return texttuple


    def readLine(self, opt :int=None, ran :int =0x0077, statesave :int = 1) -> None | str:
        """
        한글에서 한 줄을 읽어오기\n

        포인터가 위치한 지점부터 엔터키 까지를 한 줄로 인식하여 읽어들임\n
        읽어들인 text를 반환\n
        텍스트를 읽어들이고 읽어들인 위치로 이동\n
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
        if not self.hwpObject._readState:
            if opt is None:
                self.hwp.InitScan(option=None, Range=ran)
            else:
                self.hwp.InitScan(option=opt, Range=ran)
            self.hwpObject._readState = 1

        text = (-1, '')
        while text[1].strip() == '':
            text = self.hwp.GetText()

            if text[0] == 1 or text[0] == 0:
                return None
        
        # self.hwp.MovePos(201)   # 읽어들인 위치로 이동
        # hwpObject._readState 중 statesave를 하지 않을 경우 ReleaseScan() 호출
        if self.hwpObject._readState and not statesave:
            self.hwp.ReleaseScan()
            self.hwpObject._readState = 0

        return text[1]
    

    @clearReadState
    def insertLine(self, text: str) -> None:
        """
        한글에 텍스트를 입력

        입력할 위치로 포인터를 옮긴 후 실행
        text -> 문자열 형식
        hwp.HAction.Run("BreakPara") 추가
        >>> hwp.insertLine("텍스트")
        """
        
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", hwpReadWrite.unicodetoAscii(str(text)))
        act.Execute(set)
        self.hwp.HAction.Run("BreakPara")

    @clearReadState
    def insertLinebyField(self, text: str) -> None:
        """
        한글에 텍스트를 입력(누름틀 버전)

        누름틀을 이용할 경우 포인터가 입력한 문자 끝으로 넘어가지 않음

        >>> hwp.insertLinebyField("텍스트")
        """
        self.hwp.CreateField(Direction="입력칸", memo="텍스트 입력", name="textarea")
        self.hwp.PutFieldText("textarea", hwpReadWrite.unicodetoAscii(str(text)))
        self.hwp.Run("DeleteField")
        # self.hwp.Run("MoveNextParaBegin")

    @clearReadState
    def insertPicture(self, picturepath: str, Embedded=True, sizeoption=0, Reverse=False, watermark=False, Effect=0, Width=0, Height=0):
        """
        한글에 이미지를 입력\n
        입력할 위치로 포인터를 옮긴 후 실행\n

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
        미주를 삽입하는 함수\n

        :param text: 삽입할 내용
        :return: 0
        """
        self.hwp.HAction.Run("InsertEndnote")  # 미주 삽입
        act = self.hwp.CreateAction("InsertText")
        set = act.CreateSet()
        set.SetItem("Text", hwpReadWrite.unicodetoAscii(str(text)))
        act.Execute(set)
        self.hwp.HAction.Run("CloseEx") # 원래 위치로 돌아감

    @clearReadState
    def deleteLine(self) -> None:
        """
        현재 줄(엔터 전까지를) 삭제\n
        """
        self.hwp.HAction.Run("MoveSelNextParaBegin")     # 다음 문단 (Ctrl + Shift + 아래방향키) 선택
        self.hwp.HAction.Run("Delete")
        return

    @clearReadState
    def deleteWord(self) -> None:
        """
        현재 단어(공백 전까지를) 삭제\n
        """
        self.hwp.HAction.Run("MoveSelNextWord")    # 다음 단어 (Ctrl + Shift + 오른쪽방향키) 선택
        self.hwp.HAction.Run("Delete")
        return
    
    @staticmethod
    def unicodetoAscii(text: str) -> str:
        """
        text를 ascii로 바꿔주는 함수(staticmethod)\n
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
