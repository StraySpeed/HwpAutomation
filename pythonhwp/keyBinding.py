from ._decorator import clearReadState
from ._hwpObject import hwpObject

class hwpKeyBinding():
    """
    한/글 키 바인딩
    """
    def __init__(self, hwpObject: hwpObject):
        self.hwp = hwpObject.hwp
        self.hwpObject = hwpObject

    @clearReadState
    def BreakPara(self):
        """ Enter """
        self.hwp.HAction.Run("BreakPara")

    @clearReadState
    def BreakPage(self):
        """ Ctrl + Enter """
        self.hwp.HAction.Run("BreakPage")

    @clearReadState
    def BreakColumn(self):
        """ Ctrl + Shift + Enter """
        self.hwp.HAction.Run("BreakColumn")

    @clearReadState
    def DeleteBack(self):
        """ BackSpace """
        self.hwp.HAction.Run("DeleteBack")

    @clearReadState
    def MoveTopLevelBegin(self):
        """ Ctrl + PGUP(맨 위 페이지로 이동) """
        self.hwp.HAction.Run("MoveTopLevelBegin")

    @clearReadState
    def MoveSelTopLevelEnd(self):
        """ Ctrl + Shift + PGDN(맨 아래 페이지로 이동 + 선택) """
        self.hwp.HAction.Run("MoveSelTopLevelEnd")

    @clearReadState
    def MoveTopLevelEnd(self):
        """ Ctrl + PGDN(맨 아래 페이지로 이동) """
        self.hwp.HAction.Run("MoveTopLevelEnd")

    @clearReadState
    def MoveRight(self):
        """ MoveRight(우방향키) """
        self.hwp.HAction.Run("MoveRight")

    @clearReadState
    def MoveLeft(self):
        """ MoveLeft(좌방향키) """
        self.hwp.HAction.Run("MoveLeft")
    
    @clearReadState
    def MoveUp(self):
        """ MoveUp(위방향키) """
        self.hwp.HAction.Run("MoveUP")

    @clearReadState
    def MoveDown(self):
        """ MoveDown(아래방향키) """
        self.hwp.HAction.Run("MoveDown")

    @clearReadState
    def MoveNextParaBegin(self):
        """ Ctrl + MoveDown """
        self.hwp.HAction.Run("MoveNextParaBegin")

    @clearReadState
    def MovePrevParaBegin(self):
        """ Ctrl + MoveUp """
        self.hwp.HAction.Run("MovePrevParaBegin")

    @clearReadState
    def MoveLineBegin(self):
        self.hwp.HAction.Run("MoveLineBegin")

    @clearReadState
    def MoveSelNextParaBegin(self):
        """ Ctrl + Shift + MoveDown """
        self.hwp.HAction.Run("MoveSelNextParaBegin")

    @clearReadState
    def MoveSelTopLevelEnd(self):
        self.hwp.HAction.Run("MoveSelTopLevelEnd")

    @clearReadState
    def Delete(self):
        """ Delete """
        self.hwp.HAction.Run("Delete")

    def Cancel(self):
        """ ESC """
        self.hwp.HAction.Run("Cancel")

    @clearReadState
    def CloseEx(self):
        """ Shift + ESC """
        self.hwp.HAction.Run("CloseEx")

    @clearReadState
    def MoveSelNextWord(self):
        """ Ctrl + Shift + MoveRight """
        self.hwp.HAction.Run("MoveSelNextWord")
    
    @clearReadState
    def MoveSelPrevWord(self):
        """ Ctrl + Shift + MoveLeft """
        self.hwp.HAction.Run("MoveSelPrevWord")

    @clearReadState
    def MoveSelLeft(self):
        """ Shift + MoveLeft """
        self.hwp.HAction.Run("MoveSelLeft")

    @clearReadState
    def MoveSelRight(self):
        """ Shift + MoveRight """
        self.hwp.HAction.Run("MoveSelRight")

    @clearReadState
    def MoveLineEnd(self):
        self.hwp.HAction.Run("MoveLineEnd")

    def Undo(self):
        """ Ctrl + Z """
        self.hwp.Run("Undo")

    @clearReadState
    def MovePageBegin(self):
        self.hwp.HAction.Run("MovePageBegin")

    @clearReadState
    def MoveSelPageDown(self):
        self.hwp.HAction.Run("MoveSelPageDown")
    
    @clearReadState
    def MoveParaBegin(self):
        self.hwp.HAction.Run("MoveParaBegin")

    @clearReadState
    def MoveParaEnd(self):
        self.hwp.HAction.Run("MoveParaEnd")

    def ParagraphShapeDecreaseLineSpacing(self):
        """ 줄 간격 점점 줄임 (10%) -> 글자크기 9.5pt 기준 6줄을 줄일 수 있음 """
        self.hwp.HAction.Run("ParagraphShapeDecreaseLineSpacing")

    def ParagraphShapeIncreaseLineSpacing(self):
        """ 줄 간격 점점 늘림 (10%) -> 글자크기 9.5pt 기준 6줄을 늘릴 수 있음 """
        self.hwp.HAction.Run("ParagraphShapeIncreaseLineSpacing")

