from functools import wraps

# decorator의 원래 함수 반환값 주의할 것,,,
# super의 readstate를 호출할 수 있는가?
def clearReadState(func):
    """ 
    # ReleaseScan decorator
    강제로 ReleaseScan을 시행
    이렇게 구현하고 싶지 않았는데,,,
    어쩔 수 없는 듯 함
    
    한/글 캐럿이 움직이는 method에 붙여 놓을 것
    """
    @wraps(func)
    def inner_function(*args, **kwargs):
        self = args[0]  # class 함수의 첫 인자는 언제나 self
        self.hwp.ReleaseScan()
        self.hwpObject._readState = 0
        return func(*args, **kwargs)
    return inner_function