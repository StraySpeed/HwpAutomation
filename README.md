# 한/글 자동화
한컴오피스 한글 자동화를 위한 레포지토리입니다.<br>

# Language
사용하는 언어<br>
파이썬<br>

# Library
사용하는 라이브러리<br>
1. 필수<br>
```python
import win32com.client as win32
````

2. 권장<br>
필수는 아니나, 해당 라이브러리가 없을 시 제대로 동작하지 않을 수 있음<br>
```python
from unidecode import unidecode
```
