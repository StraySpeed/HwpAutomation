# 한/글 자동화
한컴오피스 한글 자동화를 위한 레포지토리입니다.

# Language
사용 언어
파이썬

# Library
사용하는 라이브러리
1. 필수
```python
import win32com.client as win32
import shutil
import os
````

2. 권장
필수는 아니나, 해당 라이브러리가 없을 시 제대로 동작하지 않을 수 있음
```python
from PIL import Image
from io import BytesIO
import base64
```

