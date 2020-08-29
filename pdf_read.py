# ocr_card.py
import os
from PIL import Image
import pyocr
import pyocr.builders
import pandas as pd
import openpyxl
import re

# 1.インストール済みのTesseractのパスを通す
path_tesseract = "C:/Program Files/Tesseract-OCR"
if path_tesseract not in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + path_tesseract

# 2.OCRエンジンの取得
tools = pyocr.get_available_tools()
tool = tools[0]

# 3.原稿画像の読み込み
img_org = Image.open("C:/Users/VARNG/Desktop/scrapping/20200829_finance_hack/mission.jpg")

# 4.ＯＣＲ実行
builder = pyocr.builders.TextBuilder()
result = tool.image_to_string(img_org, lang="kor", builder=builder)

result = str(result)

#
m = re.match(r'지 급 명 령', result)
if m is not None:
    bool = 1
else:
    bool = 0

output_line_list = [result, bool]
df = pd.DataFrame(output_line_list)
df.to_excel('./test.xlsx')
