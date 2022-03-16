from View import View
from Handlers import *

GUI = View("600x110")


GUI.addField("엑셀 파일", "파일 선택", onExcelButtonClick, 0)
GUI.addField("한글 파일", "파일 선택", onHWPButtonClick, 1)
GUI.addRunButton("매크로 실행", onMacroButtonClick, 2)

GUI.run()