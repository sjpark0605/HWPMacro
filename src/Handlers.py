from tkinter import filedialog as fd
from Logic import editHWP

def onHWPButtonClick():
    filetypes = (
        ('한글 파일', '.hwp .hwpx'),
        ('모든 파일', '*.*')
    )

    filename = fd.askopenfilename(
        title = "한글 파일 불러오기",
        initialdir = '/',
        filetypes = filetypes
    )

    return filename

def onExcelButtonClick():
        filetypes = (
            ('엑셀 파일', '.xls .xlsx'),
            ('모든 파일', '*.*')
        )

        filename = fd.askopenfilename(
            title = "엑셀 파일 불러오기",
            initialdir = '/',
            filetypes = filetypes
        )

        return filename

def onMacroButtonClick(excelPath, hwpPath):
    editHWP(excelPath, hwpPath)