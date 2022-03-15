import win32com.client as win32 
import pandas as pd 

def editHWP(excelPath, hwpFile):

    excel = pd.read_excel(excelPath)

    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.Open(hwpFile)

    field_list = [i for i in hwp.GetFieldList().split("\x02")]

    # Select and Copy All Text
    hwp.Run('SelectAll')
    hwp.Run('Copy')

    # Move Cursor to End of Document
    hwp.MovePos(3)

    for i in range(len(excel) - 1):
        # Paste Previously Copied Text
        hwp.Run('Paste')

        # Move Cursor to End of Document
        hwp.MovePos(3)

    for page in range(len(excel)):
        for field in field_list:
            # Replace Fields with Data in Fields
            hwp.PutFieldText(f'{field}{{{{{page}}}}}', excel[field].iloc[page])

    hwp.Save()
    hwp.Quit()