import win32com.client as win32 
import pandas as pd 

excel = pd.read_excel('[EXCEL INPUT PATH]')
print(excel)

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.Open("[HWP INPUT PATH]")

field_list = [i for i in hwp.GetFieldList().split("\x02")]
print(field_list)

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