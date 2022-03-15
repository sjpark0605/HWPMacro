from tkinter import Tk, Label, Frame, Button

class View:
    def __init__(self, dimension):
        self.root = Tk()
        self.root.title("엑셀 -> 한글 매크로")
        self.root.resizable(False, False)
        self.root.geometry(dimension)
        self.pathMap = {}      

    def generateField(self, labelContent, buttonText, action, row):
        fieldLabel = self.__generateFieldLabel(labelContent)
        pathFrame = self.__generateFrame()
        pathLabel = self.__generateFieldLabel(pathFrame)
        button = self.__generateButton(buttonText, action)

        self.__renderElements(fieldLabel, pathFrame, pathLabel, button, row)

    def __generateFieldLabel(self, labelContent):
        return Label(
            self.root,
            text = labelContent,
            padx = 5,
            pady = 3
        )

    def __generateFrame(self):
        return Frame(
            self.root,
            bg = "black",
            padx = 1,
            pady = 1
        )

    def __generateFieldLabel(self, frame):
        return Label(
            frame,
            bg = "white",
            width = 50
        )
    
    def __generateButton(self, text, action):
        return Button(
            self.root,
            text = text,
            command = action,
            padx = 10
        )

    def __renderElements(self, fieldLabel, pathFrame, pathLabel, button, row):
        fieldLabel.grid(row = row, column = 0)
        pathFrame.grid(row = row, column = 1)
        pathLabel.pack()
        button.grid(row = row, column = 2)