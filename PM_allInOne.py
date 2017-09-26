import tkinter as tk
from tkinter import Tk, Menu
from tkinter.ttk import Frame, Label, Entry
from tkinter import messagebox
import tkinter.filedialog
import Reply_DL
import xmlConverter2 as xmlConverter
import os, sys
import matchingAlgorithm
import EN_Attachment




class Main_Window(Frame):

    def __init__(self, parent):

        Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()

    def initUI(self):

        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)

        fileMenu = Menu(menubar, tearoff=False)
        fileMenu.add_command(label="Open CSV's for EN", command=self.openCSVs)
        menubar.add_cascade(label="Files", menu=fileMenu)

        self.parent.title('PM_allInOne')
        self.pack()
        topFrame = Frame(self.parent)
        topFrame.pack(side='top', anchor='nw')
        bottomFrame = Frame(self.parent)
        bottomFrame.pack(side='top', anchor='nw')
        self.entrySaveBatch = Entry(topFrame)
        self.entrySaveBatch.pack(anchor='nw', fill='both')
        buttonSaveBatch = tk.Button(topFrame, 
                                    text="DL-Save", 
                                    width=8, 
                                    command=self.DLbatchSave)
        buttonSaveBatch.pack(side='left')
        buttonSendBatch = tk.Button(topFrame, 
                                    text="DL-Reply", 
                                    width=8, 
                                    command=self.DLbatchSend)
        buttonSendBatch.pack(side='left')
        buttonSaveKFD = tk.Button(bottomFrame, 
                                  text="Save KFD", 
                                  width=8, 
                                  command=self.convertXMLFiles)
        buttonSaveKFD.pack(side='left')
        buttonTest = tk.Button(bottomFrame, 
                                 text="Evening", 
                                 width=8, 
                                 command=self.evening_routines)
        buttonTest.pack(side='left')

    def openCSVs(self):

        Super_Duper = os.path.expanduser('~\Documents\SUPER DUPER\\')
        if not os.path.exists(Super_Duper):
            os.makedirs(Super_Duper)

        if getattr(sys, 'frozen', False):
            script_dir = os.path.dirname(sys.executable)
        elif __file__:
            script_dir = os.path.dirname(__file__)

        ftypes = [("Comma Seperated Values", "*.csv")]
        csvs = tk.filedialog.askopenfilenames(filetypes=ftypes)
        # fl = dlg.show()
        tempA = []
        if csvs != "":
            obj = EN_Attachment.exchange_notice(script_dir)
            for oneFile in csvs:
                tempA.append(obj.getData(oneFile))
            wb = obj.openWorkbook()
            n = 2
            for elem in tempA:
                obj.fillTemplate(wb, elem, n)
                n += 1
            obj.saveWorkbook(wb, Super_Duper)

    def evening_routines(self):

        file = r'\\SE10ORGFPS01\Clearing & Custody Services\1. Product Management\Listing Management\New strikes\Evening Strikes.xlsx'
        eveningO = matchingAlgorithm.matchingAlgorithm(file)

        if eveningO.generatedAmount != eveningO.matchedAmount:
            errorText = ('{0} matched from {1} generated'.
                         format(eveningO.matchedAmount, eveningO.generatedAmount))
            messagebox.showinfo('ERROR', errorText)

    def DLbatchSave(self):

        if self.entrySaveBatch.get() != '':
            recordID = self.entrySaveBatch.get()
            
            batch = Reply_DL.BatchingSTO_Message(recordID)
            batch.getAttachments()

    def DLbatchSend(self):

        # determine if application is a script file or frozen exe
        if getattr(sys, 'frozen', False):
            script_dir = os.path.dirname(sys.executable)
        elif __file__:
            script_dir = os.path.dirname(__file__)


        if self.entrySaveBatch.get() != '':
            recordID = self.entrySaveBatch.get()

            rel_path = '\Messages\Trading_Codes.txt'
            abs_file_path = script_dir + rel_path

            # print (abs_file_path)

            with open(abs_file_path, 'r') as f:
                text = f.read()
            
            batch = Reply_DL.BatchingSTO_Message(recordID)
            batch.getAttachments()
            batch.replyMessage(text, batch.getxlsxLocation())
            batch.createTask()

    def convertXMLFiles(self):

        xmlFilesSave = Reply_DL.BatchingDK_Message()
        for xmlFile in xmlFilesSave.xmlFileLocations:
            file = xmlConverter.NewCsvToOldCsv(xmlFile)
            file.readNewCsv()
            file.csvGenerator()


def main():
    root = Tk()
    ex = Main_Window(root)
    root.resizable(width=False, height=False)
    # root.resizable(0,0)
    # root.attributes("-toolwindow",1)
    root.mainloop()

if __name__ == "__main__":
    main()