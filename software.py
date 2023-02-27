import wx
import openpyxl
import xlrd

class HomePanel(wx.Panel):

    def __init__ (self, parent):
        super().__init__(parent)

        #STATIC TEXT
        h1 = wx.StaticText(self, label="SOFTWARE PER POPOLARE NUOVA TABELLA CONTEGGI")
        h2 = wx.StaticText(self, label="Istruzioni per l'uso:")
        p1 = wx.StaticText(self, label="1. Cliccare 'Genera nuovo file excel'")
        p2 = wx.StaticText(self, label="2. Apparirà una finestra di dialogo in cui selezionare il file di ingresso, contenente i dati da inserire nella tabella")
        p3 = wx.StaticText(self, label="3. Apparirà una seconda finestra di dialogo in cui selezionare il file modello 'TABELLA GENERICA CONTEGGI'")
        p4 = wx.StaticText(self, label="4. Ora non resta che attendere x.")
        p5 = wx.StaticText(self, label="IL NUOVO FILE VIENE SALVATO CON IL NOME 'NUOVA TABELLA CONTEGGI' nella cartella che contiene il programma.")

        font = h1.GetFont()
        font.PointSize += 10
        font = font.Bold()
        h1.SetFont(font)


        #BTN INPUT
        browseBtnIn = wx.Button(self, label="File Input")
        browseBtnIn.Bind(wx.EVT_BUTTON, self.on_browse_in)
        browseBtnIn.Disable()
        browseBtnIn.Hide()

        #BTN OUTPUT
        browseBtnOut = wx.Button(self, label="File Output")
        browseBtnOut.Bind(wx.EVT_BUTTON, self.on_browse_out)
        browseBtnOut.Disable()
        browseBtnOut.Hide()

        #BTN RUN
        btnRun = wx.Button(self, label="Genera nuovo file excel")
        btnRun.Bind(wx.EVT_BUTTON, self.on_click_run)

        """ Grafica """
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(h1, 0, wx.CENTER | wx.ALL, 5)
        sizer.Add(h2, 0, wx.CENTER | wx.ALL, 2)
        sizer.Add(p1, 0, wx.LEFT, 2)
        sizer.Add(p2, 0, wx.LEFT, 2)
        sizer.Add(p3, 0, wx.LEFT, 2)
        sizer.Add(p4, 0, wx.LEFT, 2)
        sizer.Add(p5, 0, wx.CENTER | wx.ALL, 10)
        sizer.Add(btnRun, 0, wx.CENTER | wx.ALL, 25)
        self.SetSizer(sizer)


        


    def on_browse_in(self, event):
        """
        Browse an Excel file for input
        @param event: The event object         
        """

        fileIn = wx.FileSelector("Scegliere il file excel (.xls) d'ingresso, contentente i dati da inserire")
        return fileIn



    
    def on_browse_out(self, event):
        """
        Browse an Excel file for output
        @param event: The event object
        """

        fileOut = wx.FileSelector("Scegliere il file excel (.xlsx) da usare come modello")
        return fileOut




    def on_click_run(self, event):
        """
        Manipulate the value inside the input excel file, then populate the output excel file
        @param evet: The event object
        """

        input = self.on_browse_in(event)
        output = self.on_browse_out(event)

        """ Processing input file """
        wb = xlrd.open_workbook(input)
        ws = wb.sheet_by_index(0)

        startDays = []
        endDays = []

        for x in range (4, ws.nrows, 2):
            tmp = ws.cell_value(rowx=x, colx=0)
            day = tmp.replace(". ", "/")

            if day.__contains__("Gen"):
                day = day.replace("Gen", "01")
            elif day.__contains__("Feb"):
                day = day.replace("Feb", "02")
            elif day.__contains__("Mar"):
                day = day.replace("Mar", "03")
            elif day.__contains__("Apr"):
                day = day.replace("Apr", "04")
            elif day.__contains__("Mag"):
                day = day.replace("Mag", "05")
            elif day.__contains__("Giu"):
                day = day.replace("Giu", "06")
            elif day.__contains__("Lug"):
                day = day.replace("Lug", "07")
            elif day.__contains__("Ago"):
                day = day.replace("Ago", "08")
            elif day.__contains__("Set"):
                day = day.replace("Set", "09")
            elif day.__contains__("Ott"):
                day = day.replace("Ott", "10")
            elif day.__contains__("Nov"):
                day = day.replace("Nov", "11")
            elif day.__contains__("Dic"):
                day = day.replace("Dic", "12")

            startDays.append(day)

            tmp2 = ws.cell_value(rowx=x, colx=1)
            day2 = tmp2.replace(". ", "/")

            if day2.__contains__("Gen"):
                day2 = day2.replace("Gen", "01")
            elif day2.__contains__("Feb"):
                day2 = day2.replace("Feb", "02")
            elif day2.__contains__("Mar"):
                day2 = day2.replace("Mar", "03")
            elif day2.__contains__("Apr"):
                day2 = day2.replace("Apr", "04")
            elif day2.__contains__("Mag"):
                day2 = day2.replace("Mag", "05")
            elif day2.__contains__("Giu"):
                day2 = day2.replace("Giu", "06")
            elif day2.__contains__("Lug"):
                day2 = day2.replace("Lug", "07")
            elif day2.__contains__("Ago"):
                day2 = day2.replace("Ago", "08")
            elif day2.__contains__("Set"):
                day2 = day2.replace("Set", "09")
            elif day2.__contains__("Ott"):
                day2 = day2.replace("Ott", "10")
            elif day2.__contains__("Nov"):
                day2 = day2.replace("Nov", "11")
            elif day2.__contains__("Dic"):
                day2 = day2.replace("Dic", "12")
            endDays.append(day2)



        """ Processing output file """
        wbOut = openpyxl.load_workbook(output)
        wsOut = wbOut.active

        dateStart = []
        dateStop = []

        dateStaSplit = []
        dateStoSplit = []

        dateStaDef = []
        dateStoDef = []

        hourStart = []
        hourStop = []

        for x in startDays:
            dateStart.append(x.split(' '))

        for y in endDays:
            dateStop.append(y.split(' '))

        for i in dateStart:
            dateStaSplit.append(i[0].split("/"))
            hourStart.append(i[1])
        
        for index in dateStaSplit:
            dateStaDef.append(index[2] + "-" + index[1] + "-" + index[0])            
        
        # print (dateStaDef)
            
        for j in dateStop:
            dateStoSplit.append(j[0].split("/"))
            hourStop.append(j[1])
        
        for jndex in dateStoSplit:
            dateStoDef.append(jndex[2] + "-" + jndex[1] + "-" + jndex[0])

        k = 0
        #for day in dateStaDef:
        for row in wsOut.iter_rows(min_col=1, max_col=1):
            for cell in row:
                for day in dateStaDef:
                    if (str(cell.value) == str(day) + " 00:00:00"):
                        wsOut['C' + str(cell.row)] = hourStart[k]
                        wsOut['D' + str(cell.row)] = hourStop[k]
                        k = k + 1
                        break

        try:
            wbOut.save("NUOVA TABELLA CONTEGGI.xlsx")
            dial = wx.MessageDialog(self, message="File generato correttamente", caption="Fine", style=wx.OK)
            dial.ShowModal()
        except:
            print("error")
            dial = wx.MessageDialog(self, message="La creazione del nuovo file non è andata a buon fine. Chiudere il programma e riprovare!", caption="Errore", style=wx.OK)
            dial.ShowModal()

        

class MainFrame(wx.Frame):

    def __init__(self):
        super().__init__(None, title="SBM - Software Tabella Conteggi", size=(800,300))
        panel = HomePanel(self)
        self.Show()


if __name__ == '__main__':
    app = wx.App(redirect=False)
    frame = MainFrame()
    app.MainLoop()


