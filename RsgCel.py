#Gramazio Rocco & Ristè Thatiely & Solfanelli Davide
#4AS
#EsFinale

import wx, wx.grid

APP_NAME = "RsgCel"

BASE_WIDTH = 80
BASE_HEIGHT = 19

ID_DOCRec = 1
ID_SalvaNome = 2
ID_SalvaCopia = 3
ID_ImpSta = 4
ID_Proprietà = 5
ID_Ripeti = 6
ID_Seleziona = 7
ID_TrovaESos = 8

class Finestra(wx.Frame):

    def __init__(self):
        super().__init__(None, title=APP_NAME)

        # Spazio per le variabili membro della classe
        self.deviSalvare = False
        self.celleModificate = {}
        self.addizione = {}
        self.moltiplicazione = {}
        self.sottrazione = {}
        self.divisione = {}

        # -------------------------------------------

        # Chiamata alle funzioni che generano la UI
        self.creaMenubar()
        self.creaToolbar()
        self.creaStatusbar()
        # -------------------------------------------

        # Chiamata alla funzione che genera la MainView
        self.creaMainView()
        # -------------------------------------------
        
        # Bind generali
        self.Bind(wx.EVT_CLOSE, self.funzioneEsci)

        # le ultime cose, ad esempio, le impostazioni iniziali, etc...
        config = wx.FileConfig(APP_NAME)
        
        w = int( config.Read( "width" , "-10" ) )
        h = int( config.Read( "height" , "-10" ) )
        
        if (w, h) == (-10, -10):
            self.Maximize()
        else:
            self.SetSize(w,h)
        
        px = int( config.Read( "px", "0") )
        py = int( config.Read( "py", "0") )
        self.Move(px,py)
        # -------------------------------------------
        return

    # in questa funzione andremo a creare e popolare la menubar
    def creaMenubar(self):
        mb = wx.MenuBar()

        # crea un menù File
        fileMenu = wx.Menu()

        fileMenu.Append(wx.ID_NEW, "Nuovo")
        fileMenu.Append(wx.ID_OPEN, "Apri")
        customItemDOCRec = wx.MenuItem(fileMenu, 1, "Documenti recenti")
        fileMenu.Append(customItemDOCRec)
        fileMenu.Append(wx.ID_CLOSE, "Chiudi")
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_REFRESH, "Ricarica")
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_SAVE, "Salva")
        customItemSalvaNome = wx.MenuItem(fileMenu, 2, "Salva con nome ")
        fileMenu.Append(customItemSalvaNome)
        customItemSalvaCopia = wx.MenuItem(fileMenu, 3, "Salva una copia")
        fileMenu.Append(customItemSalvaCopia)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_PRINT, "Stampa")
        customItemImpSta = wx.MenuItem(fileMenu, 4, "Impostazioni stampante")
        fileMenu.Append(customItemImpSta)
        fileMenu.AppendSeparator()
        customItemProprietà = wx.MenuItem(fileMenu, 5, "Proprietà")
        fileMenu.Append(customItemProprietà)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_EXIT, "Esci da GS Excel")
        
        mb.Append(fileMenu, '&File')
        
        # crea un menù Modifica
        editMenu = wx.Menu()

        editMenu.Append(wx.ID_UNDO, "Annulla")
        editMenu.Append(wx.ID_REDO, "Ripristina")
        customItemRipeti = wx.MenuItem(editMenu, 6, "Ripeti")
        editMenu.Append(customItemRipeti)
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_CUT, "Tagia")
        editMenu.Append(wx.ID_COPY, "Copia")
        editMenu.Append(wx.ID_PASTE, "Incolla")
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_SELECTALL, "Seleziona tutto")
        customItemSele = wx.MenuItem(editMenu, 7, "Seleziona")
        editMenu.Append(customItemSele)
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_FIND, "Trova")
        customItemTrovaESos = wx.MenuItem(editMenu, 8, "Trova e sostituisci")
        editMenu.Append(customItemTrovaESos)
        
        mb.Append(editMenu, '&Modifica')
        
        self.SetMenuBar(mb)
        
        # Bind Menu File
        self.Bind(wx.EVT_MENU, self.funzioneNuovo, id=wx.ID_NEW)
        self.Bind(wx.EVT_MENU, self.funzioneApri, id=wx.ID_OPEN)
        self.Bind(wx.EVT_MENU, self.funzioneDocumentiRecenti, id=ID_DOCRec)
        self.Bind(wx.EVT_MENU, self.funzioneChiudi, id=wx.ID_CLOSE)
        self.Bind(wx.EVT_MENU, self.funzioneRicarica, id=wx.ID_REFRESH)
        self.Bind(wx.EVT_MENU, self.funzioneSalva, id=wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, self.funzioneSalvaConNome, id=ID_SalvaNome)
        self.Bind(wx.EVT_MENU, self.funzioneSalvaCopia, id=ID_SalvaCopia)
        self.Bind(wx.EVT_MENU, self.funzioneStampa, id=wx.ID_PRINT)
        self.Bind(wx.EVT_MENU, self.funzioneImpostazioniStampante, id=ID_ImpSta)
        self.Bind(wx.EVT_MENU, self.funzioneProprietà, id=ID_Proprietà)
        self.Bind(wx.EVT_MENU, self.funzioneEsci, id=wx.ID_EXIT)
        
        # Bind Modifica File
        self.Bind(wx.EVT_MENU, self.funzioneAnnulla, id=wx.ID_UNDO)
        self.Bind(wx.EVT_MENU, self.funzioneRipristina, id=wx.ID_REDO)
        self.Bind(wx.EVT_MENU, self.funzioneRipeti, id=ID_Ripeti)
        self.Bind(wx.EVT_MENU, self.funzioneTaglia, id=wx.ID_CUT)
        self.Bind(wx.EVT_MENU, self.funzioneCopia, id=wx.ID_COPY)
        self.Bind(wx.EVT_MENU, self.funzioneIncolla, id=wx.ID_PASTE)
        self.Bind(wx.EVT_MENU, self.funzioneSelezionaTutto, id=wx.ID_SELECTALL)
        self.Bind(wx.EVT_MENU, self.funzioneSeleziona, id=ID_Seleziona)
        self.Bind(wx.EVT_MENU, self.funzioneTrova, id=wx.ID_FIND)
        self.Bind(wx.EVT_MENU, self.funzioneTrovaeSostituisci, id=ID_TrovaESos)

        return

    # in questa funzione andremo a creare e popolare la toolbar
    def creaToolbar(self):
        return

    # in questa funzione aggiungeremo la statusbar
    def creaStatusbar(self):
        return

    # questa funzione implementa la vista principale del programma
    def creaMainView(self):
        panel = wx.Panel(self)
        
        mainLayout = wx.BoxSizer(wx.VERTICAL)
        
        self.mainGrid = wx.grid.Grid(panel)
        self.mainGrid.CreateGrid(100, 100)
        
        self.mainGrid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.cellaCambiata)
        
        mainLayout.Add(self.mainGrid, proportion = 1, flag = wx.EXPAND)
        
        panel.SetSizer(mainLayout)
        
        #icona
        icon = wx.Icon()
        icon.CopyFromBitmap(wx.Bitmap("icon.png", wx.BITMAP_TYPE_ANY))
        self.SetIcon(icon)
        
        return
    
    def funzioneEsci(self,evt):
        #PRIMA salvo le impostazioni... male non fa!!
        config = wx.FileConfig(APP_NAME)
        
        if self.IsMaximized():
            config.Write( "width" , str(-10) )
            config.Write( "height" , str(-10) )
        else:
            (w,h) = self.GetSize()
            config.Write( "width" , str(w) )
            config.Write( "height" , str(h) )
        
        (px,py) = self.GetPosition()
        config.Write( "px" , str(px) )
        config.Write( "py" , str(py) )

        if self.deviSalvare:
            dial = wx.MessageDialog(None, "Vuoi salvare prima di chiudere?", "Paraculo", wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION)
            risposta = dial.ShowModal()
            
            match risposta:
                case wx.ID_YES:
                    #salva
                    if not self.salva():
                        return
                    #chiudi
                    self.Destroy()
                    
                case wx.ID_NO:
                    #chiudi
                    self.Destroy()
                    
                case wx.ID_CANCEL:
                    #nulla
                    return
                
        self.Destroy()
        return
    
    def cellaCambiata(self, evt):
        row = evt.GetRow()
        col = evt.GetCol()
        
        cont = self.mainGrid.GetCellValue(row, col)
        (tr, tg, tb, ta) = self.mainGrid.GetCellTextColour(row, col).Get()
        (r, g, b, a) = self.mainGrid.GetCellBackgroundColour(row, col).Get()
        (hAlign, vAlign) = self.mainGrid.GetCellAlignment(row, col)
        font = self.mainGrid.GetCellFont(row, col)
        
        self.celleModificate[(row, col)] = [str(cont), str(tr), str(tg), str(tb), str(ta), str(r), str(g), str(b), str(a), str(hAlign), str(vAlign)]
        
        return
    
    def dictToString(self, dizionario):
        stringa = ""
        for key in dizionario:
            lista = []
            match key:
                case tuple():
                    for coord in key:
                        lista.append(str(coord))
                case str():
                    lista.append(str(key))
                case other:
                    pass
            for value in dizionario[key]:
                lista.append(str(value))
            stringa += ", ".join(lista) + "\n"
        
        return stringa
        
        
    def funzioneSalva(self, evt, stringaFile = ""): #Se gli passi il contenuto del file(di default "") puoi usarla anche per salva normale
        # Celle
        #      row, col, cont, textColour, Colour, alignment
        #      Font
        # Col
        #      width, Format
        # Row
        #      height
        
        # Quando salvi su un file già esistente controlla anche se la cella è già salvata da qualche parte
        # Formato .esgs
        dlg = wx.FileDialog(None, "Salva File", style=wx.FD_SAVE)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return False

        # la stringa che contiene il percorso della cartella selezionata
        self.percorso = dlg.GetPath()
        
        stringaFile += self.dictToString(self.celleModificate)
        
        stringaFile += "Separatore\n"
            
        for a in range(100):
            width = self.mainGrid.GetColSize(a)
            if width != BASE_WIDTH:
                stringaFile += str(a) + ", " + str(width) + "\n"
        
        stringaFile += "Separatore\n"
            
        for a in range(100):
            height = self.mainGrid.GetRowSize(a)
            if height != BASE_HEIGHT:
                stringaFile += str(a) + ", " + str(height)
        
        file = open(self.percorso + ".xlrsg", "w")
        file.write(stringaFile)
        file.close()
        
        return
    
    def funzioneApri(self, evt):
        return

    # Funzioni MenuBar
    
    # Funzioni File
    def funzioneNuovo(self, evt):
        return
    
    def funzioneApri(self, evt):
        return

    def funzioneDocumentiRecenti(self, evt):
        return

    def funzioneChiudi(self, evt):
        return

    def funzioneRicarica(self, evt):
        return

#     def funzioneSalva(self, evt):
#         return

    def funzioneSalvaConNome(self, evt):
        return
    
    def funzioneSalvaCopia(self, evt):
        return
    
    def funzioneStampa(self, evt):
        return
    
    def funzioneImpostazioniStampante(self, evt):
        return
    
    def funzioneProprietà(self, evt):
        return
    
#     def funzioneEsci(self, evt):
#         return
    
    # Funzioni Modifica
    def funzioneAnnulla(self, evt):
        return
    
    def funzioneRipristina(self, evt):
        return

    def funzioneRipeti(self, evt):
        return

    def funzioneTaglia(self, evt):
        return

    def funzioneCopia(self, evt):
        return

    def funzioneIncolla(self, evt):
        return

    def funzioneSelezionaTutto(self, evt):
        return
    
    def funzioneSeleziona(self, evt):
        return
    
    def funzioneTrova(self, evt):
        return
    
    def funzioneTrovaeSostituisci(self, evt):
        return


# ----------------------------------------
if __name__ == "__main__":
    app = wx.App()
    app.SetAppName(APP_NAME)
    window = Finestra()
    window.Show()
    app.MainLoop()
