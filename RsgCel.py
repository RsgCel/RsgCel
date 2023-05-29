#Gramazio Rocco & Ristè Thatiely & Solfanelli Davide
#4AS
#EsFinale

import wx, wx.grid

APP_NAME = "RsgCel"
EXTENSION = ".xlrsg"

ID_DOCRec = 1
ID_SalvaNome = 2
ID_SalvaCopia = 3
ID_ImpSta = 4
ID_Proprietà = 5
ID_Ripeti = 6
ID_Seleziona = 7
ID_TrovaESos = 8
ID_BarFor = 9
ID_BarStato = 10
ID_BarLat = 11
ID_Stili = 12
ID_Gal = 13
ID_Nav = 14
ID_ListaFun = 15
ID_SorDati = 16
ID_ScheInt = 17
ID_Zoom = 18
ID_Img = 20
ID_Funzione = 21
ID_Collegamento = 22
ID_CarSpeciale = 23
ID_Data = 24
ID_Ora = 25
ID_Testo = 26
ID_Spaziatura = 27
ID_Allinea = 28
ID_InfoLic=29
ID_InfoRSG=30
ID_Aiuto=31
ID_Documentazione=32
ID_Rinomina=33
ID_Donazioni=34

class Finestra(wx.Frame):

    def __init__(self):
        super().__init__(None, title=APP_NAME)

        # Spazio per le variabili membro della classe
        self.deviSalvare = False
        self.addizione = {}
        self.moltiplicazione = {}
        self.sottrazione = {}
        self.divisione = {}
        self.media = {}

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
        
        # Creazione Item Menu File
        customItemDOCRec = wx.MenuItem(fileMenu, ID_DOCRec, "Documenti recenti")
        customItemSalvaNome = wx.MenuItem(fileMenu, ID_SalvaNome, "Salva con nome ")
        customItemSalvaCopia = wx.MenuItem(fileMenu, ID_SalvaCopia, "Salva una copia")
        customItemImpSta = wx.MenuItem(fileMenu, ID_ImpSta, "Impostazioni stampante")
        customItemProprietà = wx.MenuItem(fileMenu, ID_Proprietà, "Proprietà")
        
        newItem = wx.MenuItem(fileMenu,wx.ID_NEW,"Nuovo")
        newItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_NEW))
        fileMenu.Append(newItem)
        openItem = wx.MenuItem(fileMenu, wx.ID_OPEN,"Apri")
        openItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN))
        fileMenu.Append(openItem)
        fileMenu.Append(customItemDOCRec)
        closeItem = wx.MenuItem(fileMenu, wx.ID_CLOSE,"Chiudi")
        closeItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_CLOSE))
        fileMenu.Append(closeItem)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_REFRESH, "Ricarica")
        fileMenu.AppendSeparator()
        salvaItem = wx.MenuItem(fileMenu,wx.ID_SAVE,"Salva")
        salvaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))
        fileMenu.Append(salvaItem)
        fileMenu.Append(customItemSalvaNome)
        fileMenu.Append(customItemSalvaCopia)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_PRINT, "Stampa")
        fileMenu.Append(customItemImpSta)
        fileMenu.AppendSeparator()
        fileMenu.Append(customItemProprietà)
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_EXIT, "Esci da GS Excel")
        
        mb.Append(fileMenu, '&File')
        
        # crea un menù Modifica
        editMenu = wx.Menu()
        
        # Creazione Item Menu Modifica
        customItemRipeti = wx.MenuItem(editMenu, ID_Ripeti, "Ripeti")
        customItemSele = wx.MenuItem(editMenu, ID_Seleziona, "Seleziona")
        customItemTrovaESos = wx.MenuItem(editMenu, ID_TrovaESos, "Trova e sostituisci")

        editMenu.Append(wx.ID_UNDO, "Annulla")
        editMenu.Append(wx.ID_REDO, "Ripristina")
        editMenu.Append(customItemRipeti)
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_CUT, "Tagia")
        editMenu.Append(wx.ID_COPY, "Copia")
        editMenu.Append(wx.ID_PASTE, "Incolla")
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_SELECTALL, "Seleziona tutto")
        editMenu.Append(customItemSele)
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_FIND, "Trova")
        editMenu.Append(customItemTrovaESos)
        
        mb.Append(editMenu, '&Modifica')
        
        # crea menu Visualizza
        viewMenu = wx.Menu()
        
        # Creazione Item Menu Visualizza
        customItemBarraFormula = wx.MenuItem(viewMenu, ID_BarFor, "Barra della formula")
        customItemBarraStato = wx.MenuItem(viewMenu, ID_BarStato, "Barra di stato")
        customItemBarraLat = wx.MenuItem(viewMenu, ID_BarLat, "Barra laterale")
        customItemStili = wx.MenuItem(viewMenu, ID_Stili, "Stili")
        customItemGalleria = wx.MenuItem(viewMenu, ID_Gal, "Galleria")
        customItemNavigatore = wx.MenuItem(viewMenu, ID_Nav, "Navigatore")
        customItemListaFunzioni = wx.MenuItem(viewMenu, ID_ListaFun, "Lista funzioni")
        customItemSorgenteDati = wx.MenuItem(viewMenu, ID_SorDati, "Sorgente dati")
        customItemSchermoInt = wx.MenuItem(viewMenu, ID_ScheInt, "Schermo intero")
        customItemZoom = wx.MenuItem(viewMenu, ID_Zoom, "Zoom")
        
        viewMenu.Append(customItemBarraFormula)
        viewMenu.Append(customItemBarraStato)
        viewMenu.AppendSeparator()
        viewMenu.Append(customItemBarraLat)
        viewMenu.Append(customItemStili)
        viewMenu.Append(customItemGalleria)
        viewMenu.Append(customItemNavigatore)
        viewMenu.Append(customItemListaFunzioni)
        viewMenu.Append(customItemSorgenteDati)
        viewMenu.AppendSeparator()
        viewMenu.Append(customItemSchermoInt)
        viewMenu.Append(customItemZoom)
    
        mb.Append(viewMenu, '&Visualizza')
        
        # crea Menu Inserisci
        insertMenu = wx.Menu()
        
        # Creazione Item menu Inserisci
        customItemImg = wx.MenuItem(viewMenu, ID_Img, "Immagine")
        customItemFunzione = wx.MenuItem(viewMenu, ID_Funzione, "Funzione")
        customItemCollegamento = wx.MenuItem(viewMenu, ID_Collegamento, "Collegamento")
        customItemCarattereSpeciale = wx.MenuItem(viewMenu, ID_CarSpeciale, "Carattere speciale")
        customItemData = wx.MenuItem(viewMenu, ID_Data, "Data")
        customItemOra = wx.MenuItem(viewMenu, ID_Ora, "Ora")
        
        insertMenu.Append(customItemImg)
        insertMenu.AppendSeparator()
        insertMenu.Append(customItemFunzione)
        insertMenu.AppendSeparator()
        insertMenu.Append(customItemCollegamento)
        insertMenu.Append(customItemCarattereSpeciale)
        insertMenu.AppendSeparator()
        insertMenu.Append(customItemData)
        insertMenu.Append(customItemOra)
        
        mb.Append(insertMenu, '&Inserisci')
        
        # crea menu Formato
        formatoMenu = wx.Menu()
        
        # Creazione Item Menu Formato
        customItemTesto = wx.MenuItem(viewMenu, ID_Testo, "Testo")
        customItemSpaziatura = wx.MenuItem(viewMenu, ID_Spaziatura, "Spaziatura")
        customItemAllinea = wx.MenuItem(viewMenu, ID_Allinea, "Allinea")
        
        formatoMenu.Append(customItemTesto)
        formatoMenu.Append(customItemSpaziatura)
        formatoMenu.Append(customItemAllinea)
        
        mb.Append(formatoMenu, '&Formato')
        
        #Menù stili -> ci si può mettere il font...
        stileMenu=wx.Menu()
        stileMenu.Append(wx.ID_SELECT_FONT,"Seleziona Font")
        
        mb.Append(stileMenu,"&Stile")
        
        #Menù foglio
        pageMenu=wx.Menu()
        pageMenu.Append(wx.ID_NEW,"Apri nuovo foglio di lavoro")
        pageMenu.AppendSeparator()
        pageMenu.Append(wx.ID_CLEAR,"Pulisci celle")
        pageMenu.AppendSeparator()
        pageMenu.Append(ID_Rinomina,"Rinomina foglio")
        pageMenu.AppendSeparator()
        pageMenu.Append(wx.ID_COPY,"Copia foglio di lavoro")
        
        mb.Append(pageMenu, '&Foglio')
        
        #Menù dati
        datiMenu=wx.Menu()
        datiMenu.Append(wx.ID_SORT_ASCENDING,"Ordina in modo crescente")
        datiMenu.Append(wx.ID_SORT_DESCENDING,"Ordina in modo decrescente")
        
        mb.Append(datiMenu, '&Dati')
        
        #Menù strumenti
        strMenu=wx.Menu()
        strMenu.Append(wx.ID_SPELL_CHECK,"Controllo ortografico")
        
        mb.Append(strMenu, '&Strumenti')
        
        #Menù finestra
        windowMenu = wx.Menu()
        windowMenu.Append(wx.ID_NEW,"Apri nuova finestra")
        windowMenu.Append(wx.ID_CLOSE,"Chiudi finestra")
        
        mb.Append(windowMenu, '&Finestra')
        
        #Menù Aiuto
        helpMenu = wx.Menu()     
        helpMenu.Append(ID_Aiuto, "Guida")
        helpMenu.Append(ID_Documentazione, "Documentazione programma") 
        helpMenu.AppendSeparator()
        helpMenu.Append(ID_Donazioni,"Donazione a RsgCel")
        helpMenu.AppendSeparator()
        helpMenu.Append(ID_InfoLic, "Informazioni licenza")
        helpMenu.AppendSeparator()
        helpMenu.Append(ID_InfoRSG, "Informazioni su RsgCel")
        
        mb.Append(helpMenu, '&Aiuto')
        
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
        
        # Bind Modifica
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
        
        # Bind Visualizza
        self.Bind(wx.EVT_MENU, self.funzioneBarraFormula, id=ID_BarFor)
        self.Bind(wx.EVT_MENU, self.funzioneBarraStato, id=ID_BarStato)
        self.Bind(wx.EVT_MENU, self.funzioneBarLat, id=ID_BarLat)
        self.Bind(wx.EVT_MENU, self.funzioneStili, id=ID_Stili)
        self.Bind(wx.EVT_MENU, self.funzioneGal, id=ID_Gal)
        self.Bind(wx.EVT_MENU, self.funzioneNav, id=ID_Nav)
        self.Bind(wx.EVT_MENU, self.funzioneListaFun, id=ID_ListaFun)
        self.Bind(wx.EVT_MENU, self.funzioneSorgenteDati, id=ID_SorDati)
        self.Bind(wx.EVT_MENU, self.funzioneSchermoIntero, id=ID_ScheInt)
        self.Bind(wx.EVT_MENU, self.funzioneZoom, id=ID_Zoom)
        
        # Bind Inserisci
        self.Bind(wx.EVT_MENU, self.funzioneImmagine, id=ID_Img)
        self.Bind(wx.EVT_MENU, self.funzioneFunzione, id=ID_Funzione)
        self.Bind(wx.EVT_MENU, self.funzioneCollegamento, id=ID_Collegamento)
        self.Bind(wx.EVT_MENU, self.funzioneCarattereSpeciale, id=ID_CarSpeciale)
        self.Bind(wx.EVT_MENU, self.funzioneData, id=ID_Data)
        self.Bind(wx.EVT_MENU, self.funzioneOra, id=ID_Ora)
        
        # Bind Formato
        self.Bind(wx.EVT_MENU, self.funzioneTesto, id=ID_Testo)
        self.Bind(wx.EVT_MENU, self.funzioneSpaziatura, id=ID_Spaziatura)
        self.Bind(wx.EVT_MENU, self.funzioneAllinea, id=ID_Allinea)
        
        #Bind Stili
        self.Bind(wx.EVT_MENU, self.funzioneSelFont,id=wx.ID_SELECT_FONT)

        #Bind Foglio
        self.Bind(wx.EVT_MENU, self.funzionePulisciCelle,id=wx.ID_CLEAR)
        self.Bind(wx.EVT_MENU, self.funzioneRinomina,id=ID_Rinomina)
        
        #Bind Dati
        self.Bind(wx.EVT_MENU, self.funzioneOrdinaCresc,id=wx.ID_SORT_ASCENDING)
        self.Bind(wx.EVT_MENU, self.funzioneOrdinaDecr,id=wx.ID_SORT_DESCENDING)
        
        #Bind Strumenti
        self.Bind(wx.EVT_MENU, self.funzioneCheckOrto,id=wx.ID_SPELL_CHECK)
        
        # Bind Help
        self.Bind(wx.EVT_MENU, self.funzioneInfoLic, id=ID_InfoLic)
        self.Bind(wx.EVT_MENU, self.funzioneInfoRSG, id=ID_InfoRSG)
        self.Bind(wx.EVT_MENU, self.funzioneAiuto, id=ID_Aiuto)
        self.Bind(wx.EVT_MENU, self.funzioneDocumentazione, id=ID_Documentazione)
        self.Bind(wx.EVT_MENU, self.funzioneDonazione, id=ID_Donazioni)

        return

    # in questa funzione andremo a creare e popolare la toolbar
    def creaToolbar(self):
        toolbar = self.CreateToolBar()
        
        toolbar.AddTool(wx.ID_OPEN, "Nuovo",  wx.ArtProvider.GetBitmap(wx.ART_NEW))
        toolbar.AddTool(wx.ID_OPEN, "Apri",  wx.ArtProvider.GetBitmap(wx.ART_FOLDER_OPEN))
        toolbar.AddTool(wx.ID_SAVE, "Salva", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))

        toolbar.AddSeparator()

        toolbar.AddTool(wx.ID_EXIT, "Stampa",  wx.ArtProvider.GetBitmap(wx.ART_PRINT))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_EXIT, "Taglia",  wx.ArtProvider.GetBitmap(wx.ART_CUT))
        toolbar.AddTool(wx.ID_EXIT, "Copia",  wx.ArtProvider.GetBitmap(wx.ART_COPY))
        toolbar.AddTool(wx.ID_EXIT, "Incolla",  wx.ArtProvider.GetBitmap(wx.ART_PASTE))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_EXIT, "Annulla",  wx.ArtProvider.GetBitmap(wx.ART_UNDO))
        toolbar.AddTool(wx.ID_EXIT, "Ripristina",  wx.ArtProvider.GetBitmap(wx.ART_REDO))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_EXIT, "Trova e sostituisci",  wx.ArtProvider.GetBitmap(wx.ART_FIND_AND_REPLACE))
        
        toolbar.AddSeparator()
        
        self.digitazione = wx.TextCtrl(toolbar, size=(1075,-1))
        toolbar.AddControl(self.digitazione)

        toolbar.Realize()
        
        return

    # in questa funzione aggiungeremo la statusbar
    def creaStatusbar(self):
        return

    # questa funzione implementa la vista principale del programma
    def creaMainView(self):
        panel = wx.Panel(self)
        
        mainLayout = wx.BoxSizer(wx.HORIZONTAL)
        
        vbox1 = wx.BoxSizer(wx.VERTICAL)
        
        self.mainGrid = wx.grid.Grid(panel)
        self.mainGrid.CreateGrid(100, 100)
        self.mainGrid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.cellaCambiata)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        
        indicatoreCelle = wx.ComboBox(panel, size = (150, -1))
        btnFunction = wx.Button(panel, size = (23, 23))
        btnFormula = wx.Button(panel, size = (23, 23))
        barraCella = wx.TextCtrl(panel)
        
        hbox1.Add(indicatoreCelle, proportion = 0, flag = wx.EXPAND | wx.ALL, border = 5)
        hbox1.Add(btnFunction, proportion = 0, flag = wx.EXPAND | wx.ALL, border = 5)
        hbox1.Add(btnFormula, proportion = 0, flag = wx.EXPAND | wx.ALL, border = 5)
        hbox1.Add(barraCella, proportion = 1, flag = wx.EXPAND | wx.ALL, border = 5)
        
        vbox1.Add(hbox1, proportion = 0, flag = wx.EXPAND)
        
        vbox1.Add(self.mainGrid, proportion = 1, flag = wx.EXPAND)
        
        mainLayout.Add(vbox1, proportion = 1, flag = wx.EXPAND)
        
        panel.SetSizer(mainLayout)
        
        #icona
        icon = wx.Icon()
        icon.CopyFromBitmap(wx.Bitmap("icon.png", wx.BITMAP_TYPE_ANY))
        self.SetIcon(icon)
        
        return
    def funzioneChiudi(self,evt): #CLEAR NON FUNZIONA DA RIVEDERE
        dial = wx.MessageDialog(None, "Vuoi salvare prima di chiudere il foglio di calcolo?", "Domanda", wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION)
        risposta = dial.ShowModal()
        if risposta == wx.ID_YES:
            # salva
            if not self.salva():
                return
                # chiudi
            self.mainGrid.ClearGrid()
                 
        if risposta == wx.ID_NO:
                #chiudi
            self.mainGrid.ClearGrid()
        if risposta==wx.ID_CANCEL:
                # nulla
            return
        self.deviSalvare = False
        return
    def funzioneEsci(self,evt): #NON FUNZIONA
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
            dial = wx.MessageDialog(None, "Vuoi salvare prima di chiudere?", "Domanda", wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION)
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
#         self.deviSalvare = True
        #Qui andrebbero controllate se ci sono operazioni o se la cella cambiata può cambiare il valore di qualche altra
        return
    
    def dictToString(self, dizionario:dict) -> str:
        """Creo una stringa partendo dal dizionario fornito. Per ogni riga una lista poi trasnformata in stringa di chiavi e rispettivi valori"""
        stringa = ""
        for key in dizionario:
            lista = []
            match key:
                case tuple():
                    for coord in key:
                        lista.append(str(coord))
                case str() | int():
                    lista.append(str(key))
                case other:
                    pass
            
            match dizionario[key]:
                case list():
                    for value in dizionario[key]:
                        lista.append(str(value))
                    stringa += ", ".join(lista) + "\n"
                case str() | int():
                    lista.append(dizionario[key])
                    stringa += ", ".join(lista) + "\n"
        
        return stringa
    
    def cellDictFromFileString(self, stringa:str) -> dict:
        """Prendo i valori dalla stringa fornita e li separo per metterli poi nel dizionario"""
        if stringa != "vuoto":
            dizionarioCelle = {}
            for cella in stringa.split("\n"):
                listaCella = cella.split(", ")
                row = listaCella[0]
                col = listaCella[1]
                cont = listaCella[2] # Contenuto
                tr = listaCella[3]   # TextRed
                tg = listaCella[4]   # TextGreen
                tb = listaCella[5]   # TextBlue
                ta = listaCella[6]   # TextAlpha
                r = listaCella[7]    # Red
                g = listaCella[8]    # Green
                b = listaCella[9]    #Blue
                a = listaCella[10]   #Alpha
                hAlign = listaCella[11]
                vAlign = listaCella[12]
                dizionarioCelle[(row, col)] = [str(cont), str(tr), str(tg), str(tb), str(ta), str(r), str(g), str(b), str(a), str(hAlign), str(vAlign)]
            return dizionarioCelle

    # rc = rowCol
    def rcDictFromFileString(self, rcStringa:str) -> dict:
        """Prendo i valori dalla stringa fornita e li separo per metterli poi nel dizionario"""
        if rcStringa != "vuoto":
            dizionario = {}
            for rowCol in rcStringa.split("\n"):
                listaRc = rowCol.split(", ")
                dizionario[listaRc[0]] = lisraRc[1]
        return dizionario
        
    def funzioneSalva(self, evt, stringaFile = ""): #Se gli passi il contenuto del file(di default "") puoi usarla anche per salva normale
        # Celle
        #      row, col, cont, textColour, Colour, alignment
        #      Font
        # Col
        #      width, Format
        # Row
        #      height
        
        # Quando salvi su un file già esistente controlla anche se la cella è già salvata da qualche parte
        # Formato .xlrsg
        dlg = wx.FileDialog(None, "Salva File", style=wx.FD_SAVE)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return False
        
        self.percorso = dlg.GetPath()
        
        if stringaFile == "":
            celleModificate = {}
            righeModificate = {}
            colonneModificate = {}
        else:
            listaStringaFile = stringaFile.split("Separatore\n")
            celleModificate = self.cellDictFromFileString(listaStringaFile[0])
            righeModificate = self.rcDictFromFileString(listaStringaFile[1])
            colonneModificate = self.ecDictFromFileString(listaStringaFile[2])
            
        baseHeight = self.mainGrid.GetDefaultRowSize()
        baseWidth = self.mainGrid.GetDefaultColSize()
        baseCont = ""
        (baseTr, baseTg, baseTb, baseTa) = self.mainGrid.GetDefaultCellTextColour()
        (baseR, baseG, baseB, baseA) = self.mainGrid.GetDefaultCellBackgroundColour().Get()
        (baseHAlign, baseVAlign) = self.mainGrid.GetDefaultCellAlignment()
        
        baseCell = [str(baseCont), str(baseTr), str(baseTg), str(baseTb), str(baseTa), str(baseR), str(baseG), str(baseB), str(baseA), str(baseHAlign), str(baseVAlign)]
            
        
        for row in range(self.mainGrid.GetNumberRows()):
            for col in range(self.mainGrid.GetNumberCols()):
                cont = self.mainGrid.GetCellValue(row, col)
                (tr, tg, tb, ta) = self.mainGrid.GetCellTextColour(row, col).Get()
                (r, g, b, a) = self.mainGrid.GetCellBackgroundColour(row, col).Get()
                (hAlign, vAlign) = self.mainGrid.GetCellAlignment(row, col)
                font = self.mainGrid.GetCellFont(row, col)
                
                cell = [str(cont), str(tr), str(tg), str(tb), str(ta), str(r), str(g), str(b), str(a), str(hAlign), str(vAlign)]
                
                if cell != baseCell:
                    celleModificate[(row, col)] = cell
                
                width = self.mainGrid.GetColSize(col)
                if width != baseWidth:
                    colonneModificate[col] = str(width)
            
            height = self.mainGrid.GetRowSize(row)
            if height != baseHeight:
                righeModificate[row] = str(height)
        
        stringaFile += self.dictToString(celleModificate) or "vuoto"
        
        stringaFile += "Separatore\n"
        
        stringaFile += self.dictToString(righeModificate) or "vuoto"
        
        stringaFile += "Separatore\n"
        
        stringaFile += self.dictToString(colonneModificate) or "vuoto"
        
        if EXTENSION in self.percorso:
            fileName = self.percorso
        else:
            fileName = self.percorso + EXTENSION
        
        file = open(fileName, "w")
        file.write(stringaFile)
        file.close()
        
        return
    
    def funzioneApri(self,evt):
        dlg = wx.FileDialog(None, "Apri File", style=wx.FD_OPEN, wildcard="RsgCel files (*.xlrsg)|*.xlrsg")
        if dlg.ShowModal() == wx.ID_CANCEL:
            return

        filePath = dlg.GetPath()
        
        if self.deviSalvare:
            window = Finestra()
            window.Show()
            (px,py) = self.GetPosition()
            window.Move(px + 50, py + 50)
           
            window.apri(filePath)
            return
        
        self.apri(filePath)
        return
    
    def apri(self, path):
        self.percorso = path
        
        #DA RIVEDERE
        
        #PRIMA VANNO FATTE TUTTE LE OPERAZIONI
    
        file = open(self.percorso, "r")
        contenuto = file.read()
        file.close()
        self.mainGrid.SetValue(contenuto)
        self.deviSalvare = False
        return

    # Funzioni MenuBar
    
    # Funzioni File
    def funzioneNuovo(self, evt):
        window = Finestra()
        window.Show()
        (px,py) = self.GetPosition()
        window.Move(px + 50, py + 50)
        return

    def funzioneDocumentiRecenti(self, evt):
        return

    def funzioneRicarica(self, evt):
        return

    def funzioneRicarica(self, evt):
        self.mainGrid.Refresh()
        return


    def funzioneSalvaConNome(self, evt):
        return
    
    def funzioneSalvaCopia(self, evt):
        return
    
    def funzioneStampa(self, evt):
        #PrintGrid()
        return
    
    def funzioneImpostazioniStampante(self, evt):
        return
    
    def funzioneProprietà(self, evt):
        return
   
    
    # Funzioni Modifica
    def funzioneAnnulla(self, evt):
        self.mainGrid.Undo()
        return
    
    def funzioneRipristina(self, evt):
        self.mainGrid.Redo()
        return

    def funzioneRipeti(self, evt):
        return

    def funzioneTaglia(self, evt):
        self.mainGrid.Cut()
        return

    def funzioneCopia(self, evt):
        self.mainGrid.Copy()
        return

    def funzioneIncolla(self, evt):
        self.mainGrid.Paste()
        return

    def funzioneSelezionaTutto(self, evt):
        self.mainGrid.SelectAll()
        return
    
    def funzioneSeleziona(self, evt):
        return
    
    def funzioneTrova(self, evt):
        return
    
    def funzioneTrovaeSostituisci(self, evt):
        return
    
    # Funzioni Visualizza
    def funzioneBarraFormula(self, evt):
        return
    
    def funzioneBarraStato(self, evt):
        return
    
    def funzioneBarLat(self, evt):
        return
        
    def funzioneStili(self, evt):
        return
    
    def funzioneGal(self, evt):
        return
        
    def funzioneNav(self, evt):
        return
    
    def funzioneListaFun(self, evt):
        return
    
    def funzioneSorgenteDati(self, evt):
        return
    
    def funzioneSchermoIntero(self, evt):
        return
    
    def funzioneZoom(self, evt):
        return
    
    # Funzione Inserisci
    def funzioneImmagine(self, evt):
        return
    
    def funzioneFunzione(self, evt):
        return
    
    def funzioneCollegamento(self, evt):
        return
    
    def funzioneCarattereSpeciale(self, evt):
        return
    
    def funzioneData(self, evt):
        return
    
    def funzioneOra(self, evt):
        return
    
    # Funzioni Formato
    def funzioneTesto(self, evt):
        return
    
    def funzioneSpaziatura(self, evt):
        return
    
    def funzioneAllinea(self, evt):
        return
    
    #Funzioni Stili
    def funzioneSelFont(self,evt):
        return
    
    #Funzioni Foglio
    def funzionePulisciCelle(self,evt):
        return
    def funzioneRinomina(self,evt):
        return
    
    
    #Funzioni Dati
    def funzioneOrdinaCresc(self,evt):
        return
    def funzioneOrdinaDecr(self,evt):
        return
    
    #Funzioni Strumenti
    def funzioneCheckOrto(self,evt):
        return
 
    # Funzioni Aiuto
    def funzioneInfoLic(self, evt):
        return
    def funzioneInfoRSG(self, evt):
        return
    def funzioneAiuto(self, evt):
        return
    def funzioneDocumentazione(self, evt):
        return
    def funzioneDonazione(self, evt):
        return


# ----------------------------------------
if __name__ == "__main__":
    app = wx.App()
    app.SetAppName(APP_NAME)
    window = Finestra()
    window.Show()
    app.MainLoop()
