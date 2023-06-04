#Gramazio Rocco & Ristè Thatiely & Solfanelli Davide
#4AS
#EsFinale

import wx, wx.grid
import webbrowser
import datetime
APP_NAME = "RsgCel"
EXTENSION = ".xlrsg"
TITOLO_INIZIALE = "(Senza Titolo)"

ID_DOCRec = 1
ID_SalvaNome = 2
ID_SalvaCopia = 3
ID_Proprietà = 5
ID_Ripeti = 6
ID_Seleziona = 7
ID_TrovaESos = 8
ID_BarFor = 9
ID_BarStato = 10
ID_BarLat = 11
ID_Stili = 12
ID_Nav = 14
ID_ListaFun = 15
ID_SorDati = 16
ID_ScheInt = 17
ID_Zoom = 18
ID_Img = 20
ID_Collegamento = 22
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

ID_SelezionaColoreSfondo = 35
ID_AllineaInAlto = 36
ID_AllineaAlCentroVerticalmente = 37
ID_AllineaInBasso = 38

def open_webpage(url):
    webbrowser.open(url)
    
class Finestra(wx.Frame):

    def __init__(self):
        super().__init__(None, title=TITOLO_INIZIALE + " - " + APP_NAME)

        # Spazio per le variabili membro della classe
        self.deviSalvare = False
        self.somma = {}
        self.moltiplicazione = {}
        self.sottrazione = {}
        self.divisione = {}
        self.media = {}
        self.massimo = {}
        self.minimo = {}
        self.percorso = TITOLO_INIZIALE

        # -------------------------------------------

        # Chiamata alle funzioni che generano la UI
        self.creaMenubar()
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
        oggettoImg = wx.Bitmap("ricarica.png")
        ricarica = oggettoImg.ConvertToImage()
        ricarica.Rescale(27,27)
        ricaricaItem = wx.MenuItem(fileMenu, wx.ID_REFRESH,"Ricarica")
        ricaricaItem.SetBitmap(ricarica)
        fileMenu.Append(ricaricaItem)
        fileMenu.AppendSeparator()
        salvaItem = wx.MenuItem(fileMenu,wx.ID_SAVE,"Salva")
        salvaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))
        fileMenu.Append(salvaItem)
        fileMenu.Append(customItemSalvaNome)
        fileMenu.Append(customItemSalvaCopia)
        fileMenu.AppendSeparator()
        stampaItem = wx.MenuItem(fileMenu, wx.ID_PRINT,"Stampa")
        stampaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_PRINT))
        fileMenu.Append(stampaItem)
        fileMenu.AppendSeparator()
        fileMenu.Append(customItemProprietà)
        fileMenu.AppendSeparator()
        esciItem = wx.MenuItem(fileMenu, wx.ID_EXIT,"Esci da GS Excel")
        esciItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_QUIT))
        fileMenu.Append(esciItem)
        
        mb.Append(fileMenu, '&File')
        
        # crea un menù Modifica
        editMenu = wx.Menu()
        
        # Creazione Item Menu Modifica
        customItemRipeti = wx.MenuItem(editMenu, ID_Ripeti, "Ripeti")
        customItemSele = wx.MenuItem(editMenu, ID_Seleziona, "Seleziona")
        customItemTrovaESos = wx.MenuItem(editMenu, ID_TrovaESos, "Trova e sostituisci")


        annullaItem = wx.MenuItem(editMenu,wx.ID_UNDO,"Annulla")
        annullaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_UNDO))
        editMenu.Append(annullaItem)
        ripristinaItem = wx.MenuItem(editMenu,wx.ID_REDO, "Ripristina")
        ripristinaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_REDO))
        editMenu.Append(ripristinaItem)
        editMenu.Append(customItemRipeti)
        editMenu.AppendSeparator()
        tagliaItem = wx.MenuItem(editMenu, wx.ID_CUT,"Taglia")
        tagliaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_CUT))
        editMenu.Append(tagliaItem)
        copiaItem = wx.MenuItem(editMenu, wx.ID_COPY,"Copia")
        copiaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_COPY))
        editMenu.Append(copiaItem)
        incollaItem = wx.MenuItem(editMenu,wx.ID_PASTE, "Incolla")
        incollaItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_PASTE))
        editMenu.Append(incollaItem)
        editMenu.AppendSeparator()
        editMenu.Append(wx.ID_SELECTALL, "Seleziona tutto")
        editMenu.Append(customItemSele)
        editMenu.AppendSeparator()
        oggettoImg = wx.Bitmap("trova.png")
        trova = oggettoImg.ConvertToImage()
        trova.Rescale(27,27)
        TrovaItem = wx.MenuItem(editMenu, wx.ID_FIND, "Trova")
        TrovaItem.SetBitmap(trova)
        editMenu.Append(TrovaItem)
        editMenu.Append(customItemTrovaESos)
        
        mb.Append(editMenu, '&Modifica')
        
        # crea menu Visualizza
        viewMenu = wx.Menu()
        
        # Creazione Item Menu Visualizza
        customItemBarraFormula = wx.MenuItem(viewMenu, ID_BarFor, "Barra della formula")
        customItemBarraStato = wx.MenuItem(viewMenu, ID_BarStato, "Barra di stato")
        customItemBarraLat = wx.MenuItem(viewMenu, ID_BarLat, "Barra laterale")
        customItemStili = wx.MenuItem(viewMenu, ID_Stili, "Stili")
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
        oggettoImmagine = wx.Bitmap("galleria.png")
        immagine = oggettoImmagine.ConvertToImage()
        immagine.Rescale(23,23)
        customItemImg= wx.MenuItem(insertMenu, ID_Img, "Immagine")
        customItemImg.SetBitmap(immagine)
        insertMenu.Append(customItemImg)
        insertMenu.AppendSeparator()
        oggettoImmagine = wx.Bitmap("collegamento.png")
        col = oggettoImmagine.ConvertToImage()
        col.Rescale(23,23)
        customItemCollegamento= wx.MenuItem(insertMenu,ID_Collegamento, "Collegamento")
        customItemCollegamento.SetBitmap(col)
        insertMenu.Append(customItemCollegamento)
        insertMenu.AppendSeparator()
        oggettoImg = wx.Bitmap("data.png")
        data = oggettoImg.ConvertToImage()
        data.Rescale(23,23)
        DataItem = wx.MenuItem(insertMenu, ID_Data, "Data")
        DataItem.SetBitmap(data)
        insertMenu.Append(DataItem)

        oggettoImg = wx.Bitmap("ora.png")
        ora = oggettoImg.ConvertToImage()
        ora.Rescale(23,23)
        OraItem = wx.MenuItem(insertMenu, ID_Ora, "Ora")
        OraItem.SetBitmap(ora)
        insertMenu.Append(OraItem)
   
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
        oggettoImmagine = wx.Bitmap("pulisci.png")
        pulisci = oggettoImmagine.ConvertToImage()
        pulisci.Rescale(23, 23)
        pulisciCelle = wx.MenuItem(pageMenu, wx.ID_CLEAR,"Pulisci celle")
        pulisciCelle.SetBitmap(pulisci)
        pageMenu.Append(pulisciCelle)
        pageMenu.AppendSeparator()
        pageMenu.Append(ID_Rinomina,"Rinomina foglio")
        pageMenu.AppendSeparator()
        pageMenu.Append(wx.ID_COPY,"Copia foglio di lavoro")
        
        mb.Append(pageMenu, '&Foglio')
        
        #Menù dati
        datiMenu=wx.Menu()
        oggettoImmagine = wx.Bitmap("az.png")
        cr = oggettoImmagine.ConvertToImage()
        cr.Rescale(23, 23)
        Crescente = wx.MenuItem(datiMenu,wx.ID_SORT_ASCENDING,"Ordina in modo crescente")
        Crescente.SetBitmap(cr)
        datiMenu.Append(Crescente)
        
        oggettoImmagine = wx.Bitmap("za.png")
        decr= oggettoImmagine.ConvertToImage()
        decr.Rescale(23, 23)
        Decrescente = wx.MenuItem(datiMenu,wx.ID_SORT_DESCENDING,"Ordina in modo decrescente")
        Decrescente.SetBitmap(decr)
        datiMenu.Append(Decrescente)
        
        mb.Append(datiMenu, '&Dati')
        
        #Menù strumenti
        strMenu=wx.Menu()
        strMenu.Append(wx.ID_SPELL_CHECK,"Controllo ortografico")
        
        mb.Append(strMenu, '&Strumenti')
        
        #Menù finestra
        windowMenu = wx.Menu()
        oggettoImg = wx.Bitmap("newWindow.png")
        new = oggettoImg.ConvertToImage()
        new.Rescale(23,23)
        newItem = wx.MenuItem(windowMenu, wx.ID_NEW,"Apri nuova finestra")
        newItem.SetBitmap(new)
        windowMenu.Append(newItem)
        
        closeItem = wx.MenuItem(windowMenu, wx.ID_CLOSE,"Chiudi finestra")
        closeItem.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_CLOSE))
        windowMenu.Append(closeItem)
        
        mb.Append(windowMenu, '&Finestra')
        
        #Menù Aiuto
        helpMenu = wx.Menu()
        oggettoImg = wx.Bitmap("guida.png")
        guida = oggettoImg.ConvertToImage()
        guida.Rescale(27,27)
        guidaItem = wx.MenuItem(helpMenu, ID_Aiuto,"Guida")
        guidaItem.SetBitmap(guida)
        helpMenu.Append(guidaItem)
        helpMenu.Append(ID_Documentazione, "Documentazione programma") 
        helpMenu.AppendSeparator()
        oggettoImg = wx.Bitmap("donazione.png")
        dona = oggettoImg.ConvertToImage()
        dona.Rescale(27,27)
        donaItem = wx.MenuItem(helpMenu, ID_Donazioni,"Donazione a RsgCel")
        donaItem.SetBitmap(dona)
        helpMenu.Append(donaItem)
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
        self.Bind(wx.EVT_MENU, self.funzioneNav, id=ID_Nav)
        self.Bind(wx.EVT_MENU, self.funzioneListaFun, id=ID_ListaFun)
        self.Bind(wx.EVT_MENU, self.funzioneSorgenteDati, id=ID_SorDati)
        self.Bind(wx.EVT_MENU, self.funzioneSchermoIntero, id=ID_ScheInt)
        self.Bind(wx.EVT_MENU, self.funzioneZoom, id=ID_Zoom)
        
        # Bind Inserisci
        self.Bind(wx.EVT_MENU, self.funzioneImmagine, id=ID_Img)
        self.Bind(wx.EVT_MENU, self.funzioneCollegamento, id=ID_Collegamento)
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
    def creaToolbar(self, toolbar):
        toolbar.AddTool(wx.ID_NEW, "Nuovo",  wx.ArtProvider.GetBitmap(wx.ART_NEW))
        toolbar.AddTool(wx.ID_OPEN, "Apri",  wx.ArtProvider.GetBitmap(wx.ART_FOLDER_OPEN))
        toolbar.AddTool(wx.ID_SAVE, "Salva", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))

        toolbar.AddSeparator()

        toolbar.AddTool(wx.ID_PRINT, "Stampa",  wx.ArtProvider.GetBitmap(wx.ART_PRINT))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_CUT, "Taglia",  wx.ArtProvider.GetBitmap(wx.ART_CUT))
        toolbar.AddTool(wx.ID_COPY, "Copia",  wx.ArtProvider.GetBitmap(wx.ART_COPY))
        toolbar.AddTool(wx.ID_PASTE, "Incolla",  wx.ArtProvider.GetBitmap(wx.ART_PASTE))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_UNDO, "Annulla",  wx.ArtProvider.GetBitmap(wx.ART_UNDO))
        toolbar.AddTool(wx.ID_REDO, "Ripristina",  wx.ArtProvider.GetBitmap(wx.ART_REDO))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(ID_TrovaESos, "Trova e sostituisci",  wx.ArtProvider.GetBitmap(wx.ART_FIND_AND_REPLACE))
        
        toolbar.AddSeparator()
        
        bottoneCaratteri = wx.Button(toolbar, label="Scegli Carattere", size=(115,25))
        bottoneCaratteri.Bind(wx.EVT_BUTTON, self.funzioneScegliCarattere)
        toolbar.AddControl(bottoneCaratteri)

        toolbar.Realize()
        return
    
    def creaToolbar2(self, toolbar):
        caratereComboBox = wx.ComboBox(toolbar, size = (185, -1))
        toolbar.AddControl(caratereComboBox)
        
        spazioVuoto = wx.StaticText(toolbar, size = (9, -1))
        toolbar.AddControl(spazioVuoto)
        
        
        grandezzaComboBox = wx.ComboBox(toolbar, size = (60, -1))
        toolbar.AddControl(grandezzaComboBox)
        
        toolbar.AddSeparator()

        toolbar.AddCheckTool(wx.ID_BOLD, "Grassetto",  self.toolBarImage("bold.png"))
        toolbar.AddCheckTool(wx.ID_ITALIC, "Corsivo",  self.toolBarImage("corsivo.png"))
        toolbar.AddCheckTool(wx.ID_UNDERLINE, "Sottolineato",  self.toolBarImage("sottolineato.png"))
        
        toolbar.AddSeparator()
        
        toolbar.AddTool(wx.ID_SELECT_COLOR, "Colore carattere",  self.toolBarImage("coloreFont.png"))
        toolbar.AddTool(ID_SelezionaColoreSfondo, "Colore carattere",  self.toolBarImage("backgroundcolor.png"))
        
        toolbar.AddSeparator()
        
        toolbar.AddRadioTool(wx.ID_JUSTIFY_LEFT, "Allinea a sinistra",  self.toolBarImage("alignleft.png"))
        toolbar.AddRadioTool(wx.ID_JUSTIFY_CENTER, "Allinea al centro",  self.toolBarImage("alignhorizontalcenter.png"))
        toolbar.AddRadioTool(wx.ID_JUSTIFY_RIGHT, "Allinea a destra",  self.toolBarImage("alignright.png"))
        
        toolbar.AddSeparator()
        
        toolbar.AddRadioTool(ID_AllineaInAlto, "Allinea a sinistra",  self.toolBarImage("aligntop.png"))
        toolbar.AddRadioTool(ID_AllineaAlCentroVerticalmente, "Allinea al centro",  self.toolBarImage("alignverticalcenter.png"))
        toolbar.AddRadioTool(ID_AllineaInBasso, "Allinea a destra",  self.toolBarImage("alignbottom.png"))
        
        toolbar.Realize()
        return
    
    #Prendo l'immagine grande e la rimpicciolisco a 24x24 tenendo una altra qualità
    def toolBarImage(self, imagePath):
        image = wx.Bitmap(imagePath).ConvertToImage()
        return image.Scale(24, 24, quality = wx.IMAGE_QUALITY_HIGH)

    # in questa funzione aggiungeremo la statusbar
    def creaStatusbar(self):
        self.statusBar = self.CreateStatusBar()
        self.statusBar.SetFieldsCount(8)
        self.statusBar.SetStatusText("Salvato")
        listaOperazioni = ["Somma", "Sottrazione", "Moltiplicazione", "Divisione", "Media", "Massimo", "Minimo"]
        for a in listaOperazioni:
            field = listaOperazioni.index(a) + 1
            self.statusBar.SetStatusText(a + ": 0", field)
        return

    # questa funzione implementa la vista principale del programma
    def creaMainView(self):
        panel = wx.Panel(self)
        
        mainLayout = wx.BoxSizer(wx.VERTICAL)
        
        toolbar1 = wx.ToolBar(panel, style=wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.creaToolbar(toolbar1)
        
        toolbar2 = wx.ToolBar(panel, style=wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.creaToolbar2(toolbar2)
        
        mainLayout.Add(toolbar1, proportion = 0)
        mainLayout.Add(toolbar2, proportion = 0)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        
        vbox1 = wx.BoxSizer(wx.VERTICAL)
        
        self.mainGrid = wx.grid.Grid(panel, style =  wx.TE_PROCESS_ENTER)
        self.mainGrid.CreateGrid(100, 100)
        self.mainGrid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.cellaCambiata)
        self.mainGrid.Bind(wx.grid.EVT_GRID_EDITOR_SHOWN, self.editorMostrato)
        self.mainGrid.Bind(wx.grid.EVT_GRID_EDITOR_HIDDEN, self.editorNascosto)
        self.mainGrid.Bind(wx.EVT_TEXT, self.cellaInCambiamento)
        self.mainGrid.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.cursoreSpostato)
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        
        self.indicatoreCelle = wx.TextCtrl(panel, size = (150, -1), value = "A1", style =  wx.TE_PROCESS_ENTER)
        staticTextVuota = wx.StaticText(panel, size = (25, -1))
        self.barraCella = wx.TextCtrl(panel)
        
        self.indicatoreCelle.Bind(wx.EVT_TEXT_ENTER, self.aggiornaPos)
        self.barraCella.Bind(wx.EVT_TEXT, self.aggiornaCella)
        self.barraCella.Bind(wx.EVT_KILL_FOCUS, self.operazioniOnLostFocus)
        
        hbox2.Add(self.indicatoreCelle, proportion = 0, flag = wx.EXPAND | wx.ALL, border = 5)
        hbox2.Add(staticTextVuota, proportion = 0, flag = wx.EXPAND | wx.ALL, border = 5)
        hbox2.Add(self.barraCella, proportion = 1, flag = wx.EXPAND | wx.ALL, border = 5)
        
        vbox1.Add(hbox2, proportion = 0, flag = wx.EXPAND)
        
        vbox1.Add(self.mainGrid, proportion = 1, flag = wx.EXPAND)
        
        hbox1.Add(vbox1, proportion = 1, flag = wx.EXPAND)
        
        mainLayout.Add(hbox1, proportion = 1, flag = wx.EXPAND)
        
        panel.SetSizer(mainLayout)
        
        #icona
        icon = wx.Icon()
        icon.CopyFromBitmap(wx.Bitmap("icon.png", wx.BITMAP_TYPE_ANY))
        self.SetIcon(icon)
        
        #Timer scritta StatusBar
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.aggiornaStatusBarSalvataggio, self.timer)
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
    
    def calcoloSomma(self, cellCoordsList:list) -> str:
        listaValori = []
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        for a in listaValori:
            if not a.isnumeric():
                stringa = True
                break
        else:
            stringa = False
        
        if stringa:
            somma = ""
            for a in listaValori:
                somma += a
        else:
            somma = 0
            for a in listaValori:
                somma += float(a)
            if int(somma) == somma:
                somma = int(somma)
        return str(somma)
    
    def calcoloSottrazione(self, cellCoordsList:list) -> str:
        listaValori = []
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        for a in listaValori:
            if not a.isnumeric():
                return "ERROR"
        
        listaValoriFloat = []
        for a in listaValori:
            listaValoriFloat.append(float(a))
        
        sottrazione = listaValoriFloat.pop(0)
        for a in listaValoriFloat:
            sottrazione -= a
        
        if int(sottrazione) == sottrazione:
                sottrazione = int(sottrazione)
        
        return str(sottrazione)
    
    def calcoloMoltiplicazione(self, cellCoordsList:list) -> str:
        listaValori = []
        
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        listaValoriNumerici = []
        stringa = ""
        for a in listaValori:
            if a.isnumeric():
                listaValoriNumerici.append(float(a))
            else:
                if len(stringa) == 0:
                    stringa = a
                else:
                    return "ERROR"
        
        moltiplicazione = 1
        for a in listaValoriNumerici:
            moltiplicazione *= a
        
        if int(moltiplicazione) == moltiplicazione:
            moltiplicazione = int(moltiplicazione)
        else:
            return "ERROR"
        
        if len(stringa) != 0:
            return str(stringa * moltiplicazione)
        return str(moltiplicazione)
    
    def calcoloDivisione(self, cellCoordsList:list) -> str:
        listaValori = []
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        for a in listaValori:
            if not a.isnumeric():
                return "ERROR"
        
        listaValoriFloat = []
        for a in listaValori:
            listaValoriFloat.append(float(a))
        
        divisione = listaValoriFloat.pop(0)
        for a in listaValoriFloat:
            divisione /= a
        
        if int(divisione) == divisione:
                divisione = int(divisione)
        
        return str(divisione)
    
    def calcoloMedia(self, cellCoordsList:list) -> str:
        media = float(self.calcoloSomma(cellCoordsList)) / len(cellCoordsList)
        if int(media) == media:
            return str(int(media))
        return str(media)
    
    def calcoloMax(self, cellCoordsList:list) -> str:
        listaValori = []
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        for a in listaValori:
            if not a.isnumeric():
                return "ERROR"
        
        listaValoriFloat = []
        for a in listaValori:
            listaValoriFloat.append(float(a))
        
        massimo = max(listaValoriFloat)
        
        if int(massimo) == massimo:
            massimo = int(massimo)
        
        return str(massimo)

    def calcoloMin(self, cellCoordsList:list) -> str:
        listaValori = []
        for a in cellCoordsList:
            listaValori.append(self.mainGrid.GetCellValue(a[0], a[1]))
        
        for a in listaValori:
            if not a.isnumeric():
                return "ERROR"
        
        listaValoriFloat = []
        for a in listaValori:
            listaValoriFloat.append(float(a))
        
        minimo = min(listaValoriFloat)
        
        if int(minimo) == minimo:
            minimo = int(minimo)
        
        return str(minimo)
    
    def alfanumericCellsListToCoordCellList(self, listaCelle:list) -> list:
        listaCoordsCelle = []
        for a in listaCelle:
            listaCoordsCelle.append(self.alfanumericToNumberCellCoord(a))
        return listaCoordsCelle
    
    def alfanumericToNumberCellCoord(self, stringa:str) -> tuple:
        for a in stringa:
            if a.isnumeric():
                tupla = stringa.partition(a)
                break
        col = tupla[0]
        
        for a in range(self.mainGrid.GetNumberCols()):
            if self.mainGrid.GetColLabelValue(a) == col:
                colNumber = a
                break
        
        rowNumber = tupla[1] + tupla[2]
        return (int(rowNumber) - 1, int(colNumber))
    
    def aggiornaPos(self, evt):
        value = evt.GetEventObject().GetValue()
        (row, col) = self.alfanumericToNumberCellCoord(value)
        self.mainGrid.GoToCell(row, col)
        self.mainGrid.SetFocus()        
        evt.Skip()
        return
    
    def aggiornaCella(self, evt):
        if evt.GetEventObject().IsModified():
            self.deviSalvare = True
            self.statusBar.SetStatusText("Non salvato")
            row = self.mainGrid.GetGridCursorRow()
            col = self.mainGrid.GetGridCursorCol()
            cont = evt.GetEventObject().GetValue()
            if cont != "":
                self.mainGrid.SetCellValue(row, col, cont)
        return
    
    def operazioniOnLostFocus(self, evt):
        if evt.GetEventObject().IsModified():
            row = self.mainGrid.GetGridCursorRow()
            col = self.mainGrid.GetGridCursorCol()
            cont = evt.GetEventObject().GetValue()
            if "=" in cont:
                self.operazioni(row, col)
            self.cellaOperazione(row, col)
        return
        
    
    def editorNascosto(self, evt):
        row = evt.GetRow()
        col = evt.GetCol()
        row = evt.GetRow()
        col = evt.GetCol()
        cont = self.mainGrid.GetCellValue(row, col)
        if "=" in cont:
            self.operazioni(row, col)
        return
    
    def editorMostrato(self, evt):
        row = evt.GetRow()
        col = evt.GetCol()
        cont = self.mainGrid.GetCellValue(row, col)
        cont = self.aggiornaBarraCella(row, col, cont)
        self.mainGrid.SetCellValue(row, col, cont)
        return
    
    def cursoreSpostato(self, evt):
        row = evt.GetRow()
        col = evt.GetCol()
        cont = self.mainGrid.GetCellValue(row, col)
        self.aggiornaIndicatorePos(evt)
        self.aggiornaBarraCella(row, col, cont)
        return
    
    def aggiornaIndicatorePos(self, evt):
        row = evt.GetRow()
        col = evt.GetCol()
        self.aggiornaComboBox(row, col)
        return
    
    def cellaInCambiamento(self, evt):
#         self.deviSalvare = True
#         self.statusBar.SetStatusText("Non salvato")
        row = self.mainGrid.GetGridCursorRow()
        col = self.mainGrid.GetGridCursorCol()
        cont = self.mainGrid.GetCellEditor(row, col).GetValue()
        self.barraCella.SetValue(cont)
        return
    
    def aggiornaBarraCella(self, row, col, cont):
        cella = (row, col)
        for a in self.somma:
            if a == cella:
                cont = self.creaContenutoOperazione(self.somma[a], "+")
                break
        
        for a in self.sottrazione:
            if a == cella:
                cont = self.creaContenutoOperazione(self.sottrazione[a], "-")
                break
        
        for a in self.divisione:
            if a == cella:
                cont = self.creaContenutoOperazione(self.divisione[a], "/")
                break
        
        for a in self.moltiplicazione:
            if a == cella:
                cont = self.creaContenutoOperazione(self.moltiplicazione[a], "*")
                break
        
        for a in self.media:
            if a == cella:
                cont = self.creaContenutoOperazione(self.media[a], "MEDIA")
                break
        
        for a in self.massimo:
            if a == cella:
                cont = self.creaContenutoOperazione(self.massimo[a], "MAX")
                break
        
        for a in self.minimo:
            if a == cella:
                cont = self.creaContenutoOperazione(self.minimo[a], "MIN")
                break
        
        self.barraCella.SetValue(cont)
        return cont
    
    def cellaCambiata(self, evt):
        self.deviSalvare = True
        self.statusBar.SetStatusText("Non salvato")
        #Qui andrebbero controllate se ci sono operazioni o se la cella cambiata può cambiare il valore di qualche altra
        row = evt.GetRow()
        col = evt.GetCol()
        cont = self.mainGrid.GetCellValue(row, col)
        if "=" in cont:
            self.operazioni(row, col)
        self.cellaOperazione(row, col)
        evt.Skip()
        return
    
    def aggiornaComboBox(self, row, col):
        cella = str(self.mainGrid.GetColLabelValue(col)) + str(row + 1)
        self.indicatoreCelle.SetValue(cella)
        return
    
    def cellaOperazione(self, row:int, col:int) -> None:
        cella = (row, col)
        for a in self.somma:
            for b in self.somma[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "+")
        
        for a in self.sottrazione:
            for b in self.sottrazione[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "-")
        
        for a in self.moltiplicazione:
            for b in self.moltiplicazione[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "*")
        
        for a in self.divisione:
            for b in self.divisione[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "/")
        
        for a in self.media:
            for b in self.media[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "MEDIA")
        
        for a in self.massimo:
            for b in self.massimo[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "MAX")
        
        for a in self.minimo:
            for b in self.minimo[a]:
                if b == cella:
                    self.operazioni(a[0], a[1], main = False, op = "MIN")
            
    def operazioni(self, row:int, col:int, main = True, op = "") -> None:
        cont = self.mainGrid.GetCellEditor(row, col).GetValue()
        cont = cont.replace("=", "")
        cont = cont.replace(" ", "")
        if "+" in cont or "+" in op:
            if main:
                listaCelle = cont.split("+")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.somma[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.somma[(row, col)]
            
            
            self.statusBar.SetStatusText("Somma: " + str(len(self.somma)), 1)
            output = self.calcoloSomma(listaCoordCelle)
            
        elif "-" in cont or "-" in op:
            if main:
                listaCelle = cont.split("-")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.sottrazione[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.sottrazione[(row, col)]
            
            self.statusBar.SetStatusText("Sottrazione: " + str(len(self.sottrazione)), 2)
            output = self.calcoloSottrazione(listaCoordCelle)
            
        elif "*" in cont or "*" in op:
            if main:
                listaCelle = cont.split("*")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.moltiplicazione[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.moltiplicazione[(row, col)]
            
            self.statusBar.SetStatusText("Moltiplicazione: " + str(len(self.moltiplicazione)), 2)
            output = self.calcoloMoltiplicazione(listaCoordCelle)
        
        elif "/" in cont or "/" in op:
            if main:
                listaCelle = cont.split("/")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.divisione[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.divisione[(row, col)]
            
            self.statusBar.SetStatusText("Divisione: " + str(len(self.divisione)), 3)
            output = self.calcoloDivisione(listaCoordCelle)
            
        elif "MEDIA" in cont or "MEDIA" in op:
            if main:
                cont = cont.replace(")", "")
                cont = cont.replace("MEDIA(", "")
                listaCelle = cont.split(";")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.media[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.media[(row, col)]
            
            self.statusBar.SetStatusText("Media: " + str(len(self.media)), 4)
            output = self.calcoloMedia(listaCoordCelle)
        
        elif "MAX" in cont or "MAX" in op:
            if main:
                cont = cont.replace(")", "")
                cont = cont.replace("MAX(", "")
                listaCelle = cont.split(";")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.massimo[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.massimo[(row, col)]
            
            self.statusBar.SetStatusText("Massimo: " + str(len(self.massimo)), 5)
            output = self.calcoloMax(listaCoordCelle)
        
        elif "MIN" in cont or "MIN" in op:
            if main:
                cont = cont.replace(")", "")
                cont = cont.replace("MIN(", "")
                listaCelle = cont.split(";")
                listaCoordCelle = self.alfanumericCellsListToCoordCellList(listaCelle)
                self.minimo[(row, col)] = listaCoordCelle
            else:
                listaCoordCelle = self.minimo[(row, col)]
            
            self.statusBar.SetStatusText("Minimo: " + str(len(self.minimo)), 6)
            output = self.calcoloMin(listaCoordCelle)
            
        self.mainGrid.SetCellValue(row, col, output)
        return
    
    def dictToString(self, dizionario:dict) -> str:
        """Creo una stringa partendo dal dizionario fornito. Per ogni riga una lista poi trasnformata in stringa di chiavi e rispettivi valori"""
        stringa = ""
        if dizionario != None:
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
        return ""
    
    def creaContenutoOperazione(self, cellCoordsList, segno):
        listaCelle = []
        for a in cellCoordsList:
            row = a[0]
            col = a[1]
            colLabel = self.mainGrid.GetColLabelValue(col)
            listaCelle.append(str(colLabel) + str(row + 1))
        if not segno in ("MEDIA", "MAX", "MIN"):
            cont = "=" + segno.join(listaCelle)
        elif segno == "MEDIA":
            cont = "=MEDIA(" + ";".join(listaCelle) + ")"
        elif segno == "MAX":
            cont = "=MAX(" + ";".join(listaCelle) + ")"
        elif segno == "MIN":
            cont = "=MIN(" + ";".join(listaCelle) + ")"
        return cont
                
    
    def cellDictFromFileString(self, stringa:str) -> dict:
        """Prendo i valori dalla stringa fornita e li separo per metterli poi nel dizionario"""
        if stringa != "vuoto\n":
            dizionarioCelle = {}
            listaCelle = stringa.split("\n")
            if listaCelle[-1] == "":
                listaCelle.pop(-1)
            for cella in listaCelle:
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
                fontString = listaCella[13]
                
                dizionarioCelle[(row, col)] = [str(cont), str(tr), str(tg), str(tb), str(ta), str(r), str(g), str(b), str(a), str(hAlign), str(vAlign), str(fontString)]
            return dizionarioCelle

    # rc = rowCol
    def rcDictFromFileString(self, rcStringa:str) -> dict:
        """Prendo i valori dalla stringa fornita e li separo per metterli poi nel dizionario"""
        if rcStringa != "vuoto\n":
            dizionario = {}
            listaRowCol = rcStringa.split("\n")
            if listaRowCol[-1] == "":
                listaRowCol.pop(-1)
            for rowCol in listaRowCol:
                listaRc = rowCol.split(", ")
                dizionario[listaRc[0]] = listaRc[1]
            return dizionario
        return
    
    def aggiornaStatusBarSalvataggio(self, evt):
        if self.deviSalvare:
            self.statusBar.SetStatusText("Non salvato")
        else:
            self.statusBar.SetStatusText("Salvato")
        
    def salva(self, stringaFile = "", copia = False, saveAs = False): #Se gli passi il contenuto del file(di default "") puoi usarla anche per salva normale
        # Celle
        #      row, col, cont, textColour, Colour, alignment
        #      Font
        # Col
        #      width, Format
        # Row
        #      height
        
        # Quando salvi su un file già esistente controlla anche se la cella è già salvata da qualche parte
        # Formato .xlrsg
        
        lastPercorso = self.percorso
        
        if saveAs or self.percorso == TITOLO_INIZIALE:
            dlg = wx.FileDialog(None, "Salva File", style=wx.FD_SAVE, wildcard="RsgCel files (*.xlrsg)|*.xlrsg")
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
            colonneModificate = self.rcDictFromFileString(listaStringaFile[2])
            stringaFile = ""
            
        baseHeight = self.mainGrid.GetDefaultRowSize()
        baseWidth = self.mainGrid.GetDefaultColSize()
        baseCont = ""
        (baseTr, baseTg, baseTb, baseTa) = self.mainGrid.GetDefaultCellTextColour()
        (baseR, baseG, baseB, baseA) = self.mainGrid.GetDefaultCellBackgroundColour().Get()
        (baseHAlign, baseVAlign) = self.mainGrid.GetDefaultCellAlignment()
        fontString = self.mainGrid.GetDefaultCellFont().GetNativeFontInfoDesc()
        
        baseCell = [str(baseCont), str(baseTr), str(baseTg), str(baseTb), str(baseTa), str(baseR), str(baseG), str(baseB), str(baseA), str(baseHAlign), str(baseVAlign), str(fontString)]
            
        
        for row in range(self.mainGrid.GetNumberRows()):
            for col in range(self.mainGrid.GetNumberCols()):
                cont = self.mainGrid.GetCellValue(row, col)
                
                cella = (row, col)
                for a in self.somma:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.somma[a], "+")
                        break
                
                for a in self.sottrazione:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.sottrazione[a], "-")
                        break
                
                for a in self.divisione:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.divisione[a], "/")
                        break
                
                for a in self.moltiplicazione:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.moltiplicazione[a], "*")
                        break
                
                for a in self.media:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.media[a], "MEDIA")
                        break
                
                for a in self.massimo:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.massimo[a], "MAX")
                        break
                
                for a in self.minimo:
                    if a == cella:
                        cont = self.creaContenutoOperazione(self.minimo[a], "MIN")
                        break
                
                (tr, tg, tb, ta) = self.mainGrid.GetCellTextColour(row, col).Get()
                (r, g, b, a) = self.mainGrid.GetCellBackgroundColour(row, col).Get()
                (hAlign, vAlign) = self.mainGrid.GetCellAlignment(row, col)
                fontString = self.mainGrid.GetCellFont(row, col).GetNativeFontInfoDesc()
                
                cell = [str(cont), str(tr), str(tg), str(tb), str(ta), str(r), str(g), str(b), str(a), str(hAlign), str(vAlign), str(fontString)]
                
                if cell != baseCell:
                    celleModificate[(row, col)] = cell
                
                width = self.mainGrid.GetColSize(col)
                if width != baseWidth:
                    colonneModificate[col] = str(width)
            
            height = self.mainGrid.GetRowSize(row)
            if height != baseHeight:
                righeModificate[row] = str(height)
        
        stringaFile += self.dictToString(celleModificate) or "vuoto\n"
        
        stringaFile += "Separatore\n"
        
        stringaFile += self.dictToString(righeModificate) or "vuoto\n"
        
        stringaFile += "Separatore\n"
        
        stringaFile += self.dictToString(colonneModificate) or "vuoto\n"
        
        if EXTENSION in self.percorso:
            fileName = self.percorso
        else:
            fileName = self.percorso + EXTENSION
        
        file = open(fileName, "w")
        file.write(stringaFile)
        file.close()
        
        self.deviSalvare = False
        self.statusBar.SetStatusText("SALVATAGGIO COMPLETATO")
        self.timer.StartOnce(5000)
        
        if copia:
            self.percorso =lastPercorso
        
        self.SetTitle(self.percorso + " - " + APP_NAME)
        
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
    
        file = open(self.percorso, "r")
        contenuto = file.read()
        file.close()
        
        [celle, righe, colonne] = contenuto.split("Separatore\n")
        
        if celle != "vuoto\n":
        
            dizionarioCelle = self.cellDictFromFileString(celle)
            
            for cella in dizionarioCelle:
                row = int(cella[0])
                col = int(cella[1])
                cont = dizionarioCelle[cella][0]
                tr = int(dizionarioCelle[cella][1])
                tg = int(dizionarioCelle[cella][2])
                tb = int(dizionarioCelle[cella][3])
                ta = int(dizionarioCelle[cella][4])
                r = int(dizionarioCelle[cella][5])
                g = int(dizionarioCelle[cella][6])
                b = int(dizionarioCelle[cella][7])
                a = int(dizionarioCelle[cella][8])
                hAlign = int(dizionarioCelle[cella][9])
                vAlign = int(dizionarioCelle[cella][10])
                fontString = dizionarioCelle[cella][11]
                 
                alignment = (hAlign, vAlign)
                
                self.mainGrid.SetCellValue(row, col, cont)
                self.mainGrid.SetCellTextColour(row, col, wx.Colour(tr, tg, tb, ta))
                self.mainGrid.SetCellBackgroundColour(row, col, wx.Colour(r, g, b, a))
                self.mainGrid.SetCellAlignment(row, col, hAlign, vAlign)
                
            for cella in dizionarioCelle:
                cont = dizionarioCelle[cella][0]
                if "=" in cont:
                    row = int(cella[0])
                    col = int(cella[1])
                    self.operazioni(row, col)
        
        if righe != "vuoto\n":
        
            dizionarioRighe = self.rcDictFromFileString(righe)
            
            for riga in dizionarioRighe:
                row = int(riga[0])
                height = int(dizionarioRighe[riga])
                
                self.mainGrid.SetRowSize(row, height)
        
        if colonne != "vuoto\n":
        
            dizionarioColonne = self.rcDictFromFileString(colonne)
            
            for colonna in dizionarioColonne:
                row = int(colonna[0])
                width = int(dizionarioColonne[colonna])
                
                self.mainGrid.SetColSize(row, width)
        
        self.deviSalvare = False
        self.statusBar.SetStatusText("Salvato")
        self.SetTitle(self.percorso + " - " + APP_NAME)
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
        self.mainGrid.Refresh()
        return
    
    def funzioneSalva(self, evt):
        stringaFile = ""
        if self.percorso != TITOLO_INIZIALE:
            f = open(self.percorso)
            stringaFile = f.read()
            f.close()
        self.salva(stringaFile)
        return

    def funzioneSalvaConNome(self, evt):
        self.salva("", saveAs = True)
        return
    
    def funzioneSalvaCopia(self, evt):
        self.salva("", copia = True, saveAs = True)
        return
    
    def funzioneStampa(self, evt):
        printout = GridPrintout(self.mainGrid, "Stampa")
        printer = wx.Printer()
        #prompt True così mostra la finestra di dialogo nella stampa
        printer.Print(self, printout, prompt=True)
    
    
    def funzioneProprietà(self, evt):
        proprieta = "Nome file: " + "\nTipo: Foglio elettronico" + "\nPosizione:" + "\nDimensione: sconosciuto" + "\nCreato:"
        
        dial = wx.MessageDialog(None, proprieta, "Proprietà", wx.OK | wx.ICON_INFORMATION)
        dial.ShowModal()
        
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
        row = self.mainGrid.GetGridCursorRow()
        col = self.mainGrid.GetGridCursorCol()
        self.mainGrid.SelectBlock(row, col, row, col, False)
        return
    
    def funzioneTrova(self, evt):
        trovadial = wx.TextEntryDialog(None, "Cosa cerchi?", "Trova")
        trovadial.ShowModal()
        
        return
    
    def funzioneTrovaeSostituisci(self, evt):
        trovaEsosdial = wx.FindReplaceDialog(None, "Trova e Sostituisci") # Manca qualcosa (data)
        trovaEsosdial.ShowModal()
        
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
        dlg = wx.FileDialog(None, "Apri File", style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_CANCEL:
            return

        percorso = dlg.GetPath()
        
        # Capire come inserire immagine nella griglia
        return
    
    def funzioneCollegamento(self, evt):
        dialog=wx.TextEntryDialog(None, "inserisci l'URL per il collegamento", "Domanda", "")
        if dialog.ShowModal() == wx.ID_OK:
            url = dialog.GetValue()
            open_webpage(url)
        dialog.Destroy()
        return
    
    def funzioneData(self, evt):
        rows = self.mainGrid.GetNumberRows()
        cols = self.mainGrid.GetNumberCols()
        oggi = datetime.date.today()
        data = oggi.strftime("%d/%m/%Y")
        for row in range(rows):
            for col in range(cols):
                if self.mainGrid.IsInSelection(row, col):
                    self.mainGrid.SetCellValue(row, col,data)
        return
    
    def funzioneOra(self, evt):
        rows = self.mainGrid.GetNumberRows()
        cols = self.mainGrid.GetNumberCols()
        adesso = datetime.datetime.now()
        orario = adesso.strftime("%H:%M:%S")
        for row in range(rows):
            for col in range(cols):
                if self.mainGrid.IsInSelection(row, col):
                    self.mainGrid.SetCellValue(row, col, orario)
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
        datiIniziali = wx.FontData()
        dialog = wx.FontDialog(self, datiIniziali)
        if dialog.ShowModal() != wx.ID_OK:
            return

        datiFinali = dialog.GetFontData()
        fontSelezionato = datiFinali.GetChosenFont()
       
        self.mainGrid.SetFont(fontSelezionato)
        return
    
    #Funzioni Foglio
    def funzionePulisciCelle(self,evt):
        self.mainGrid.ClearGrid()
        return
    def funzioneRinomina(self,evt):
        dlg = wx.TextEntryDialog(self, "nuovo titolo:", "inserisci titolo", self.GetTitle())
        if dlg.ShowModal() == wx.ID_OK:
            titolo = dlg.GetValue()
            self.SetTitle( titolo + " - " + APP_NAME)
        return
    
    
    #Funzioni Dati
    def funzioneOrdinaCresc(self,evt):
        rows = self.mainGrid.GetNumberRows()
        cols = self.mainGrid.GetNumberCols()
        selected_words = []
        for row in range(rows):
            for col in range(cols):
                word = self.mainGrid.GetCellValue(row, col)
                if self.mainGrid.IsInSelection(row, col):
                    selected_words.append((word, row, col))
        selected_words.sort()# Ordino le parole selezionate in ordine crescent
        index = 0
        for row in range(rows):
            for col in range(cols):
                if self.mainGrid.IsInSelection(row, col):
                    self.mainGrid.SetCellValue(row, col, selected_words[index][0])
                    index += 1
        return
                           
    def funzioneOrdinaDecr(self,evt):
        rows = self.mainGrid.GetNumberRows()
        cols = self.mainGrid.GetNumberCols()
        selected_words = []
        for row in range(rows):
            for col in range(cols):
                word = self.mainGrid.GetCellValue(row, col)
                if self.mainGrid.IsInSelection(row, col):
                    selected_words.append((word, row, col))
        selected_words.sort()  # Ordino le parole selezionate in ordine crescente
        selected_words.reverse()  # Inverto l'ordine per ottenere l'ordinamento decrescente
        index = 0
        for row in range(rows):
            for col in range(cols):
                if self.mainGrid.IsInSelection(row, col):
                    self.mainGrid.SetCellValue(row, col, selected_words[index][0])
                    index += 1
        return
    
    #Funzioni Strumenti
    def funzioneCheckOrto(self,evt):
        return
 
    # Funzioni Aiuto
    def funzioneInfoLic(self, evt):
        dial = wx.MessageDialog(None, "Licenza: Open Source,programma disponibile per tutti", "Informazione Licenza", wx.OK | wx.ICON_INFORMATION)
        dial.ShowModal()
        return
    
    def funzioneInfoRSG(self, evt):
        dial = wx.MessageDialog(None, "RSG Cel\nVersione: 1.0.0\nSviluppatori: Gramazio Rocco,Ristè Thatiely,Solfanelli Davide\nRSG Cel è un programma per fare fogli di calcolo", "Informazione RSG cel", wx.OK | wx.ICON_INFORMATION)
        dial.ShowModal()
        return
    
    def funzioneAiuto(self, evt):
        open_webpage("https://ti-aiuto.it/")
        return
    
    def funzioneDocumentazione(self, evt):
        open_webpage("https://support.microsoft.com/it-it/excel")
        return
    
    def funzioneDonazione(self, evt):
        dial = wx.MessageDialog(None, "Sei sicuro di voler darci dei soldi?", "Domanda", wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION)
        risposta = dial.ShowModal()
        if risposta == wx.ID_YES:
            open_webpage("file:///C:/Users/Giovanni/Desktop/Nuova%20cartella/sito.html")
        elif risposta == wx.ID_NO:
            return
            
        elif risposta == wx.ID_CANCEL:
            return
        return
    
    def funzioneScegliCarattere(self, evt):
        datiIniziali = wx.FontData()
        dialog = wx.FontDialog(self, datiIniziali)
        if dialog.ShowModal() != wx.ID_OK:
            return

        datiFinali = dialog.GetFontData()
        fontSelezionato = datiFinali.GetChosenFont()

        self.mainGrid.SetFont(fontSelezionato) # Vedere come impostare il font nella griglia

        return
        
#la classe ereditata da wx.Printout definisce il contenuto della stampa del foglio di calcolo
class GridPrintout(wx.Printout):
    def __init__(self, grid, title):
        super().__init__(title)
        self.mainGrid = grid

#a inizio stampa
    def OnPreparePrinting(self):
        self.page_conta = 1

#FUNZIONE chiamata per ogni pagina
    def OnPrintPage(self, page):
        #disegno la griglia -> poi utilizzo DrawGrid
        dc = self.GetDC()
        if dc is not None:
            dc.SetFont(self.grid.GetDefaultCellFont())
            dc.SetPen(wx.BLACK_PEN)
            dc.SetBrush(wx.WHITE_BRUSH)
            dc.Clear()

            pageSize = self.GetPageSizePixels()
            pageRect = self.GetPageRectPixels()

            # Calcola le dimensioni della griglia in base alle dimensioni della pagina di stampa
            nR, nC = self.grid.GetNumberRows(), self.grid.GetNumberCols()
            total_width = sum([self.grid.GetColSize(col) for col in range(nC)])
            total_height = sum([self.grid.GetRowSize(row) for row in range(nR)])

            # Calcola la posizione di partenza per disegnare la griglia
            start_x = pageRect.x + ((page * pageSize.x) // pageSize.GetWidth())
            start_y = pageRect.y + ((page * pageSize.y) // pageSize.GetHeight())

            # Disegna la griglia
            self.grid.DrawGrid(start_x, start_y, total_width, total_height, dc)

            return True

        return False


# ----------------------------------------
if __name__ == "__main__":
    app = wx.App()
    app.SetAppName(APP_NAME)
    window = Finestra()
    window.Show()
    app.MainLoop()
