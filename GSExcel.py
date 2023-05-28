#Gramazio Rocco & Solfanelli Davide
#4AS
#EsFinale

import wx

APP_NAME = "GS Excel"

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
ID_Info = 19
ID_Img = 20
ID_Funzione = 21
ID_Collegamento = 22
ID_CarSpeciale = 23
ID_Data = 24
ID_Ora = 25
ID_Testo = 26
ID_Spaziatura = 27
ID_Allinea = 28


class Finestra(wx.Frame):

    def __init__(self):
        super().__init__(None, title=APP_NAME)

        # Spazio per le variabili membro della classe

        # -------------------------------------------

        # Chiamata alle funzioni che generano la UI
        self.creaMenubar()
        self.creaToolbar()
        self.creaStatusbar()
        # -------------------------------------------

        # Chiamata alla funzione che genera la MainView
        self.creaMainView()
        # -------------------------------------------

        # le ultime cose, ad esempio, le impostazioni iniziali, etc...

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
        
        fileMenu.Append(wx.ID_NEW, "Nuovo")
        fileMenu.Append(wx.ID_OPEN, "Apri")
        fileMenu.Append(customItemDOCRec)
        fileMenu.Append(wx.ID_CLOSE, "Chiudi")
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_REFRESH, "Ricarica")
        fileMenu.AppendSeparator()
        fileMenu.Append(wx.ID_SAVE, "Salva")
        fileMenu.Append(customItemSalvaNome)
        fileMenu.Append(customItemDOCRec)
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
        
        # crea Menu Aiuto
        helpMenu = wx.Menu()
        
        # Creazione Item menu Aiuto
        customItemInfo = wx.MenuItem(helpMenu, 19, "Informazioni su RsgCel")
        
        helpMenu.Append(customItemInfo)
        
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
        
        # Bind Help
        self.Bind(wx.EVT_MENU, self.funzioneInfo, id=ID_Info)

        return

    # in questa funzione andremo a creare e popolare la toolbar
    def creaToolbar(self):
        return

    # in questa funzione aggiungeremo la statusbar
    def creaStatusbar(self):
        return

    # questa funzione implementa la vista principale del programma
    def creaMainView(self):
        
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

    def funzioneSalva(self, evt):
        return

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
    
    def funzioneEsci(self, evt):
        return
    
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
    
    # Funzioni Aiuto
    def funzioneInfo(self, evt):
        return
    
# ----------------------------------------
if __name__ == "__main__":
    app = wx.App()
    app.SetAppName(APP_NAME)
    window = Finestra()
    window.Show()
    app.MainLoop()
