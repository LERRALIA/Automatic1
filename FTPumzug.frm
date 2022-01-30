VERSION 5.00
Begin VB.Form FTPumzug 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "FTPumzug"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Nein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Budni :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "das neue FTP-Verfahren übernehmen ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   6255
   End
End
Attribute VB_Name = "FTPumzug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR:
  
 If BudniArtikelnummerAktualisieren Then
 
  'create sperrtabelle [FTPumzugFertig]
   gdBase.Execute "Create Table FTPumzugFertig(EsIstFertig bit)", dbFailOnError
   gbBudniNeuesFtpVerfahren = True
   Label3.Caption = "Fertig"
   Label3.Refresh
   Unload Me
 End If
 
 
 

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FTPumzug"
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = ""
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()

 Me.BackColor = glH1
 Label1.BackColor = glH1
 Label2.BackColor = glH1
 Label3.BackColor = glH1
End Sub


Private Function BudniArtikelnummerAktualisieren() As Boolean
On Error GoTo LOKAL_ERROR:
 
 Dim neuBudniArtikelPfad As String
 neuBudniArtikelPfad = gcDBPfad & "\neuBudniArtikel.MDB"
 
 Label3.Caption = "Budni-ArtikelNr werden aktualisiert . . ."
 Label3.Refresh
 
 gdBase.Execute ("UPDATE ARTIKEL A INNER JOIN [MS Access;Database=" & neuBudniArtikelPfad & "].neuBudniArtikelNr NA ON A.EAN=NA.[GTIN-Code] SET A.LIBESNR=NA.[EDK Artik]")
   
 BudniArtikelnummerAktualisieren = True

Exit Function
LOKAL_ERROR:
    
    BudniArtikelnummerAktualisieren = False
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FTPumzug"
    Fehler.gsFunktion = "BudniArtikelnummerAktualisieren"
    Fehler.gsFehlertext = ""
    
    Fehlermeldung1


End Function

