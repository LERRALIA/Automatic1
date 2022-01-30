VERSION 5.00
Begin VB.Form FTPwechsel 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "FTPumzug"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtPasw 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtEdekaNr 
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Passwort :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "neue EDEKA-KundenNr :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   " geben sie diese ein, und drücken Sie [Ja]  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   " erhalten haben,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "KundenNr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "wenn Sie schon Ihre neue EDEKA - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Umstellung auf den BHSG- Bestellvorgang "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "zum neuen FTP-Verfahren wechseln ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Budni :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FTPwechsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 If Trim(txtEdekaNr.Text) = "" Then
 
  MsgBox ("Bitte geben Sie Ihre neue EDEKA-Nummer ein und versuchen Sie es erneut !")
  
 ElseIf txtPasw.Text = "brnas2030" Then
   
  Dim edkNr As String
  Dim lWert As Long
  edkNr = txtEdekaNr.Text

  lWert = MsgBox("ist Ihre EDEKA-Nummer : [ " & edkNr & " ] richtig ?", vbYesNo + vbQuestion, "Winkiss Frage:")
  If lWert = vbYes Then

        'Budni wird auf EDEKA-Nr geändert
        gdBase.Execute ("UPDATE LISRT SET KUNDNR='" & edkNr & "',GLN='" & edkNr & "',FORMAT='EDIBHSG' WHERE FORMAT='EDIBUDNI'")

        'Artikel libesnr Aktualisieren
        If BudniArtikelNrAktualisieren Then

         gdBase.Execute "Create Table FTPumzugFertig(EsIstFertig bit)", dbFailOnError
         Label3.Caption = "Fertig"
         Label3.Refresh
         gbBudniNeuesFtpVerfahren = True
         Unload Me

        End If

  End If
   
 Else
 
 MsgBox ("falsches Passwort ! ! !")
 
 End If
 
 

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Activate()

 Me.BackColor = glH1
 Label1.BackColor = glH1
 Label2.BackColor = glH1
 Label3.BackColor = glH1
 Label4.BackColor = glH1
 Label5.BackColor = glH1
 Label6.BackColor = glH1
 Label7.BackColor = glH1
 Label8.BackColor = glH1
 Label9.BackColor = glH1
 Label10.BackColor = glH1
 
 MsgBox ("Achtung: " & vbNewLine & "Sie dürfen diese Umstellung durchführen, wenn Sie von Dronova hierfür aufgefordert werden. ansonsten führen Sie die Umstellung bitte nicht durch.")
 
End Sub

Private Function BudniArtikelNrAktualisieren() As Boolean
On Error GoTo LOKAL_ERROR

 BudniArtikelNrAktualisieren = False
 
 Label3.Caption = " [bestellnummern] werden aktualisiert . . ."
 Label3.Refresh
 
 Dim neuArtikelBudniPfad As String
 neuArtikelBudniPfad = gcDBPfad & "\neuBudniArtikel.mdb"
 gdBase.Execute ("UPDATE Artikel A INNER JOIN [MS Access;Database=" & neuArtikelBudniPfad & "].neuBudniArtikelNr NA ON A.EAN=NA.[GTIN-Code] SET A.libesnr=NA.[EDK Artik]")

 BudniArtikelNrAktualisieren = True

Exit Function
LOKAL_ERROR:

    BudniArtikelNrAktualisieren = False
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "FTPwechsel"
    Fehler.gsFunktion = "BudniArtikelNrAktualisieren"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

 
