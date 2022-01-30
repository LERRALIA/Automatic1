VERSION 5.00
Begin VB.Form frmWKL132 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Ergänzungen zur Mahnung"
   ClientHeight    =   6120
   ClientLeft      =   2055
   ClientTop       =   1770
   ClientWidth     =   10845
   Icon            =   "frmWKL132.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6120
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schließen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6720
      TabIndex        =   14
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Anschrift im Rechnungsfuß"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   4215
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Löschen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Auswählen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   7680
      MaxLength       =   35
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Drucken"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   8760
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Leeren"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   4
      Top             =   4560
      Width           =   1935
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Speichern"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   4680
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   9735
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "vorhandene Texte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Kurzbeschreibung:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Mahnungstexte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmWKL132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim iRet As Integer
    
    Select Case Index
        Case Is = 0     'speichern
            If Text1(2).Text = "" Then
                Text1(2).Text = Left(Text1(1).Text, 35)
                Text1(2).Text = SwapStr(Text1(2).Text, Chr(10), "")
                Text1(2).Text = SwapStr(Text1(2).Text, Chr(13), "")
            End If
            
            Speicherblock "MAHNUNG", Text1(1).Text, Text1(2).Text
            ZeigeblockinList "MAHNUNG", List1
            
        Case Is = 1     'leeren
'            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""

        Case Is = 2 'Übernehmen
        
            ctmp = Text1(1).Text
            ctmp = Trim$(ctmp)

            frmWKL24!Label5(1).Caption = ctmp

            If Check1.Value = vbChecked Then
                iRet = vbYes
            Else
                iRet = vbNo
            End If
            frmWKL24.SchreibeMahnung iRet
        Case 5
            
            Unload frmWKL132
        Case Is = 3     'auswählen
            If List1.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag auswählen!", Label1(5)
                Exit Sub
            End If
            ZeigeblockinEinzelteile "MAHNUNG", List1, Text1(2), Text1(1), CLng(Right$(List1.list(List1.ListIndex), 10))
        Case Is = 4     'löschen
            If List1.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag auswählen!", Label1(5)
                Exit Sub
            End If
            DELblock CLng(Right$(List1.list(List1.ListIndex), 10))
            ZeigeblockinList "MAHNUNG", List1
            anzeige "normal", "Eintrag wurde gelöscht", Label1(5)
            Text1(2).Text = ""
            Text1(1).Text = ""
            
    End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Label1(0)
    
'    Text1(0).Text = "" 'frmWKL24!Label5(0).Caption
    Text1(1).Text = "" 'frmWKL24!Label5(1).Caption
    Text1(2).Text = ""
    
    ZeigeblockinList "MAHNUNG", List1
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
LogtoEnd Me
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List1_Click()
On Error GoTo LOKAL_ERROR

    ZeigeblockinEinzelteile "MAHNUNG", List1, Text1(2), Text1(1), CLng(Right$(List1.list(List1.ListIndex), 10))
              
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    cZeichen = Chr$(KeyAscii)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 0 'Porto
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text1_LostFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    Text1(Index).BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



