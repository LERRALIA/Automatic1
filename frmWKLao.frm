VERSION 5.00
Begin VB.Form frmWKLao 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Ihre Telefonnummer? Wir möchten Sie zurückrufen."
   ClientHeight    =   3240
   ClientLeft      =   2055
   ClientTop       =   2865
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3240
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin sevCommand3.Command Command2 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Speichern"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5040
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Leer"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "0"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   5520
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "9"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "8"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   4320
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "7"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "6"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "5"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "4"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "3"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "2"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "1"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaxLength       =   18
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWKLao.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmWKLao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 To 9
            Text1.Text = Text1.Text & Command1(Index).Caption
        Case 10
            Text1.Text = ""
    End Select

    Text1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    Dim iRet As Integer
    Dim sSQL As String
    
    Select Case Index
        Case Is = 0     'Speichern
            iRet = fnPruefeTelefonNr()
            Select Case iRet
                Case Is = 0     'alles okay
                
                    sSQL = "Update FIRMA SET TEL = '" & Text1.Text & "'"
                    gdBase.Execute sSQL, dbFailOnError
                        
                    LeseFirmenDaten
                    Unload frmWKLao
                    
                Case Is = 1     'keine Eingabe
                    Screen.MousePointer = 11
                    MsgBox "Bitte geben Sie Ihre Rückrufnummer ein!", vbInformation, "Winkiss Hinweis:"
                    Text1.SetFocus
            End Select
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeTelefonNr() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    
    fnPruefeTelefonNr = 0
    
    ctmp = Text1.Text
    ctmp = Trim$(ctmp)
    
    If ctmp = "" Then
        fnPruefeTelefonNr = 1
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeTelefonNr"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    frmWKLao.Top = Screen.Height / 2 - frmWKLao.Height / 2
    frmWKLao.Left = Screen.Width / 2 - frmWKLao.Width / 2
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command2_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rückrufnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



