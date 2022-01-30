VERSION 5.00
Begin VB.Form frmWK21c 
   Caption         =   "Verbinden mit ?"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3045
   ControlBox      =   0   'False
   Icon            =   "frmWK21c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3045
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command2 
      Caption         =   "Trennen"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verbinden"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmWK21c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Me.Hide
    
    If List1.ListIndex = -1 Then Exit Sub
    
    InternetDial Me.hwnd, ConName, DIAL_FORCE_UNATTENDED, ConID, 0

    gbVerbindungstarten = True
    Unload frmWK21c

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil 'Verbindung mit?' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    If ConID Then InternetHangUp ConID, 0
    ConID = 0
    
    gbVerbindungstarten = False
    Unload frmWK21c
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil 'Verbindung mit?' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim S&, LN&, X%
    Dim R(255) As RASENTRYNAME95

    Screen.MousePointer = 11
    '### Namen der bestehenden DFÜ-Verbindungen einlesen
    R(0).dwSize = 264
    S = 256 * R(0).dwSize
    Call RasEnumEntries(vbNullString, vbNullString, R(0), S, LN)
    
    If LN <> 0 Then
      '### Es besteht mindestens eine DFÜ-Verbindung
      For X = 0 To LN - 1
        ConName = StrConv(R(X).szEntryName(), vbUnicode)
        
        If Trim(Left$(ConName, InStr(ConName, vbNullChar) - 1)) = "Esüdro" Then
            List1.AddItem Left$(ConName, InStr(ConName, vbNullChar) - 1)
        End If
      Next X
      List1.ListIndex = 0
      
    Else
      '### Keine DFÜ da
        Screen.MousePointer = 0
      
       frmWKL37.Label16.Caption = "Keine DFÜ mit dem Namen 'ESÜDRO' vorhanden."
       frmWKL37.Label16.Refresh
      
        Command1.Enabled = False
        Command2.Enabled = False
    End If
    
    gbVerbindungstarten = False
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil 'Verbindung mit?' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List1_Click()
    On Error GoTo LOKAL_ERROR
    
    ConName = List1.list(List1.ListIndex)
    DFÜname = ConName
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil 'Verbindung mit?' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
