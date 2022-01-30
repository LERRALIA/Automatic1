VERSION 5.00
Begin VB.Form frmZEN42 
   Caption         =   " - Kassenprotokolle"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command5 
      Caption         =   "Löschen"
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
      Index           =   2
      Left            =   9600
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "frmZEN42.frx":0000
      Left            =   120
      List            =   "frmZEN42.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   8760
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ansehen"
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
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Schließen"
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
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   $"frmZEN42.frx":0004
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kassenprotokolle"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmZEN42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

  

    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Kassprot"

    Select Case Index
        Case 0
            Unload frmZEN42
        Case 1 'ansehen
            If List1.ListIndex >= 0 Then
                Anzeige "normal", "", Label1(4)
                Screen.MousePointer = 11
                zeigeHilfeDabapfad "KassProt", List1.list(List1.ListIndex)
                Screen.MousePointer = 0
            Else
                Anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label1(4)
            End If
        Case 2
            If List1.ListIndex >= 0 Then
                Anzeige "normal", "", Label1(4)
                Kill cPfad & "\" & List1.list(List1.ListIndex)
                fuelleliste
            Else
                Anzeige "rot", "Wählen Sie bitte eine Datei aus!", Label1(4)
            End If
            
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    
    fuelleliste
    

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
'    Anzeige "normal", "", Label1(4)
    
    

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub fuelleliste()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    Dim lCount  As Long
    Dim cExt    As String
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "KASSPROT"
    
    File2.Path = cPfad
    File2.Pattern = "*.txt"
    File2.Refresh
    
    List1.Clear
    For lCount = 0 To File2.ListCount - 1
        ctmp = File2.list(lCount)
        ctmp = Trim$(ctmp)
        ctmp = UCase$(ctmp)
        
        cExt = Right$(ctmp, 3)
        If UCase$(cExt) = "TXT" Then
            List1.AddItem ctmp
            
        End If
    Next lCount
    If List1.ListCount > 1 Then
        Anzeige "normal", List1.ListCount & " Dateien stehen zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 1 Then
        Anzeige "normal", "1 Datei steht zur Verfügung", Label1(4)
    ElseIf List1.ListCount = 0 Then
        Anzeige "normal", "Es steht keine Datei zur Verfügung", Label1(4)
    End If
    
    List1.Refresh
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelleliste"
    Fehler.gsFehlertext = "Im Programmteil Kassenprotokolle ist ein Fehler aufgetreten."
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
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
