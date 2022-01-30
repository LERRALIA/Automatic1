VERSION 5.00
Begin VB.Form frmWK21p 
   BackColor       =   &H000000C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         BackColor       =   &H000000C0&
         Caption         =   "Stimmt nicht, an keiner Kasse wird ein Tagesabschluss durchgeführt"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         Caption         =   "Die statistischen Daten sind erstellt . Bitte übertragen Sie nach dem Kassen- abschluss die Daten per Telekiss."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmWK21p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()    'sofort weiter und löschen
    On Error GoTo LOKAL_ERROR

    Dim cPfad23 As String
    Dim j As Integer
    Dim i As Integer
    
    j = 0
    cPfad23 = gcDBPfad               'Datenbankpfad
    If Right(cPfad23, 1) <> "\" Then
        cPfad23 = cPfad23 & "\"
    End If
    
    Kill cPfad23 & Label1.Caption & ".txt"
    
    Unload Me
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 0
    lbl6(0).Caption = gsAnzeigeText
    Timer1.Interval = 2000
    Timer1.Enabled = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad23 As String
    
    cPfad23 = gcDBPfad               'Datenbankpfad
    If Right(cPfad23, 1) <> "\" Then
        cPfad23 = cPfad23 & "\"
    End If
    
    anzeige "artikel", "", Label2
    
    Command1.Visible = True
    
    If Not FileExists(cPfad23 & Label1.Caption & ".txt") Then
        Unload Me
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

