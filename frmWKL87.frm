VERSION 5.00
Begin VB.Form frmWKL87 
   Caption         =   "Farbmerkmale der Bestellvorschlagszahlen"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   Icon            =   "frmWKL87.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10110
   StartUpPosition =   2  'Bildschirmmitte
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
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   18
      Top             =   3600
      Width           =   1695
   End
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
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Tag             =   "Shape"
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundenbestellungen + Konditionsvorgaben + Inhaltsvorgabe (Packungsgröße) wurden berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   16
      Top             =   3240
      Width           =   9495
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Tag             =   "Shape"
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundenbestellung + Inhaltsvorgabe (Packungsgröße) wurden berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   14
      Top             =   2880
      Width           =   9495
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Tag             =   "Shape"
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Konditionsvorgaben + Inhaltsvorgabe (Packungsgröße) wurden berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Top             =   2520
      Width           =   9495
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Tag             =   "Shape"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundenbestellungen + Konditionsvorgaben wurden berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Konditionsvorgaben wurden berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Tag             =   "Shape"
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Inhaltsvorgabe (Packungsgröße) wurde berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Tag             =   "Shape"
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "ohne Merkmal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Kundenbestellung wurde berücksichtigt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Tag             =   "Shape"
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '2D
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "Shape"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Farbmerkmale der Bestellvorschlagszahlen"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmWKL87"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 1
            Unload frmWKL87
        Case 0
            drucken
    End Select
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmale auf. "
    Fehlermeldung1
End Sub
Private Sub drucken()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim sSQL As String
    Dim lfarbwert As Long
    Dim cfbBE As String
    
    
    loeschNEW "BVOFARBE", gdBase
    CreateTable "BVOFARBE", gdBase
    
    For i = 0 To 7
    
        lfarbwert = CDec(Label5(i).BackColor)
        cfbBE = Label4(i).Caption
        
        sSQL = "Insert into BVOFARBE (FARBWERT,FARBBESCHREIB) values "
        sSQL = sSQL & " ( " & lfarbwert & ",'" & cfbBE & "')"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    
    Next i
    
    reportbildschirm "", "aWKL87"
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmale auf. "
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Modul6.alternativFarbform Me, Label1
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Farbmerkmale auf. "
    Fehlermeldung1
End Sub
