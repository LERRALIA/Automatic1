VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL217 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   8145
   ClientLeft      =   2445
   ClientTop       =   2010
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8145
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command15 
      Height          =   735
      Left            =   6480
      TabIndex        =   0
      Top             =   7200
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "jetzt per SMS erinnern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "jetzt per SMS erinnern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "jetzt per SMS erinnern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "jetzt per SMS erinnern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   8
      Top             =   5160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "jetzt per SMS erinnern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Übersicht der Erinnerungen verwerfen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      MouseIcon       =   "frmWKL217.frx":0000
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   26
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Sie haben jederzeit die Möglichkeit für einzelne Tage die SMS - Erinnerung über den Terminkalender vorzunehmen."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   6480
      Width           =   9135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Caption         =   "Heute"
      BeginProperty Font 
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
      TabIndex        =   24
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Caption         =   "Heute"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Caption         =   "Heute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Caption         =   "Heute"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      Caption         =   "Heute"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "erinnert am:"
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
      Index           =   3
      Left            =   6240
      TabIndex        =   19
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Behandlungstag:"
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
      Index           =   2
      Left            =   600
      TabIndex        =   18
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "erinnert am?"
      BeginProperty Font 
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
      Left            =   6240
      TabIndex        =   17
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "erinnert am?"
      BeginProperty Font 
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
      Left            =   6240
      TabIndex        =   16
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "erinnert am?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "erinnert am?"
      BeginProperty Font 
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
      Left            =   6240
      TabIndex        =   14
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Erinnerung ausstehend"
      BeginProperty Font 
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
      Left            =   6240
      TabIndex        =   13
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Behandlungsdatum"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Behandlungsdatum"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Behandlungsdatum"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Behandlungsdatum"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Termine dieser Behandlungstage per SMS erinnern."
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9135
   End
   Begin VB.Label Label2 
      Caption         =   "Behandlungsdatum"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00808000&
      Caption         =   "SMS Erinnerung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "frmWKL217"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim dateAuswertungstag      As Date
    Dim sMess                   As String

    
    
    dateAuswertungstag = DateValue(Label2(Index).Caption)
    
    sMess = "Möchten Sie jetzt die Termine für den " & dateAuswertungstag & " per SMS versenden?"

    Dim iRet As Integer

    iRet = MsgBox(sMess, vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        LeseOpeningsWKL82
        VersendeTermineSMS dateAuswertungstag
        TrageSMSBenachrichtigungEin dateAuswertungstag
        WasBisherGeschah
    End If
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command15_Click()
On Error GoTo LOKAL_ERROR
    
    Unload frmWKL217

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command15_Click"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
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

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    WasBisherGeschah

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WasBisherGeschah()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim rsRs2 As DAO.Recordset
    Dim lBehDat                 As Long
    Dim lWeekday                As Long
    
    Screen.MousePointer = 11
    
    If NewTableSuchenDBKombi("SMS_UEBERSICHT", gdBase) = False Then
        'dann heute
        'heute + 1
        'heute + 2
        'anbieten
        
        
        Label2(0).Visible = True
        Label3(0).Visible = True
        Label3(0).Caption = ""
        Label4(0).Visible = True
        Command1(0).Visible = True
        
        Command1(0).Enabled = True
        Command1(1).Enabled = True
        Command1(2).Enabled = True
        Command1(3).Enabled = True
        Command1(4).Enabled = True
        
        
        
        Label2(0).Caption = DateValue(Now)
        Label2(1).Caption = DateAdd("d", 1, DateValue(Now))
        Label2(2).Caption = DateAdd("d", 2, DateValue(Now))
        Label2(3).Caption = DateAdd("d", 3, DateValue(Now))
        Label2(4).Caption = DateAdd("d", 4, DateValue(Now))
        
        Label3(0).Caption = "Erinnerung ausstehend"
        Label3(1).Caption = "Erinnerung ausstehend"
        Label3(2).Caption = "Erinnerung ausstehend"
        Label3(3).Caption = "Erinnerung ausstehend"
        Label3(4).Caption = "Erinnerung ausstehend"
        
        Label3(0).ForeColor = glS1
        Label3(1).ForeColor = glS1
        Label3(2).ForeColor = glS1
        Label3(3).ForeColor = glS1
        Label3(4).ForeColor = glS1
        
        zeigmalWochentagan 0
        zeigmalWochentagan 1
        zeigmalWochentagan 2
        zeigmalWochentagan 3
        zeigmalWochentagan 4
        
    Else
        'alles auf anfang
        
        Label2(0).Caption = "BEHANDLUNGSDATUM"
        Label2(1).Caption = "BEHANDLUNGSDATUM"
        Label2(2).Caption = "BEHANDLUNGSDATUM"
        Label2(3).Caption = "BEHANDLUNGSDATUM"
        Label2(4).Caption = "BEHANDLUNGSDATUM"
        
        Label3(0).Caption = "Erinnerung ausstehend"
        Label3(1).Caption = "Erinnerung ausstehend"
        Label3(2).Caption = "Erinnerung ausstehend"
        Label3(3).Caption = "Erinnerung ausstehend"
        Label3(4).Caption = "Erinnerung ausstehend"
        
        
        
        Label4(0).Caption = ""
        Label4(1).Caption = ""
        Label4(2).Caption = ""
        Label4(3).Caption = ""
        Label4(4).Caption = ""
     
        'DATUMBEHANDLUNG,DATUMERINNERUNG
        Label2(0).Visible = False
        Label3(0).Visible = False
        Label3(0).Caption = ""
        Label4(0).Visible = False
        Command1(0).Visible = False

        cSQL = "Select TOP 2 DATUMBEHANDLUNG from SMS_UEBERSICHT"
        cSQL = cSQL & " order by DATUMBEHANDLUNG desc "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
        
            Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!DATUMBEHANDLUNG) Then
                If UCase(Label2(1).Caption) = "BEHANDLUNGSDATUM" Then
                    Label2(1).Caption = rsrs!DATUMBEHANDLUNG
                    Command1(1).Enabled = False
                    
                    cSQL = "Select DATUMERINNERUNG from SMS_UEBERSICHT where DATUMBEHANDLUNG = " & CLng(DateValue(Label2(1).Caption))
                    Set rsRs2 = gdBase.OpenRecordset(cSQL)
                    If Not rsRs2.EOF Then
                        If Not IsNull(rsRs2!DATUMERINNERUNG) Then
                            Label3(1).Caption = rsRs2!DATUMERINNERUNG
                        End If
                    End If
                    rsRs2.Close: Set rsRs2 = Nothing
                    
                    zeigmalWochentagan 1
                    
                Else
                    Label2(0).Caption = rsrs!DATUMBEHANDLUNG
                    
                    cSQL = "Select DATUMERINNERUNG from SMS_UEBERSICHT where DATUMBEHANDLUNG = " & CLng(DateValue(Label2(0).Caption))
                    Set rsRs2 = gdBase.OpenRecordset(cSQL)
                    If Not rsRs2.EOF Then
                        If Not IsNull(rsRs2!DATUMERINNERUNG) Then
                            Label3(0).Caption = rsRs2!DATUMERINNERUNG
                        End If
                    End If
                    rsRs2.Close: Set rsRs2 = Nothing
                    
                    zeigmalWochentagan 0
                    
                    Label2(0).Visible = True
                    Label3(0).Visible = True
                    Label4(0).Visible = True
                    Command1(0).Visible = True
                    Command1(0).Enabled = False
                End If
            End If
            
            
            
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        
        
        Label2(2).Caption = DateAdd("d", 1, DateValue(Label2(1).Caption))
        Label2(3).Caption = DateAdd("d", 2, DateValue(Label2(1).Caption))
        Label2(4).Caption = DateAdd("d", 3, DateValue(Label2(1).Caption))
        
        
        Label3(2).Caption = "Erinnerung ausstehend"
        Label3(3).Caption = "Erinnerung ausstehend"
        Label3(4).Caption = "Erinnerung ausstehend"
        
        Label3(2).ForeColor = vbRed
        Label3(3).ForeColor = vbRed
        Label3(4).ForeColor = vbRed
        
        'Welcher Wochentag ist das
        zeigmalWochentagan 2
        zeigmalWochentagan 3
        zeigmalWochentagan 4
    End If
   

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WasBisherGeschah"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigmalWochentagan(iIndex As Integer)
On Error GoTo LOKAL_ERROR

    
    Dim lBehDat                 As Long
    Dim lWeekday                As Long
    
    lBehDat = CLng(DateValue(Label2(iIndex).Caption))
    lWeekday = Weekday(lBehDat, vbMonday)

    Select Case lWeekday
        Case Is = 1 '"MO"
            Label4(iIndex).Caption = "Mo"
        Case Is = 2 '"DI"
            Label4(iIndex).Caption = "Di"
        Case Is = 3 '"MI"
            Label4(iIndex).Caption = "Mi"
        Case Is = 4 '"DO"
            Label4(iIndex).Caption = "Do"
        Case Is = 5 '"FR"
            Label4(iIndex).Caption = "Fr"
        Case Is = 6 '"SA"
            Label4(iIndex).Caption = "Sa"
        Case Is = 7 '"SO"
            Label4(iIndex).Caption = "So"
    End Select
    
    If CLng(DateValue(Label2(iIndex).Caption)) = CLng(DateValue(Now)) Then
        Label4(iIndex).Caption = "Heute"
    End If
    
    If CLng(DateValue(Label2(iIndex).Caption)) = CLng(DateValue(Now) - 1) Then
        Label4(iIndex).Caption = "Gestern"
    End If
    
    If CLng(DateValue(Label2(iIndex).Caption)) = CLng(DateValue(Now) + 1) Then
        Label4(iIndex).Caption = "Morgen"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WasBisherGeschah"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 7
            loeschNEW "SMS_UEBERSICHT", gdBase
            WasBisherGeschah
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil SMS-Erinnerung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
