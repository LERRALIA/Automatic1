VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL72 
   Caption         =   "Plakate"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL72.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboFormat 
      Height          =   330
      Left            =   8640
      TabIndex        =   28
      Top             =   6120
      Width           =   3015
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   21
      Top             =   5520
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "durchgestrichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   20
      Top             =   4920
      Width           =   1575
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   11
      Top             =   6600
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Druckvorschau"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   10
      Top             =   4920
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   11535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmWKL72.frx":0442
      Left            =   6480
      List            =   "frmWKL72.frx":0444
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   6360
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Laden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   11160
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   1
      Left            =   2520
      TabIndex        =   24
      Top             =   1080
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   2
      Left            =   2520
      TabIndex        =   25
      Top             =   2040
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   3
      Left            =   2520
      TabIndex        =   26
      Top             =   4440
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   4
      Left            =   2520
      TabIndex        =   27
      Top             =   5640
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   300
      Index           =   5
      Left            =   2520
      TabIndex        =   22
      Top             =   6840
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "Farbe wählen"
      BevelWidth      =   1
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Plakate"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9015
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
      BackColor       =   &H00C0C000&
      Caption         =   "Zeile 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Zeile 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Zeile 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "alter Preis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "neuer Preis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   14
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Zeile 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   9375
   End
End
Attribute VB_Name = "frmWKL72"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
Dim sSQL As String

anzeige "normal", "", Label6

Select Case Index

    Case 0
        
        sSQL = "Delete from  SondText where name is null "
        gdBase.Execute sSQL, dbFailOnError
        
        
        voreinstellungspeichernE72B
        
        
        Unload frmWKL72
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub delete()
    On Error GoTo LOKAL_ERROR

    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    
    Dim i As Integer
    For i = 0 To 5
        SSCommand1(i).ForeColor = vbBlack
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delete"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
    
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    Dim sdateiname As String
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0 'Speichern
                
                sdateiname = InputBox("Wollen Sie diese Zusammenstellung speichern?. Dann vergeben Sie bitte einen Namen!", "Sonderangebot Frage:")
                speichern Trim(sdateiname)
                listeerneuern
                
                delete
                
                Text1(0).SetFocus
            
        Case Is = 1 'Drucken
            Text1(0).SetFocus
            
            drucken
        Case Is = 2
            list1_KeyUp 46, 1
        Case Is = 3
            delete
            Text1(0).SetFocus
        Case Is = 4
            sdateiname = Trim(List1.list(List1.ListIndex))
            anzeigeN sdateiname
        
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

loeschNEW "SONDDRU", gdBase

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub

Private Sub list1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sdatname As String
    Dim sSQL As String
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not List1.ListIndex = -1 Then
                sdatname = Trim(List1.list(List1.ListIndex))
                List1.RemoveItem (List1.ListIndex)
                
                sSQL = " Delete From Sondtext where name = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "list1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub listeerneuern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "Delete from Sondtext where name is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Sondtext where name = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select name from Sondtext where name <> null order by name"
    Set rs = gdBase.OpenRecordset(sSQL)
    List1.Clear
    Do While Not rs.EOF
        List1.AddItem rs!name
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "listeerneuern"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub anzeigeN(name As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsSo As Recordset
    
    sSQL = "Select * From Sondtext where name = '" & name & "' "
    
    Set rsSo = gdBase.OpenRecordset(sSQL)
        
        If Not rsSo.EOF Then
            rsSo.MoveFirst
        
            If Not IsNull(rsSo!Ueber) Then
                Text1(0).Text = rsSo!Ueber
            Else
                Text1(0).Text = ""
            End If
            
            If Not IsNull(rsSo!Zeile1) Then
                Text1(1).Text = rsSo!Zeile1
            Else
                Text1(1).Text = ""
            End If
            
            If Not IsNull(rsSo!altPr) Then
                Text1(3).Text = rsSo!altPr
            Else
                Text1(3).Text = ""
            End If
            
            If Not IsNull(rsSo!neuPr) Then
                Text1(4).Text = rsSo!neuPr
            Else
                Text1(4).Text = ""
            End If
            
            If Not IsNull(rsSo!Zeile2) Then
                Text1(2).Text = rsSo!Zeile2
            Else
                Text1(2).Text = ""
            End If
            
            If Not IsNull(rsSo!Zeile3) Then
                Text1(5).Text = rsSo!Zeile3
            Else
                Text1(5).Text = ""
            End If
            
            
            
            If Not IsNull(rsSo!UeberS) Then
                SSCommand1(0).ForeColor = rsSo!UeberS
            Else
                SSCommand1(0).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!Zeile1S) Then
                SSCommand1(1).ForeColor = rsSo!Zeile1S
            Else
                SSCommand1(1).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!Zeile2S) Then
                SSCommand1(2).ForeColor = rsSo!Zeile2S
            Else
                SSCommand1(2).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!altPrS) Then
                SSCommand1(3).ForeColor = rsSo!altPrS
            Else
                SSCommand1(3).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!neuPrS) Then
                SSCommand1(4).ForeColor = rsSo!neuPrS
            Else
                SSCommand1(4).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!Zeile3S) Then
                SSCommand1(5).ForeColor = rsSo!Zeile3S
            Else
                SSCommand1(5).ForeColor = vbBlack
            End If
            
            If Not IsNull(rsSo!durchge) Then
                If rsSo!durchge = True Then
                    Check1.Value = vbChecked
                Else
                    Check1.Value = vbUnchecked
                End If
            Else
                Check1.Value = vbUnchecked
            End If
            
            
        End If
    rsSo.Close
    
    Text1(0).Text = SwapStr(Text1(0).Text, "$", "'")
    Text1(1).Text = SwapStr(Text1(1).Text, "$", "'")
    Text1(2).Text = SwapStr(Text1(2).Text, "$", "'")
    Text1(3).Text = SwapStr(Text1(3).Text, "$", "'")
    Text1(4).Text = SwapStr(Text1(4).Text, "$", "'")
    Text1(5).Text = SwapStr(Text1(5).Text, "$", "'")
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigen"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub speichern(name As String)
    On Error GoTo LOKAL_ERROR
    
'    Text1(0).Text = SwapStr(Text1(0).Text, "'", " ")
'    Text1(1).Text = SwapStr(Text1(1).Text, "'", " ")
'    Text1(2).Text = SwapStr(Text1(2).Text, "'", " ")
'    Text1(3).Text = SwapStr(Text1(3).Text, "'", " ")
'    Text1(4).Text = SwapStr(Text1(4).Text, "'", " ")
'    Text1(5).Text = SwapStr(Text1(5).Text, "'", " ")
    
    
    
    
    Dim cText1 As String
    Dim cText2 As String
    Dim cText3 As String
    Dim cText4 As String
    Dim cText5 As String
    Dim cText6 As String
    
    cText1 = Text1(0).Text
    cText2 = Text1(1).Text
    cText3 = Text1(2).Text
    cText4 = Text1(3).Text
    cText5 = Text1(4).Text
    cText6 = Text1(5).Text
    
    cText1 = SwapStr(cText1, "'", "$")
    cText2 = SwapStr(cText2, "'", "$")
    cText3 = SwapStr(cText3, "'", "$")
    cText4 = SwapStr(cText4, "'", "$")
    cText5 = SwapStr(cText5, "'", "$")
    cText6 = SwapStr(cText6, "'", "$")
    
    
    
    
    
    Dim sSQL As String
    sSQL = " Insert into SondText "
    sSQL = sSQL & "( Name  "
    sSQL = sSQL & ", Ueber "
    sSQL = sSQL & ", Zeile1 "
    sSQL = sSQL & ", Zeile2 "
    sSQL = sSQL & ", altPr "
    sSQL = sSQL & ", neuPr "
    sSQL = sSQL & ", Zeile3 "
    
    sSQL = sSQL & ", UeberS "
    sSQL = sSQL & ", Zeile1S "
    sSQL = sSQL & ", Zeile2S "
    sSQL = sSQL & ", altPrS "
    sSQL = sSQL & ", neuPrS "
    sSQL = sSQL & ", Zeile3S "
    sSQL = sSQL & ", durchge "
    
    sSQL = sSQL & ") "
    sSQL = sSQL & " Values ( "
    sSQL = sSQL & "'" & name & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText1 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText2 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText3 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText4 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText5 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText6 & "' "

    sSQL = sSQL & " , " & SSCommand1(0).ForeColor
    sSQL = sSQL & " , " & SSCommand1(1).ForeColor
    sSQL = sSQL & " , " & SSCommand1(2).ForeColor
    sSQL = sSQL & " , " & SSCommand1(3).ForeColor
    sSQL = sSQL & " , " & SSCommand1(4).ForeColor
    sSQL = sSQL & " , " & SSCommand1(5).ForeColor
    
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " , True "
    Else
        sSQL = sSQL & " , False "
    End If
    
    sSQL = sSQL & ") "
'    MsgBox sSQL
    gdBase.Execute sSQL, dbFailOnError
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub drucken()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "SONDDRU", gdBase
    CreateTable "SONDDRU", gdBase
    
    Dim cText1 As String
    Dim cText2 As String
    Dim cText3 As String
    Dim cText4 As String
    Dim cText5 As String
    Dim cText6 As String
    
    cText1 = Text1(0).Text
    cText2 = Text1(1).Text
    cText3 = Text1(2).Text
    cText4 = Text1(3).Text
    cText5 = Text1(4).Text
    cText6 = Text1(5).Text
    
    cText1 = SwapStr(cText1, "'", "''")
    cText2 = SwapStr(cText2, "'", "''")
    cText3 = SwapStr(cText3, "'", "''")
    cText4 = SwapStr(cText4, "'", "''")
    cText5 = SwapStr(cText5, "'", "''")
    cText6 = SwapStr(cText6, "'", "''")
    
    sSQL = " Insert into SondDru "
    sSQL = sSQL & "( "
    sSQL = sSQL & " Ueber "
    sSQL = sSQL & ", Zeile1 "
    sSQL = sSQL & ", Zeile2 "
    sSQL = sSQL & ", altPr "
    sSQL = sSQL & ", neuPr "
    sSQL = sSQL & ", Zeile3 "
    
    sSQL = sSQL & ", UeberS "
    sSQL = sSQL & ", Zeile1S "
    sSQL = sSQL & ", Zeile2S "
    sSQL = sSQL & ", altPrS "
    sSQL = sSQL & ", neuPrS "
    sSQL = sSQL & ", Zeile3S "
    sSQL = sSQL & ", durchge "
    sSQL = sSQL & ") "
    sSQL = sSQL & " Values ( "
    sSQL = sSQL & "'" & cText1 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText2 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText3 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText4 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText5 & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & cText6 & "' "
    
    sSQL = sSQL & " , " & SSCommand1(0).ForeColor
    sSQL = sSQL & " , " & SSCommand1(1).ForeColor
    sSQL = sSQL & " , " & SSCommand1(2).ForeColor
    sSQL = sSQL & " , " & SSCommand1(3).ForeColor
    sSQL = sSQL & " , " & SSCommand1(4).ForeColor
    sSQL = sSQL & " , " & SSCommand1(5).ForeColor
    
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " , True "
    Else
        sSQL = sSQL & " , False "
    End If
    
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
    
    If cboFormat.Text <> "bitte wählen" Then
        Select Case cboFormat.Text
            Case "Variante 1"
                reportbildschirm "", "awkl72a"
            Case "Variante 2"
                reportbildschirm "", "awkl72b"
            Case "DroNova (Querformat)"
                reportbildschirm "", "awkl72c"
            Case "DroNova (Querformat) 2"
                reportbildschirm "", "awkl72d"
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    If Not NewTableSuchenDBKombi("SONDTEXT", gdBase) Then
        CreateTable "SONDTEXT", gdBase
    End If
   
    listeerneuern
    
    delete
    
    cboFormat.AddItem "bitte wählen"
    cboFormat.AddItem "Variante 1"
    cboFormat.AddItem "Variante 2"
    cboFormat.AddItem "DroNova (Querformat)"
    cboFormat.AddItem "DroNova (Querformat) 2"
    
    cboFormat.Text = "bitte wählen"
    
    If NewTableSuchenDBKombi("E72B", gdBase) Then
        
        voreinstellungladenE72B
    
    End If
    
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE72B()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    
    loeschNEW "E72B", gdBase
    CreateTableT2 "E72B", gdBase
    
    sSQL = "Insert into E72B ( Druckformat) "
    sSQL = sSQL & " values ('" & cboFormat.Text & "')"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE72B"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE72B()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("E72B")
    If Not rs.EOF Then

        If Not IsNull(rs!Druckformat) Then
            cboFormat.Text = rs!Druckformat
        Else
            cboFormat.Text = "bitte wählen"
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE72B"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        Case 1
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        Case 2
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        Case 3
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        Case 4
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        Case 5
            cdl1.ShowColor
            SSCommand1(Index).ForeColor = cdl1.Color
        
    End Select
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
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
    Fehler.gsFehlertext = "Im Programmteil Plakate ist ein Fehler aufgetreten."
    Fehlermeldung1

End Sub



