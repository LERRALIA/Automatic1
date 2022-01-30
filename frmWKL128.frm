VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL128 
   Caption         =   "Kassenbuch"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL128.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   405
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   6
      Tag             =   "2"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11535
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
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
      Caption         =   "Drucken"
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
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11535
   End
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   20
      Left            =   6120
      TabIndex        =   10
      ToolTipText     =   "Kalender"
      Top             =   360
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Wählen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   21
      Left            =   8400
      TabIndex        =   11
      ToolTipText     =   "Kalender"
      Top             =   360
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Wählen Sie hier das Datum aus."
      ToolTipTitle    =   "Kalender"
      ButtonStyle     =   2
      Caption         =   ""
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   165
      Left            =   8040
      TabIndex        =   12
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Zurück"
      ToolTipTitle    =   "Zurück"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   165
      Left            =   8040
      TabIndex        =   13
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Vor"
      ToolTipTitle    =   "Vor"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command8 
      Height          =   165
      Left            =   5760
      TabIndex        =   14
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Zurück"
      ToolTipTitle    =   "Zurück"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command7 
      Height          =   165
      Left            =   5760
      TabIndex        =   15
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Vor"
      ToolTipTitle    =   "Vor"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   17
      Top             =   240
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
      Caption         =   "Export"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Achtung: Das Kassenbuch wird nur dann korrekt geführt, wenn jede Kasse täglich am Ende des Tages einen Kassenabschluss vornimmt."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   7200
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Datum von:"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Datum bis:"
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
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   1215
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
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kassenbuch"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmWKL128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            ZeigeKaBuch Text1(0).Text, Text1(1).Text
            
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            ZeigeKaBuch Text1(0).Text, Text1(1).Text
            'fertig
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL128
        Case 1
            DruckeKaBuch
        Case 2
            Export_KaBuch_alle Text1(0).Text, Text1(1).Text
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeKaBuch()
On Error GoTo LOKAL_ERROR

Dim lcount As Long
Dim cZeile As String
Dim cSQL As String


loeschNEW "PRINTQ", gdBase
CreateTable "PRINTQ", gdBase

loeschNEW "FIRMADRUCK", gdBase
CreateTable "FIRMADRUCK", gdBase

cSQL = "Insert into FIRMADRUCK select"
cSQL = cSQL & " NAME "
cSQL = cSQL & ", STRASSE "
cSQL = cSQL & ", PLZ "
cSQL = cSQL & ", ORT "
cSQL = cSQL & ", TEL "
cSQL = cSQL & ", FAX "
cSQL = cSQL & ", STEUERNR "
cSQL = cSQL & ", EMAIL "
cSQL = cSQL & ", " & gcKasNum & " as KASNUM "
cSQL = cSQL & " from Firma "
gdBase.Execute cSQL, dbFailOnError

If List1.ListCount > 0 Then
    For lcount = 0 To List1.ListCount - 1
    
        cZeile = List1.list(lcount)
        
        cSQL = "Insert into PRINTQ (Zeile) values ('" & cZeile & "')"
        gdBase.Execute cSQL, dbFailOnError
        
    Next lcount
    reportbildschirm "WKL029", "aWKL128"
End If

Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKaBuch"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKaBuch(cVon As String, cBis As String)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset
Dim cFeld As String
Dim cLBSatz As String
Dim lMerkdat    As Long

List1.Clear
List2.Clear
List2.AddItem Space(66) & "Umsatz   Bargeld"

sSQL = " Select * from KABUCH where Kasnum = " & gcKasNum

If cVon <> "" Then
    sSQL = sSQL & " and Datum >= " & CLng(DateValue(cVon))
End If

If cBis <> "" Then
    sSQL = sSQL & " and Datum <= " & CLng(DateValue(cBis))
End If

sSQL = sSQL & "  order by autopos"
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Datum) Then
            cFeld = Format(rsrs!Datum, "DD.MM.YY")
        Else
            cFeld = ""
        End If
        cFeld = cFeld & Space$(9 - Len(cFeld))
        cLBSatz = cFeld
    
        If Not IsNull(rsrs!Pos) Then
            cFeld = rsrs!Pos
        Else
            cFeld = ""
        End If
        
        If cFeld = "1" Then
            List1.AddItem "__________________________________________________________________________________"
        End If
        
        cFeld = cFeld & Space$(3 - Len(cFeld))
        cLBSatz = cLBSatz & cFeld
        
        If Not IsNull(rsrs!BEZUMS) Then
            cFeld = rsrs!BEZUMS
        Else
            cFeld = ""
        End If
        
        If cFeld = "gesamt" Then
            List1.AddItem Space(64) & "__________________"
        End If
        cFeld = cFeld & Space$(50 - Len(cFeld))
        cLBSatz = cLBSatz & cFeld
        
        If Not IsNull(rsrs!EURUMS) Then
            cFeld = Format(rsrs!EURUMS, "#####0.00")
        Else
            cFeld = ""
        End If
        cFeld = Space$(10 - Len(cFeld)) & cFeld
        cLBSatz = cLBSatz & cFeld
        
        If Not IsNull(rsrs!EURBAR) Then
            cFeld = Format(rsrs!EURBAR, "#####0.00")
        Else
            cFeld = ""
        End If
        cFeld = Space$(10 - Len(cFeld)) & cFeld
        cLBSatz = cLBSatz & cFeld
    
        List1.AddItem cLBSatz
        
    
    
    rsrs.MoveNext
    Loop
End If
rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKaBuch"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Export_KaBuch_alle(cVon As String, cBis As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsDisKas        As Recordset
    Dim cFeld           As String
    Dim cLBSatz         As String
    Dim sdat            As String
    Dim cPfad1          As String
    Dim cdatei          As String
    Dim cPfad           As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim iRet            As Integer
    Dim cStrich         As String
    Dim cKasse          As String
    Dim bAnd            As Boolean
    Dim sAusgabedatname As String
    bAnd = False

    sdat = Format$(cVon, "DDMMYY") & "_" & Format$(cBis, "DDMMYY")
    sAusgabedatname = "Kassenbuch_" & sdat & ".txt"

    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    
    
    
    
    loeschNEW "KABUCH_EXPORT", gdBase
    sSQL = "Select * into KABUCH_EXPORT from KABUCH "
    If cVon <> "" Then
    
        If bAnd Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        sSQL = sSQL & " Datum >= " & CLng(DateValue(cVon))
        bAnd = True
    End If
    
    If cBis <> "" Then
        If bAnd Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        sSQL = sSQL & " Datum <= " & CLng(DateValue(cBis))
        bAnd = True
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Select distinct(kasnum) as kasse from KABUCH_EXPORT order by kasnum"
    Set rsDisKas = gdBase.OpenRecordset(sSQL)
    If Not rsDisKas.EOF Then
        rsDisKas.MoveFirst
        Do While Not rsDisKas.EOF
        
            cKasse = "-1"
        
            If Not IsNull(rsDisKas!Kasse) Then
                cKasse = Trim(rsDisKas!Kasse)
            End If
            
            
            
            cSatz = "Kasse: " & cKasse & " Filiale: " & gcFilNr & Space(47) & "Umsatz   Bargeld" & Chr$(13) & Chr$(10)

            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            
            
            sSQL = " Select * from KABUCH_EXPORT where Kasnum = " & cKasse
            sSQL = sSQL & "  order by autopos"
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                    
                    If Not IsNull(rsrs!Datum) Then
                        cFeld = Format(rsrs!Datum, "DD.MM.YY")
                    Else
                        cFeld = ""
                    End If
                    cFeld = cFeld & Space$(9 - Len(cFeld))
                    cLBSatz = cFeld
                
                    If Not IsNull(rsrs!Pos) Then
                        cFeld = rsrs!Pos
                    Else
                        cFeld = ""
                    End If
                    
                    If cFeld = "1" Then
                        cStrich = "__________________________________________________________________________________" & vbCrLf
                        
                        lPos = LOF(iFileNr)
                        lPos = lPos + 1
                        Put #iFileNr, lPos, cStrich
                        
                    End If
                    
                    cFeld = cFeld & Space$(3 - Len(cFeld))
                    cLBSatz = cLBSatz & cFeld
                    
                    If Not IsNull(rsrs!BEZUMS) Then
                        cFeld = rsrs!BEZUMS
                    Else
                        cFeld = ""
                    End If
                    
                    If cFeld = "gesamt" Then
                        cStrich = Space(64) & "__________________" & vbCrLf
                        
                        lPos = LOF(iFileNr)
                        lPos = lPos + 1
                        Put #iFileNr, lPos, cStrich
                        
                    End If
                    cFeld = cFeld & Space$(50 - Len(cFeld))
                    cLBSatz = cLBSatz & cFeld
                    
                    If Not IsNull(rsrs!EURUMS) Then
                        cFeld = Format(rsrs!EURUMS, "#####0.00")
                    Else
                        cFeld = ""
                    End If
                    cFeld = Space$(10 - Len(cFeld)) & cFeld
                    cLBSatz = cLBSatz & cFeld
                    
                    If Not IsNull(rsrs!EURBAR) Then
                        cFeld = Format(rsrs!EURBAR, "#####0.00")
                    Else
                        cFeld = ""
                    End If
                    cFeld = Space$(10 - Len(cFeld)) & cFeld
                    cLBSatz = cLBSatz & cFeld
                
                    cLBSatz = cLBSatz & Chr$(13) & Chr$(10)
                        
                    lPos = LOF(iFileNr)
                    lPos = lPos + 1
                    Put #iFileNr, lPos, cLBSatz
                    
                rsrs.MoveNext
                Loop
            End If
            rsrs.Close: Set rsrs = Nothing
                    
            
            
            
            cLBSatz = Chr$(13) & Chr$(10)
                        
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cLBSatz
            
            
            
        
        
        
        rsDisKas.MoveNext
        Loop
    End If
    rsDisKas.Close: Set rsDisKas = Nothing
    
    
    
    
    
    Close iFileNr
    
    If Datendrin("KABUCH_EXPORT", gdBase) Then
        iRet = MsgBox("Möchten Sie diese Textdatei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            gcBestellEmail.Subject = "Kassenbücher (" & cVon & " - " & cBis & ") Filiale: " & gcFilNr
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
      
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Export_KaBuch_alle"
        Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
    
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        Else
            
        End If
        
        
        lDat = lDat + 1
        
        
        Text1(0).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    ZeigeKaBuch Text1(0).Text, Text1(1).Text
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    
    Else
    
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        Else
            
        End If
        
        
        lDat = lDat - 1
        
        
        Text1(0).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    
    ZeigeKaBuch Text1(0).Text, Text1(1).Text
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
    
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        Else
           
        End If
        
        
        lDat = lDat + 1
        
        
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    ZeigeKaBuch Text1(0).Text, Text1(1).Text
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long
    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        Else
            
        End If
        
        lDat = lDat - 1
        
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
    End If
    ZeigeKaBuch Text1(0).Text, Text1(1).Text
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
    Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")

    ZeigeKaBuch Text1(0).Text, Text1(1).Text
    
    Label1(2).ForeColor = glWarn
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "KABUCH_EXPORT", gdBase
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

