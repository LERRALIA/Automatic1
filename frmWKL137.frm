VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL137 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Bonansichten"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check17 
      Caption         =   "als Nettorechnung?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6360
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   360
      Index           =   4
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   27
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "alle angezeigten drucken"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   26
      Top             =   480
      Width           =   2655
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
      Height          =   360
      Index           =   3
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   19
      Top             =   1680
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
   Begin VB.CheckBox Check5 
      Caption         =   "im DIN A4 Format"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   2040
      Width           =   2055
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
      Height          =   360
      Index           =   2
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1920
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
      Height          =   360
      Index           =   1
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "nur Storno"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   1200
      Width           =   3930
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   8520
      TabIndex        =   8
      Top             =   1680
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   7815
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
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cmbFilialen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2730
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   5
      Top             =   360
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   4
      Top             =   960
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
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
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   7815
   End
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   0
      Left            =   1800
      TabIndex        =   23
      ToolTipText     =   "Kalender"
      Top             =   1200
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
   Begin sevCommand3.Command Command8 
      Height          =   165
      Left            =   1320
      TabIndex        =   24
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
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
      Left            =   1320
      TabIndex        =   25
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bontext:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   28
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kasse:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   22
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunde:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   18
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bon:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Zahlungsarten:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   12
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Filialen:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonansichten"
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
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5040
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
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL137"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim KASSBON_DB        As Database

Private Sub Check5_Click()
On Error GoTo LOKAL_ERROR

    If Check5.Value = vbChecked Then
        Check17.Visible = True
        Check17.Value = vbUnchecked
    Else
        Check17.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim rec         As Recordset
    Dim cSatz       As String
    Dim cFeld       As String
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    Dim cPfad           As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\KASSBON.MDB"
    
    Set KASSBON_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsKASSBON_Passwort)
'        KASSBON_DB.Close
        
    Text1(0).Text = DateValue(Now) - 1
    
    Combo2.Clear
    Combo2.AddItem "Alle Zahlungsarten"
    Combo2.AddItem "BA Barzahlung"
    Combo2.AddItem "EC Kartenzahlung EC"
    Combo2.AddItem "AE Kartenzahlung American Express"
    Combo2.AddItem "VI Kartenzahlung VISA"
    Combo2.AddItem "EU Kartenzahlung Eurocard / Mastercard"
    Combo2.AddItem "DC Kartenzahlung Diners Club"
    Combo2.AddItem "BC Kartenzahlung Barclaycard"
    Combo2.AddItem "SO Kartenzahlung sonstige"
    Combo2.AddItem "GZ gemischte Zahlung"
    Combo2.AddItem "LS Lastschrift"
    Combo2.AddItem "KR Kreditkauf"
    Combo2.Text = "Alle Zahlungsarten"
    
    Command5_Click 1
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 0
            
        Case 1
            Command5_Click 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub cmbFilialen_Click()
On Error GoTo LOKAL_ERROR
    
    Command5_Click 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmbFilialen_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Combo2_Click()
On Error GoTo LOKAL_ERROR
    
    Command5_Click 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
            Command5_Click 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
     
    Select Case Index
        Case 11
            gsHelpstring = "Bonansichten"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
    Else
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        End If
        lDat = lDat + 1
        Text1(0).Text = Format(lDat, "DD.MM.YY")
    End If
    
    Command5_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
    Else
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        End If
        lDat = lDat - 1
        Text1(0).Text = Format(lDat, "DD.MM.YY")
    End If
    
    Command5_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cKKart As String
    Dim cFil As String
    Dim cBon As String
    Dim cKunde As String
    Dim cKassnum As String
    Dim cDatum As String
    Dim cBonNr As String
    Dim cBontext As String
    Dim bSton   As Boolean

    Select Case Index
        Case 0
            Unload frmWKL137
        Case 1 'suchen
            cFil = "0"
            If Combo2.Text = "Alle Zahlungsarten" Then
                cKKart = "0"
            Else
                cKKart = Left(Combo2.Text, 2)
            End If
            
            cBon = ""
            If Text1(1).Text <> "" Then
                If IsNumeric(Text1(1).Text) Then
                    cBon = Text1(1).Text
                End If
            End If
            
            cKunde = ""
            If Text1(2).Text <> "" Then
                If IsNumeric(Text1(2).Text) Then
                    cKunde = Text1(2).Text
                End If
            End If
            
            cKassnum = ""
            If Text1(3).Text <> "" Then
                If IsNumeric(Text1(3).Text) Then
                    cKassnum = Text1(3).Text
                End If
            End If
            
            cDatum = ""
            If Text1(0).Text <> "" Then
                If IsDate(Text1(0).Text) Then
                    cDatum = Text1(0).Text
                End If
            End If
            
            cBontext = ""
            If Text1(4).Text <> "" Then
                
                cBontext = Text1(4).Text
                
            End If
            
            If Check1(1).Value = vbChecked Then
                bSton = True
            Else
                bSton = False
            End If
            
            If IsDate(cDatum) = False Then
                anzeige "rot", "Datum ist falsch", Label1(4)
                Text1(0).SetFocus
                Exit Sub
            End If
            
            AnlistenBonDatenNeu Text1(0).Text, cFil, cKKart, bSton, cBon, cKunde, cKassnum, cBontext
            List2.Clear
            If List1.ListCount > 0 Then
                List1.ListIndex = 0
            End If
        Case 2
        
            If Check1(0).Value = vbChecked Then
            
                If List1.ListCount > 0 Then
                
                
                    For i = 0 To List1.ListCount - 1
                        
                        
                        List1.Selected(i) = True
                        DruckeZweitBonAusListe List2, False
                        
                    Next i
                
                
                End If
            
            Else
                If List1.ListIndex < 0 Then
                    anzeige "rot", "Bitte einen Kassenbon in der Liste auswählen!", Label1(4)
                    List1.SetFocus
                Else
    
                    cBonNr = List1.list(List1.ListIndex)
                    
                    If Check5.Value = vbChecked Then 'also dina4 druck
                        setzedrucker gcListenDrucker
                        
                        
                        If Check17.Value = vbChecked Then 'Nettorechnung
                            DruckeKassenZweitBonDinA4_Netto CLng(DateValue(Text1(0).Text)), cBonNr, "Rechnung", gbDINA4RECHFU
                        Else
                            DruckeKassenZweitBonDinA4 CLng(DateValue(Text1(0).Text)), cBonNr, "2. Bon", gbDINA4RECHFU ', cKasnum
                        End If
                        
                        
                        
                        
                        
                        setzedrucker gcBonDrucker
                    Else
                        DruckeZweitBonAusListe List2, False
    
                    End If
                End If
            End If

    End Select
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeKassenZweitBonDinA4(lDate As Long, cAuswahl As String, cZahlart As String, bReFuss As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cFirma          As String
    Dim cVname          As String
    Dim cNName          As String
    Dim cTitel          As String
    Dim cPlz            As String
    Dim cStadt          As String
    Dim cStrasse        As String
    Dim cAnrede         As String
    Dim cPrintFiTiAnVoNa    As String
    
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim cLBSatz         As String
    Dim rs              As Recordset
    Dim sSQL            As String
    Dim dUSTV           As Double
    Dim dUSTE           As Double
    
    loeschNEW "ANGEBOTNOW", gdBase
    CreateTable "ANGEBOTNOW", gdBase
    
    loeschNEW "DAGKOPF", gdBase
    CreateTable "DAGKOPF", gdBase
    
    '******
    Dim lAnzZeile         As Long
    Dim lHeute            As Long
    Dim lWeekday          As Long
    Dim lSuchtag          As Long
    Dim lPos              As Long
    Dim lStart            As Long
    Dim lZeile            As Long
    
    Dim aDeviceName       As String
    Dim cEscapeSequenz    As String
    Dim cSQL              As String
    Dim cFeld             As String
    Dim cZeile            As String
    Dim cBonNr            As String
    Dim cUhrZeit          As String
    Dim cBetrag           As String
    Dim cKundnr           As String
    Dim rsrs              As Recordset
    
    ReDim cDruckZeile(1 To 1) As String
    
    cBonNr = Trim$(Left(cAuswahl, 6))
    cUhrZeit = Right(cAuswahl, 5)
    
    lHeute = lDate
    
    cSQL = "Select * from ANGEBOTNOW "
    FnOpenrecordset rs, cSQL, 1, gdBase
    
    cSQL = "Select * from Kassjour "
    cSQL = cSQL & "where adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and BELEGNR = " & cBonNr & " "
    lAktSatz = 0
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
       
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            rs.AddNew
            rs!artnr = rsrs!artnr
            rs!BEZEICH = rsrs!BEZEICH
            If Not IsNull(rsrs!Menge) Then
                If Val(rsrs!Menge) <> 0 Then
                    rs!vkpr = rsrs!Preis / rsrs!Menge
                Else
                    rs!vkpr = rsrs!Preis
                End If
            Else
                rs!vkpr = rsrs!Preis
            End If
            rs!ANZAHL = rsrs!Menge
            rs!MWST = rsrs!MWST
            rs!KVKPR1 = rsrs!vkpr
            rs!lfnr = cBonNr
            rs!posinr = lAktSatz
            rs!BEDNR = rsrs!BEDIENER
            
            
            rs!MOPREIS = rsrs!MOPREIS
            
            rs.Update
            
            rsrs.MoveNext
        Loop
    End If
    
    '*****

    rsrs.Close: Set rsrs = Nothing
    rs.Close: Set rs = Nothing
    
    
    cSQL = "Select kundnr from Kassjour "
    cSQL = cSQL & "where adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and BELEGNR = " & cBonNr & " "
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Kundnr) Then
            cKundnr = rsrs!Kundnr
        Else
            cKundnr = "0"
        End If
    
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    
    cKundnr = Trim(cKundnr)
    
    If cKundnr = "" Then
        cKundnr = "0"
    End If

    dUSTV = ermUstv("ANGEBOTNOW")
    dUSTE = ermUste("ANGEBOTNOW")

    
    If cKundnr <> "0" Then
        
        cFirma = lookingForKundendaten(Trim(cKundnr)).firma
        cVname = lookingForKundendaten(Trim(cKundnr)).vorname
        cNName = lookingForKundendaten(Trim(cKundnr)).nachname
        cTitel = lookingForKundendaten(Trim(cKundnr)).titel
        cPlz = lookingForKundendaten(Trim(cKundnr)).Plz
        cStadt = lookingForKundendaten(Trim(cKundnr)).Ort
        cStrasse = lookingForKundendaten(Trim(cKundnr)).strasse
        cAnrede = lookingForKundendaten(Trim(cKundnr)).anrede
        
        cPrintFiTiAnVoNa = ""
        
        If cFirma <> "" Then
            cPrintFiTiAnVoNa = cFirma
        End If
        
        cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & vbCrLf
        
        If cAnrede <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cAnrede & Space(1)
        End If
        
        If cTitel <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cTitel & Space(1)
        End If
        
        If cVname <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cVname & Space(1)
        End If
        
        If cNName <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cNName
        End If
        
        sSQL = "Insert into DAGKOPF (kundnr,PrintFiTiAnVoNa,name,vorname,titel,plz,stadt,strasse,Firma,anrede,datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        sSQL = sSQL & cKundnr
        sSQL = sSQL & ", '" & cPrintFiTiAnVoNa & "' "
        sSQL = sSQL & ", '" & cNName & "' "
        sSQL = sSQL & ", '" & cVname & "' "
        sSQL = sSQL & ", '" & cTitel & "' "
        sSQL = sSQL & ", '" & cPlz & "' "
        sSQL = sSQL & ", '" & cStadt & "' "
        sSQL = sSQL & ", '" & cStrasse & "' "
        sSQL = sSQL & ", '" & cFirma & "' "
        sSQL = sSQL & ", '" & cAnrede & "' "
        sSQL = sSQL & " , '" & cZahlart & "' "
        sSQL = sSQL & " , '" & gcBediener & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
        
    Else
    
        sSQL = "Insert into DAGKOPF (datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        
        sSQL = sSQL & "  '' "
        sSQL = sSQL & " , '" & gcBediener & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Dim cFirmName       As String
    Dim cFirmAdress     As String
    Dim cFirmBank       As String
    Dim cFirmKomm       As String
    Dim cSteuernr       As String
    Dim cKommentar      As String
    
    If bReFuss Then
        loeschNEW "REFUSS", gdBase
        CreateTableT2 "REFUSS", gdBase
    
        cSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
            cSteuernr = ""
        End If
        
        cSQL = "Insert into REFUSS ( "
        cSQL = cSQL & " STEUERNR"
        cSQL = cSQL & ", FIRMNAME"
        cSQL = cSQL & ", FIRMADRESS"
        cSQL = cSQL & ", FIRMBANK"
        cSQL = cSQL & ", FIRMKOMM"
        cSQL = cSQL & ") values ("
        cSQL = cSQL & " '" & cSteuernr & "'"
        cSQL = cSQL & ", '" & cFirmName & "'"
        cSQL = cSQL & ", '" & cFirmAdress & "'"
        cSQL = cSQL & ", '" & cFirmBank & "'"
        cSQL = cSQL & ", '" & cFirmKomm & "'"
        cSQL = cSQL & ") "
        gdBase.Execute cSQL, dbFailOnError

    End If
    
    If Modul6.FindFile(gcDBPfad, "alrs3s.rpt") Then
        If gbDINA4RECHFU Then
            If Modul6.FindFile(gcDBPfad, "alrs3sfu.rpt") Then
                reportbildschirm "alr", "alrs3sfu"
            Else
                MsgBox "Ihre spezielle Druckansicht(alrs3sfu.rpt) muss angepasst werden, bitte rufen Sie die Hotline (Thomas Heinz, 0511 9559112) an!", vbOKOnly, "Winkiss Hinweis:"
            End If
        Else
            reportbildschirm "alr", "alrs3s"
        End If
    Else
        If gbDINA4RECHFU Then
            reportbildschirm "alr", "alrs3fu"
        Else
            reportbildschirm "alr", "alrs3"
        End If
    End If
    

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3134 Then
        schreibeProtokollDINA4Err "DruckeKassenZweitBonDinA4 Insert into Syntaxfehler"
        schreibeProtokollDINA4Err sSQL
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeKassenZweitBonDinA4"
        Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub

Private Sub DruckeKassenZweitBonDinA4_Netto(lDate As Long, cAuswahl As String, cZahlart As String, bReFuss As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cFirma          As String
    Dim cVname          As String
    Dim cNName          As String
    Dim cTitel          As String
    Dim cPlz            As String
    Dim cStadt          As String
    Dim cStrasse        As String
    Dim cAnrede         As String
    Dim cPrintFiTiAnVoNa    As String
    
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim cLBSatz         As String
    Dim rs              As Recordset
    Dim sSQL            As String
    Dim dUSTV           As Double
    Dim dUSTE           As Double
    
    loeschNEW "ANGEBOTNOW", gdBase
    CreateTable "ANGEBOTNOW", gdBase
    
    loeschNEW "DAGKOPF", gdBase
    CreateTable "DAGKOPF", gdBase
    
    '******
    Dim lAnzZeile         As Long
    Dim lHeute            As Long
    Dim lWeekday          As Long
    Dim lSuchtag          As Long
    Dim lPos              As Long
    Dim lStart            As Long
    Dim lZeile            As Long
    
    Dim aDeviceName       As String
    Dim cEscapeSequenz    As String
    Dim cSQL              As String
    Dim cFeld             As String
    Dim cZeile            As String
    Dim cBonNr            As String
    Dim cUhrZeit          As String
    Dim cBetrag           As String
    Dim cKundnr           As String
    Dim rsrs              As Recordset
    
    ReDim cDruckZeile(1 To 1) As String
    
    cBonNr = Trim$(Left(cAuswahl, 6))
    cUhrZeit = Right(cAuswahl, 5)
    
    lHeute = lDate
    
    cSQL = "Select * from ANGEBOTNOW "
    FnOpenrecordset rs, cSQL, 1, gdBase
    
    cSQL = "Select * from Kassjour "
    cSQL = cSQL & "where adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and BELEGNR = " & cBonNr & " "
    lAktSatz = 0
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
       
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            rs.AddNew
            rs!artnr = rsrs!artnr
            rs!BEZEICH = rsrs!BEZEICH
            If Not IsNull(rsrs!Menge) Then
                If Val(rsrs!Menge) <> 0 Then
                    rs!vkpr = rsrs!Preis / rsrs!Menge
                Else
                    rs!vkpr = rsrs!Preis
                End If
            Else
                rs!vkpr = rsrs!Preis
            End If
            rs!ANZAHL = rsrs!Menge
            rs!MWST = rsrs!MWST
            rs!KVKPR1 = rsrs!vkpr
            rs!lfnr = rsrs!BELEGNR
            rs!posinr = lAktSatz
            rs!BEDNR = rsrs!BEDIENER
            
            
            rs!MOPREIS = rsrs!MOPREIS
            
            rs.Update
            
            rsrs.MoveNext
        Loop
    End If
    
    '*****

    rsrs.Close: Set rsrs = Nothing
    rs.Close: Set rs = Nothing
    
    
    cSQL = "Update ANGEBOTNOW "
    cSQL = cSQL & " set "
    cSQL = cSQL & " vkpr = ((vkpr/(100 + " & gdMWStV & "))*100)"
    cSQL = cSQL & " where MWST = 'V' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ANGEBOTNOW "
    cSQL = cSQL & " set "
    cSQL = cSQL & " vkpr = ((vkpr/(100 + " & gdMWStE & "))*100)"
    cSQL = cSQL & " where MWST = 'E' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update ANGEBOTNOW "
    cSQL = cSQL & " set "
    cSQL = cSQL & " KVKPR1 = ((KVKPR1/(100 + " & gdMWStV & "))*100)"
    cSQL = cSQL & " where MWST = 'V' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ANGEBOTNOW "
    cSQL = cSQL & " set "
    cSQL = cSQL & " KVKPR1 = ((KVKPR1/(100 + " & gdMWStE & "))*100)"
    cSQL = cSQL & " where MWST = 'E' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    cSQL = "Select kundnr from Kassjour "
    cSQL = cSQL & "where adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & "and BELEGNR = " & cBonNr & " "
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Kundnr) Then
            cKundnr = rsrs!Kundnr
        Else
            cKundnr = "0"
        End If
    
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    
    cKundnr = Trim(cKundnr)
    
    If cKundnr = "" Then
        cKundnr = "0"
    End If

    dUSTV = ermUstv("ANGEBOTNOW")
    dUSTE = ermUste("ANGEBOTNOW")

    
    If cKundnr <> "0" Then
        
        cFirma = lookingForKundendaten(Trim(cKundnr)).firma
        cVname = lookingForKundendaten(Trim(cKundnr)).vorname
        cNName = lookingForKundendaten(Trim(cKundnr)).nachname
        cTitel = lookingForKundendaten(Trim(cKundnr)).titel
        cPlz = lookingForKundendaten(Trim(cKundnr)).Plz
        cStadt = lookingForKundendaten(Trim(cKundnr)).Ort
        cStrasse = lookingForKundendaten(Trim(cKundnr)).strasse
        cAnrede = lookingForKundendaten(Trim(cKundnr)).anrede
        
        cPrintFiTiAnVoNa = ""
        
        If cFirma <> "" Then
            cPrintFiTiAnVoNa = cFirma
        End If
        
        cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & vbCrLf
        
        If cAnrede <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cAnrede & Space(1)
        End If
        
        If cTitel <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cTitel & Space(1)
        End If
        
        If cVname <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cVname & Space(1)
        End If
        
        If cNName <> "" Then
            cPrintFiTiAnVoNa = cPrintFiTiAnVoNa & cNName
        End If
        
        sSQL = "Insert into DAGKOPF (kundnr,PrintFiTiAnVoNa,name,vorname,titel,plz,stadt,strasse,Firma,anrede,datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        sSQL = sSQL & cKundnr
        sSQL = sSQL & ", '" & cPrintFiTiAnVoNa & "' "
        sSQL = sSQL & ", '" & cNName & "' "
        sSQL = sSQL & ", '" & cVname & "' "
        sSQL = sSQL & ", '" & cTitel & "' "
        sSQL = sSQL & ", '" & cPlz & "' "
        sSQL = sSQL & ", '" & cStadt & "' "
        sSQL = sSQL & ", '" & cStrasse & "' "
        sSQL = sSQL & ", '" & cFirma & "' "
        sSQL = sSQL & ", '" & cAnrede & "' "
        sSQL = sSQL & " , '" & cZahlart & "' "
        sSQL = sSQL & " , '" & gcBediener & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
        
    Else
    
        sSQL = "Insert into DAGKOPF (datname,bedname,USTV,USTE)"
        sSQL = sSQL & " values ( "
        
        sSQL = sSQL & "  '' "
        sSQL = sSQL & " , '" & gcBediener & "' "
        sSQL = sSQL & " , '" & dUSTV & "' "
        sSQL = sSQL & " , '" & dUSTE & "' "
        
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Dim cFirmName       As String
    Dim cFirmAdress     As String
    Dim cFirmBank       As String
    Dim cFirmKomm       As String
    Dim cSteuernr       As String
    Dim cKommentar      As String
    
    If bReFuss Then
        loeschNEW "REFUSS", gdBase
        CreateTableT2 "REFUSS", gdBase
    
        cSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
            cSteuernr = ""
        End If
        
        cSQL = "Insert into REFUSS ( "
        cSQL = cSQL & " STEUERNR"
        cSQL = cSQL & ", FIRMNAME"
        cSQL = cSQL & ", FIRMADRESS"
        cSQL = cSQL & ", FIRMBANK"
        cSQL = cSQL & ", FIRMKOMM"
        cSQL = cSQL & ") values ("
        cSQL = cSQL & " '" & cSteuernr & "'"
        cSQL = cSQL & ", '" & cFirmName & "'"
        cSQL = cSQL & ", '" & cFirmAdress & "'"
        cSQL = cSQL & ", '" & cFirmBank & "'"
        cSQL = cSQL & ", '" & cFirmKomm & "'"
        cSQL = cSQL & ") "
        gdBase.Execute cSQL, dbFailOnError

    End If
    
    If gbDINA4RECHFU Then
    
        If Modul6.FindFile(gcDBPfad, "alrs3fuNettoS.rpt") Then
            reportbildschirm "alr", "alrs3fuNettoS"
        Else
            reportbildschirm "alr", "alrs3fuNetto"
        End If
    
    
        
    Else
    
    
        If Modul6.FindFile(gcDBPfad, "alrs3NettoS.rpt") Then
            reportbildschirm "alr", "alrs3NettoS"
        Else
            reportbildschirm "alr", "alrs3Netto"
        End If
    
    
    
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3134 Then
        schreibeProtokollDINA4Err "DruckeKassenZweitBonDinA4_Netto Insert into Syntaxfehler"
        schreibeProtokollDINA4Err sSQL
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeKassenZweitBonDinA4_Netto"
        Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ZeigeZweitenKassenBon(cAuswahl As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzZeile         As Long
    Dim lDatum            As Long
    Dim lWeekday          As Long
    Dim lSuchtag          As Long
    Dim lPos              As Long
    Dim lStart            As Long
    Dim lZeile            As Long
    
    Dim aDeviceName         As String
    Dim cEscapeSequenz      As String
    Dim cSQL                As String
    Dim cFeld               As String
    Dim cZeile              As String
    Dim cBonNr              As String
    Dim cUhrZeit            As String
    Dim cBetrag             As String
    Dim rsrs                As Recordset
    Dim cWoTag              As String
    Dim iCount              As Integer
    
    ReDim cDruckZeile(1 To 1) As String
    
    cBonNr = Trim$(Left(cAuswahl, 6))
    cUhrZeit = Trim(Right(cAuswahl, 8))
    cFeld = Mid(cAuswahl, 56, 10)
    lDatum = DateValue(Trim$(Mid(cAuswahl, 56, 8)))
    
    cSQL = "Select * from KASSBOND "
    cSQL = cSQL & " where DATUM = " & Trim$(Str$(lDatum)) & " "
'    cSQL = cSQL & "and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & " and BONNR = " & cBonNr & " "
    cSQL = cSQL & " and UHRZEIT = '" & cUhrZeit & "' "
    
    Set rsrs = KASSBON_DB.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BONTEXT) Then
            cFeld = rsrs!BONTEXT
        Else
            cFeld = ""
        End If
    Else
        cFeld = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If cFeld = "" Then
        Exit Sub
    End If
    
    
    lStart = 1
    
    List2.Clear
    cZeile = ""
    Do While lStart < Len(cFeld)
        lPos = InStr(lStart, cFeld, vbCrLf)
        If lPos <> 0 Then
            cZeile = Mid(cFeld, lStart, lPos - lStart + 2)
            If Right(cZeile, 2) <> vbCrLf Then
                cZeile = cZeile & vbCrLf
            End If
            lZeile = lZeile + 1
            cZeile = SwapStr(cZeile, Chr(13), "")
            cZeile = SwapStr(cZeile, Chr(10), "")
            List2.AddItem cZeile
        End If
        lStart = lPos + 2
        If lStart = 0 Then
            Exit Do
        End If
    Loop

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeZweitenKassenBon"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "ANGEBOTNOW", gdBase
    loeschNEW "DAGKOPF", gdBase
    LogtoEnd Me
    
    KASSBON_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List1_Click()
On Error GoTo LOKAL_ERROR

    ZeigeZweitenKassenBon List1.list(List1.ListIndex)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub AnlistenBonDatenNeu(cDatum As String, cFil As String, cKKart As String, bStorno As Boolean, cBon As String, cKun As String, ckassnr As String, Optional cBontext As String = "")
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum      As Long
    Dim lWeekday    As Long
    Dim lSuchtag    As Long
    Dim lcount      As Long
    
    Dim cSQL        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim cName       As String
    Dim rsrs        As Recordset
    
    List1.Clear
    List3.Clear
    
    List3.AddItem " Bon    Betrag  ZA Fil  Kunde                                      Uhrzeit"
    

    cSQL = "Select * from KASSBOND "
    cSQL = cSQL & " where KASNUM is not null "
    
    If cDatum <> "" Then
        lDatum = DateValue(cDatum)
        cSQL = cSQL & "and DATUM = " & Trim$(Str$(lDatum)) & " "
    End If
    
    If cFil <> "0" Then
        cSQL = cSQL & "and Filiale = " & Val(cFil) & " "
    End If
    If cKKart <> "0" Then
        cSQL = cSQL & "and KK_art = '" & cKKart & "' "
    End If
    
    If cBon <> "" Then
        cSQL = cSQL & "and BONNR = " & cBon & " "
    End If
    
    If cKun <> "" Then
        cSQL = cSQL & "and KUNDNR = " & cKun & " "
    End If
    
    If ckassnr <> "" Then
        cSQL = cSQL & "and kasnum = " & ckassnr & " "
    End If
    
    If cBontext <> "" Then
        cSQL = cSQL & "and Bontext like '*" & cBontext & "*' "
    End If
    
    
    If bStorno Then
        cSQL = cSQL & "and betrag < 0 "
    End If
    
    If gbUmsAnz = False Then
        cSQL = cSQL & "and BONNR >= 0 "
    End If

    cSQL = cSQL & "order by DATUM desc, UHRZEIT asc, BONNR asc "
    
    Set rsrs = KASSBON_DB.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BONNR) Then
                dWert = rsrs!BONNR
            Else
                dWert = 0
            End If
            
            
            
            cFeld = Format$(dWert, "#####0")
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cFeld
            
           
            
            If Not IsNull(rsrs!Betrag) Then
                dWert = rsrs!Betrag
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "#####0.00")
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
        
            If Not IsNull(rsrs!kk_art) Then
                cFeld = rsrs!kk_art
            Else
                cFeld = "  "
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!FILIALE) Then
                cFeld = rsrs!FILIALE
            Else
                cFeld = " "
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!Kundnr) Then
                cFeld = rsrs!Kundnr
            Else
                cFeld = " "
            End If
            
            If cFeld = "" Then cFeld = "0"
            
            If cFeld = "0" Then
                cLBSatz = cLBSatz & Space(31)
            Else
                cName = WhatIsXfromKu(rsrs!Kundnr, "Name")
                cFeld = Space$(9 - Len(cFeld)) & cFeld & " " & Left(cName, 20) & Space$(20 - Len(Left(cName, 20)))
                cLBSatz = cLBSatz & cFeld & " "
            End If
            
            If Not IsNull(rsrs!Datum) Then
                cFeld = Format(rsrs!Datum, "DD.MM.YY")
            Else
                cFeld = "        "
            End If
            cFeld = Space$(11 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!Uhrzeit) Then
                cFeld = rsrs!Uhrzeit
            Else
                cFeld = "00:00"
            End If
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            List1.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    Else
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AnlistenBonDatenNeu"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            If Len(Text1(Index)) = 0 Then
                Command5_Click 1
            End If
        Case 1
            If Len(Text1(Index)) = 4 Then
                If IsNumeric(Text1(Index)) Then
                    Command5_Click 1
                End If
            ElseIf Len(Text1(Index)) = 0 Then
                Command5_Click 1
            End If
        Case 2
            If Len(Text1(Index)) > 4 Then
                If IsNumeric(Text1(Index)) Then
                    Command5_Click 1
                End If
            ElseIf Len(Text1(Index)) = 0 Then
                Command5_Click 1
            End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 0 ' Datum
            cValid = "1234567890." & Chr$(8)
        Case 1, 2, 3  'Kundnr bonnr kasnum
            cValid = "1234567890" & Chr$(8)
        Case 4
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß#:"
    End Select

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Bonansichten ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
