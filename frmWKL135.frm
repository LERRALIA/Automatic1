VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL135 
   Caption         =   "Garantieüberprüfung"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL135.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   22
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
      Caption         =   "Auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   21
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   7
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
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
      Height          =   4380
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   11
      ToolTipText     =   "Hilfe"
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
      Index           =   0
      Left            =   9600
      TabIndex        =   8
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
      Height          =   540
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   11535
   End
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   20
      Left            =   2040
      TabIndex        =   23
      ToolTipText     =   "Kalender"
      Top             =   1080
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
      Index           =   0
      Left            =   2040
      TabIndex        =   24
      ToolTipText     =   "Kalender"
      Top             =   1680
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
   Begin sevCommand3.Command Command7 
      Height          =   165
      Index           =   1
      Left            =   1560
      TabIndex        =   25
      Top             =   1320
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
      Index           =   0
      Left            =   1560
      TabIndex        =   26
      Top             =   1080
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
   Begin sevCommand3.Command Command7 
      Height          =   165
      Index           =   3
      Left            =   1560
      TabIndex        =   27
      Top             =   1920
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
      Index           =   2
      Left            =   1560
      TabIndex        =   28
      Top             =   1680
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
      Caption         =   "Artikelbezeichnung:"
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   20
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Belegnr:"
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Kundnr:"
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Bemerkung:"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Seriennummer:"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Datum bis:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Datum von:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantieüberprüfung"
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
      TabIndex        =   10
      Top             =   120
      Width           =   10335
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
      TabIndex        =   9
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL135"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(0).SetFocus
        Case Is = 0        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
       
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 11
            gsHelpstring = "Garantieüberprüfung"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lLfnr As Long
    Dim lcount As Long
    Dim bFound As Boolean
    
    Select Case Index
        Case 0
            Unload frmWKL135
        Case 1
            zeigegarantie
        Case 2
        
            bFound = False
    
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                End If
            Next lcount
            
            If Not bFound Then
                anzeige "rot", "Bitte markieren Sie einen Artikel", Label1(4)
                Exit Sub
            End If
            
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    lLfnr = Val(Right(List2.list(lcount), 5))
                End If
            Next lcount
            
            Loeschegarantie lLfnr
            zeigegarantie
            
        Case 3
        
            bFound = False
    
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                End If
            Next lcount
            
            If Not bFound Then
                anzeige "rot", "Bitte markieren Sie einen Artikel", Label1(4)
                Exit Sub
            End If
            
            glGarantienummer = 0
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    glGarantienummer = Val(Right(List2.list(lcount), 5))
                End If
            Next lcount
            
            Command5_Click 0
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim lDat As Long

Select Case Index
    Case 0
        If IsDate(Text1(1).Text) = False Then
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
        Else
            If IsDate(Text1(1).Text) = True Then
                lDat = CLng(DateValue(Text1(1).Text))
            End If
            lDat = lDat + 1
            Text1(1).Text = Format(lDat, "DD.MM.YY")
        End If

    Case 1
        If IsDate(Text1(1).Text) = False Then
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
        Else
            If IsDate(Text1(1).Text) = True Then
                lDat = CLng(DateValue(Text1(1).Text))
            End If
            lDat = lDat - 1
            Text1(1).Text = Format(lDat, "DD.MM.YY")
        End If
    Case 2
        If IsDate(Text1(0).Text) = False Then
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
        Else
            If IsDate(Text1(0).Text) = True Then
                lDat = CLng(DateValue(Text1(0).Text))
            End If
            lDat = lDat + 1
            Text1(0).Text = Format(lDat, "DD.MM.YY")
        End If

    Case 3
        If IsDate(Text1(0).Text) = False Then
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
        Else
            If IsDate(Text1(0).Text) = True Then
                lDat = CLng(DateValue(Text1(0).Text))
            End If
            lDat = lDat - 1
            Text1(0).Text = Format(lDat, "DD.MM.YY")
        End If
End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    zeigegarantie
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigegarantie()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cVon As String
    Dim cBis As String
    Dim bAnd As Boolean
    
    bAnd = False
    
    List2.Clear
    List1.Clear
    List1.AddItem "Datum     Uhrzeit   Artnr   Artikelbezeichnung" & Space(19) & "Seriennr       Bemerkung      Kundnr     Bon"
    sSQL = "Select * from Garantie "
    
    'Bonnummer
    If Val(Text1(5).Text) > 0 Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " belegnr = " & Val(Text1(5).Text)
        bAnd = True
    End If
    
    'KUndnummer
    If Val(Text1(4).Text) > 0 Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " KUNDNR = " & Val(Text1(4).Text)
        bAnd = True
    End If
    
    'bezeich
    If Trim(Text1(6).Text) <> "" Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " bezeich like '*" & Trim(Text1(6).Text) & "*'"
        bAnd = True
    End If
    
    'serie
    If Trim(Text1(2).Text) <> "" Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " SerienNr like '*" & Trim(Text1(2).Text) & "*'"
        bAnd = True
    End If
    
    'bemerk
    If Trim(Text1(3).Text) <> "" Then
        If bAnd = True Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " bemerk like '*" & Trim(Text1(3).Text) & "*'"
        bAnd = True
    End If
    
    'von
    If Text1(1).Text <> "" Then
        If IsDate(Text1(1).Text) Then
            cVon = Trim(Text1(1).Text)
            lVon = DateValue(cVon)
            
            If bAnd = True Then
                sSQL = sSQL & " And "
            Else
                sSQL = sSQL & " where "
            End If
            
            sSQL = sSQL & " adate >= " & lVon
            bAnd = True
        End If
    End If
    
    'bis
    If Text1(0).Text <> "" Then
        If IsDate(Text1(0).Text) Then
            cBis = Trim(Text1(0).Text)
            lBis = DateValue(cBis)
            
            If bAnd = True Then
                sSQL = sSQL & " And "
            Else
                sSQL = sSQL & " where "
            End If
            
            
            sSQL = sSQL & " adate <= " & lBis
            bAnd = True
        End If
    End If
    
    
    sSQL = sSQL & " order by lfnr desc "
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cSatz = ""
            
            If Not IsNull(rsrs!Adate) Then
                cFeld = Format(rsrs!Adate, "DD.MM.YY") & Space(2)
            Else
                cFeld = Space(10)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!AZEIT) Then
                cFeld = Format(rsrs!AZEIT, "HH:MM:SS") & Space(2)
            Else
                cFeld = Space(10)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!artnr) Then
                cFeld = Space(6 - Len(rsrs!artnr)) & rsrs!artnr & Space(2)
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH & Space(35 - Len(rsrs!BEZEICH)) & Space(2)
            Else
                cFeld = Space(37)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!Seriennr) Then
                If Len(rsrs!Seriennr) > 12 Then
                    cFeld = Left(rsrs!Seriennr, 9) & "..."
                    cFeld = cFeld & Space(15 - Len(cFeld))
                Else
                    cFeld = rsrs!Seriennr & Space(13 - Len(rsrs!Seriennr)) & Space(2)
                End If
                
            Else
                cFeld = Space(15)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!Bemerk) Then
                If Len(rsrs!Bemerk) > 12 Then
                    cFeld = Left(rsrs!Bemerk, 9) & "..."
                    cFeld = cFeld & Space(15 - Len(cFeld))
                Else
                    cFeld = rsrs!Bemerk & Space(13 - Len(rsrs!Bemerk)) & Space(2)
                End If
                
            Else
                cFeld = Space(15)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!Kundnr) Then
                cFeld = Space(6 - Len(rsrs!Kundnr)) & rsrs!Kundnr & Space(2)
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!BELEGNR) Then
                cFeld = Space(6 - Len(rsrs!BELEGNR)) & rsrs!BELEGNR & Space(2)
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!lfnr) Then
                cFeld = Space(6 - Len(rsrs!lfnr)) & rsrs!lfnr & Space(2)
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & cFeld
            
            List2.AddItem cSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigegarantie"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Loeschegarantie(lLaufendNr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    sSQL = "Delete from Garantie where lfnr = " & lLaufendNr
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Loeschegarantie"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
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
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 2, 6, 3 'Bemerkung seriennr bezeich
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%"
        Case 1, 0 ' Datum
            cValid = "1234567890." & Chr$(8)
        Case 4, 5  'Kundnr bonnr
            cValid = "1234567890" & Chr$(8)
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
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command5_Click 1
    End If
    
    If KeyCode = vbKeyEscape Then
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Text1(Index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Garantieüberprüfung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
