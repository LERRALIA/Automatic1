VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL115 
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Auswertung"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   615
      Left            =   9600
      TabIndex        =   29
      Top             =   2160
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Lieferanten"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Marken"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "auch Bediener ohne Nein Verkäufe anzeigen"
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      Top             =   3480
      Width           =   2175
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   23
      Top             =   3960
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Bedieneraktivität"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   22
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
      Caption         =   "Entfernen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   11
      Left            =   10800
      TabIndex        =   19
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   9600
      TabIndex        =   16
      Top             =   6120
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "als Packliste"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "als Bestellvorschlag"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Entwicklung"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   12
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Ändern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   9
      Top             =   7320
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
         Height          =   6615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   11668
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorFixed  =   12632256
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   2
      Top             =   1080
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
      Caption         =   "Suche Daten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   1
      Top             =   7920
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
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ComboBox cboBed 
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
      Left            =   120
      TabIndex        =   4
      Text            =   "alle"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3840
      TabIndex        =   21
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1080
      TabIndex        =   25
      Tag             =   "2"
      Top             =   2355
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   1875
      Width           =   1215
   End
   Begin sevCommand3.Command Command11 
      Height          =   360
      Left            =   11280
      TabIndex        =   32
      Top             =   360
      Width           =   405
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
      ToolTipTitle    =   "Spaltenanordung"
      ButtonStyle     =   2
      Caption         =   ""
      Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
      Picture         =   "frmWKL115.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   14
      Left            =   2760
      TabIndex        =   33
      ToolTipText     =   "Kalender"
      Top             =   1845
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command7 
      Height          =   165
      Index           =   1
      Left            =   2400
      TabIndex        =   34
      Top             =   2040
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
      Index           =   0
      Left            =   2400
      TabIndex        =   35
      Top             =   1845
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
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   15
      Left            =   2760
      TabIndex        =   36
      ToolTipText     =   "Kalender"
      Top             =   2325
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command7 
      Height          =   165
      Index           =   3
      Left            =   2400
      TabIndex        =   37
      Top             =   2520
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
      Index           =   2
      Left            =   2400
      TabIndex        =   38
      Top             =   2325
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
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Bediener"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Grund des Nein Verkaufs"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblAnzeige 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Auswertung Nein Verkäufe"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
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
      Left            =   120
      TabIndex        =   28
      Top             =   1920
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
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmWKL115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerArtnr          As Byte
Dim Spaltennummerlfnr           As Byte
Dim SpaltennummerArtikelStatus  As Byte
Dim SpaltennummerKUNDNR         As Byte
Dim SpaltennummerBESTELLTAM     As Byte
Dim SpaltennummerBESTELLTUM     As Byte
Dim SpaltennummerAWM            As Byte
Dim SpaltennummerAWMKund        As Byte
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 14       ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(1).SetFocus
        Case Is = 15        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
            vorbereitungTab
            ZeigDaten " order by Filiale"
            zeigKnöpfe True
        Case Is = 1
            Unload frmWKL115
        Case Is = 2
            
            Frame1.Visible = False
            zeigKnöpfe False
            anzeige "normal", "", lblanzeige
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigDaten(corder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer

    Tabcheck "NVK"
    FormatGridOverTablay "NVK"

    With MSHFLEX1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With
    
    'Grid fuellen
    anzeige "normal", "Die Daten werden angezeigt...", lblanzeige
    
    GridFuellen ermittleDaten(corder)
    
    ermittlespalten
    
    FaerbenHGrid MSHFLEX1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    
    If MSHFLEX1.Visible Then
        MSHFLEX1.Col = 0
        MSHFLEX1.Row = 0
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigDaten"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "LFNR"
                Spaltennummerlfnr = i
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "KUNDNR"
                SpaltennummerKUNDNR = i
            Case Is = "ADATE"
                SpaltennummerBESTELLTAM = i
            Case Is = "AZEIT"
                SpaltennummerBESTELLTUM = i
            Case Is = "FARBE"
                SpaltennummerAWM = i
            Case Is = "FARBNRKU"
                SpaltennummerAWMKund = i
            Case Is = "STATUSARTIKEL"
                SpaltennummerArtikelStatus = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermittleDaten(corder As String) As String
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim iBed            As Integer
    Dim iFil            As Integer
    Dim bAnd            As Boolean
    Dim cSTATUS         As String
    Dim lcount          As Long
    Dim lVon            As Long
    Dim lBis            As Long
    Dim cFeld           As String
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) Then
            cSTATUS = Trim(List1.list(lcount))
        End If
    Next lcount
            
    ermittleDaten = "Select * from KC" & srechnertab
    
    iBed = 0
    iFil = 0
    bAnd = False
    
    cFeld = Text1(0).Text
    If cFeld <> "" Then
        If IsDate(cFeld) Then
            lVon = DateValue(cFeld)
        Else
            anzeige "rot", "Bitte ein gültiges Datum eingeben!", lblanzeige
            Text1(0).SetFocus
            Exit Function
        End If
    Else
        lVon = 0
    End If
    
    cFeld = Text1(1).Text
    If cFeld <> "" Then
        If IsDate(cFeld) Then
            lBis = DateValue(cFeld)
        Else
            anzeige "rot", "Bitte ein gültiges Datum eingeben!", lblanzeige
            Text1(1).SetFocus
            Exit Function
        End If
    Else
        lBis = 0
    End If
    

    If cboBed.Text <> "alle Bediener" Then
        iBed = CInt(Left$(cboBed.Text, 4))
    End If
    
'    If cboFil.Text <> "alle Filialen" Then
'        iFil = CInt(Left$(cboFil.Text, 3))
'    End If
'
    sSQL = "Select * from KC" & srechnertab
    
    If cSTATUS <> "" Then
        If bAnd Then
            sSQL = sSQL & " and neinart = '" & cSTATUS & "'"
        Else
            sSQL = sSQL & " where neinart = '" & cSTATUS & "'"
        End If
        bAnd = True
    End If
    
'    If iFil <> "0" Then
'        If bAnd Then
'            sSQL = sSQL & " and FILIALE = " & iFil
'        Else
'            sSQL = sSQL & " where FILIALE = " & iFil
'        End If
'        bAnd = True
'    End If
    
    If iBed <> "0" Then
        If bAnd Then
            sSQL = sSQL & " and BEDNU = " & iBed
        Else
            sSQL = sSQL & " where BEDNU = " & iBed
        End If
        bAnd = True
    End If
    
    If lVon > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        
        sSQL = sSQL & " ADATE >= " & Trim$(Str$(lVon)) & " "
        
        
        If lBis > 0 Then
            sSQL = sSQL & " and ADATE <= " & Trim$(Str$(lBis)) & " "
        Else
            sSQL = sSQL & " and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
        End If
        bAnd = True
    Else
        If lBis > 0 Then
        
            If bAnd Then
                sSQL = sSQL & " and "
            Else
                sSQL = sSQL & " where "
            End If
            sSQL = sSQL & " ADATE >= " & Trim$(Str$(lBis)) & " "
            sSQL = sSQL & " and ADATE <= " & Trim$(Str$(lBis)) & " "
            bAnd = True
        End If
        
    End If
    
    ermittleDaten = sSQL & corder

    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleDaten"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Long
    Dim i           As Long
    Dim j           As Long
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GridFuellen(cSQL As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim lMax        As Long
    Dim sSQL2       As String
    
    loeschNEW "DRUCK133", gdBase
    CreateTableT2 "DRUCK133", gdBase
    
    sSQL2 = "Insert into DRUCK133 " & cSQL
    gdBase.Execute sSQL2, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSHFLEX1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        anzeige "normal", "Es werden " & lMax & " Artikel angezeigt...", lblanzeige
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            .Rows = lrow + 1
            .Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    Select Case sSpaltenname(i)
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                    End Select
                    
                End If
            Next i
            rsrs.MoveNext
        Loop
        
        Frame1.Visible = True
        anzeige "normal", "Es wurden " & lMax & " Artikel ermittelt.", lblanzeige
    Else
        Frame1.Visible = False
        anzeige "rot", "Es wurden keine Artikel ermittelt.", lblanzeige
    End If
        
    rsrs.Close
    If byAnzahlSpalten < 2 Then
    Else
        .FixedCols = 1
    End If
    .RowHeight(1) = 0
    lrow = lrow - 1
    
    .Redraw = True
    .Visible = True
    
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "lfnr"
    gstab = "NVK"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    
    Select Case Index
    
    Case 0 'Artikeldaten
        EntwicklungNeinVK
    Case 1
'        ZeigDaten ""
        If Option1(0).Value = True Then
            drucken "Packliste"
        ElseIf Option1(1).Value = True Then
            drucken "Bestellvorschlag"
        End If
    Case 2 'Bedieneraktivität
        AUSwertung lblanzeige
    Case 3
        If MSHFLEX1.RowSel > 1 Then
            FlexGrid_Update MSHFLEX1
            
        Else
            anzeige "rot", "Markieren Sie eine Zeile!", lblanzeige
        End If
    Case 4 'Auswertung
        AuswertungNeinVK lblanzeige
    Case 5
        leeren
    Case 6 'sele
        If MSHFLEX1.RowSel > 1 Then
            FlexGrid_Delete MSHFLEX1
            
            ZeigDaten ""
        Else
            anzeige "rot", "Markieren Sie eine Zeile!", lblanzeige
        End If
    Case 11
        gsHelpstring = "Auswertung Nein Verkäufe"
        frmWKL110.Show 1

        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function AUSwertung(lblx As Label) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnzNeinVK As Long
    Dim dErtrag As Double
    Dim dKassErtrag As Double
    Dim sdat As String
    
    loeschNEW "MITNVK", gdBase
    CreateTableT2 "MITNVK", gdBase
    
    loeschNEW "PRINVK", gdBase
    CreateTableT2 "PRINVK", gdBase
    
    loeschNEW "KassBed", gdBase
    CreateTableT2 "KASSBED", gdBase
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", lblx
    
    sSQL = "insert Into KassBed select artnr, preis,mwst,bediener,adate,linr,ekpr,menge  from kassjour where adate > " & CLng(DateValue(Now) - 100)
    sSQL = sSQL & " and UMS_OK = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Erträge werden errechnet(mwst = V)...", lblx
    
    sSQL = "Update KassBed set ertrag = ((preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Erträge werden errechnet(mwst = E)...", lblx
    
    sSQL = "Update KassBed set ertrag = ((preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Erträge werden errechnet(mwst = O)...", lblx
    
    sSQL = "Update KassBed set ertrag = ((preis * 100)/(100 + " & gdMWStO & ")) - (EKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Index (Bediener) wird erstellt...", lblx
    CheckIndex "KassBed", "Bediener", "", gdBase
    anzeige "normal", "Index (Datum) wird erstellt...", lblx
    CheckIndex "KassBed", "Adate", "", gdBase
    
    sSQL = "Insert Into MITNVK select bednu,bedname from bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MITNVK")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                lAnzNeinVK = ermNVKproMit(rsrs!BEDNU, 2)
                lblx.Caption = rsrs!bedname & " " & lAnzNeinVK & " Neinverkäufe"
                lblx.Refresh
                rsrs.Edit
                rsrs!MengeNVK = lAnzNeinVK
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If Check1.Value = vbUnchecked Then
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK where MengeNVK > 0 order by MengeNVK desc"
    Else
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK  order by MengeNVK desc"
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
            
                lblx.Caption = rsrs!bedname & " " & rsrs!MengeNVK & " Neinverkäufe"
                lblx.Refresh
                
                dErtrag = ermNVKErtragproMit(rsrs!BEDNU, 2)
                dKassErtrag = ermErtragproMit(rsrs!BEDNU, 2)
                rsrs.Edit
                rsrs!ERTRAG = dErtrag
                rsrs!KassErtrag = dKassErtrag
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sdat = Format(DateValue(Now), "mmmm") & " " & Format(DateValue(Now), "YYYY")
    sSQL = "Insert Into PRINvk select bednu,bedname,MengeNVK,ertrag,KassErtrag,'" & sdat & "' as monat from MITnvk "
    gdBase.Execute sSQL, dbFailOnError
    
    'Jetzt vormonat
    
    loeschNEW "MITNVK", gdBase
    CreateTableT2 "MITNVK", gdBase
    
    sSQL = "Insert Into MITNVK select bednu,bedname from bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MITNVK")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                lAnzNeinVK = ermNVKproMit(rsrs!BEDNU, 1)
                lblx.Caption = rsrs!bedname & " " & lAnzNeinVK & " Neinverkäufe"
                lblx.Refresh
                rsrs.Edit
                rsrs!MengeNVK = lAnzNeinVK
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If Check1.Value = vbUnchecked Then
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK where MengeNVK > 0 order by MengeNVK desc"
    Else
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK  order by MengeNVK desc"
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
            
                lblx.Caption = rsrs!bedname & " " & rsrs!MengeNVK & " Neinverkäufe"
                lblx.Refresh
                
                dErtrag = ermNVKErtragproMit(rsrs!BEDNU, 1)
                dKassErtrag = ermErtragproMit(rsrs!BEDNU, 1)
                rsrs.Edit
                rsrs!ERTRAG = dErtrag
                rsrs!KassErtrag = dKassErtrag
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If Month(DateValue(Now)) = 1 Then
        sdat = MonthName(12) & " " & Year(Now) - 1
    Else
        sdat = MonthName(Month(DateValue(Now)) - 1) & " " & Format(DateValue(Now), "YYYY")
    End If
    
    sSQL = "Insert Into PRINVK select bednu,bedname,MengeNVK,ertrag,KassErtrag,'" & sdat & "' as monat from MITnvk "
    gdBase.Execute sSQL, dbFailOnError
    
    'Jetzt vorvormonat
    
    loeschNEW "MITNVK", gdBase
    CreateTableT2 "MITNVK", gdBase
    
    sSQL = "Insert Into MITNVK select bednu,bedname from bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MITNVK")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                lAnzNeinVK = ermNVKproMit(rsrs!BEDNU, 3)
                lblx.Caption = rsrs!bedname & " " & lAnzNeinVK & " Neinverkäufe"
                lblx.Refresh
                rsrs.Edit
                rsrs!MengeNVK = lAnzNeinVK
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    If Check1.Value = vbUnchecked Then
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK where MengeNVK > 0 order by MengeNVK desc"
    Else
        sSQL = "select MengeNVK, bednu,bedname,ertrag,KassErtrag from MITNVK  order by MengeNVK desc"
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
            
                lblx.Caption = rsrs!bedname & " " & rsrs!MengeNVK & " Neinverkäufe"
                lblx.Refresh
                
                dErtrag = ermNVKErtragproMit(rsrs!BEDNU, 3)
                dKassErtrag = ermErtragproMit(rsrs!BEDNU, 3)
                rsrs.Edit
                rsrs!ERTRAG = dErtrag
                rsrs!KassErtrag = dKassErtrag
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    
    If Month(DateValue(Now)) = 2 Then
        sdat = MonthName(12) & " " & Year(Now) - 1
    ElseIf Month(DateValue(Now)) = 1 Then
        sdat = MonthName(11) & " " & Year(Now) - 1
    Else
        sdat = MonthName(Month(DateValue(Now)) - 2) & " " & Format(DateValue(Now), "YYYY")
    End If
    
    
    sSQL = "Insert Into PRINVK select bednu,bedname,MengeNVK,ertrag,Kassertrag,'" & sdat & "' as monat from MItnvk "
    gdBase.Execute sSQL, dbFailOnError
    
    If Check1.Value = vbUnchecked Then
        sSQL = "Delete from PRINVK where MengeNVK = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblx
    loeschNEW "KassBed", gdBase
    
    reportbildschirm "", "aZEN133e"
    Screen.MousePointer = 0
    anzeige "normal", "", lblx
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Auswertung"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function AuswertungNeinVK(lblx As Label) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bAnd As Boolean
    
    bAnd = False
    
    loeschNEW "NVKAUSW", gdBase
    CreateTableT2 "NVKAUSW", gdBase
    
    Screen.MousePointer = 0
    anzeige "normal", "Schritt 1", lblx
 
    sSQL = "Insert Into NVKAUSW select * from NEINVK"
'    If cboFil.Text = "alle Filialen" Then
'
'    Else
'        sSQL = sSQL & " where Filiale  = " & CInt(Left$(cboFil.Text, 3))
'        bAnd = True
'    End If
    
    If cboBed.Text = "alle Bediener" Then
        
    Else
        If bAnd Then
            sSQL = sSQL & " and "
        Else
            sSQL = sSQL & " where "
        End If
        sSQL = sSQL & " bednu  = " & CInt(Left$(cboBed.Text, 4))
        bAnd = True
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 2", lblx
    
    sSQL = "Update NVKAUSW inner join ARTIKEL on NVKAUSW.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " set NVKAUSW.LINR = Artikel.linr "
    sSQL = sSQL & " , NVKAUSW.lpz = Artikel.lpz "
    sSQL = sSQL & " , NVKAUSW.KVKPR1 = Artikel.KVKPR1 "
    sSQL = sSQL & " , NVKAUSW.MWST = Artikel.MWST "
    sSQL = sSQL & " , NVKAUSW.FARBNR = val(Artikel.awm) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 3", lblx
    sSQL = "Update NVKAUSW set filauswahl = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update NVKAUSW set Bediener = '" & cboBed.Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 4", lblx
    
    sSQL = "Update NVKAUSW inner join Artlief on NVKAUSW.linr = Artlief.linr "
    sSQL = sSQL & " and NVKAUSW.artnr = Artlief.artnr "
    sSQL = sSQL & " Set NVKAUSW.LEKPR = Artlief.LEKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    
    anzeige "normal", "Schritt 5", lblx
    sSQL = "Update NVKAUSW Set Ertrag = (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStV & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update NVKAUSW Set Ertrag =  (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStE & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update NVKAUSW Set Ertrag =  (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStO & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 6", lblx
    BringFarbeInsSpiel "NVKAUSW", gdBase
     
    anzeige "normal", "Schritt 7", lblx
    sSQL = "Update NVKAUSW inner join LINBEZ on NVKAUSW.linr = LINBEZ.linr and NVKAUSW.lpz = LINBEZ.lpz "
    sSQL = sSQL & " Set NVKAUSW.marke = LINBEZ.Marke "
    sSQL = sSQL & " , NVKAUSW.markeKUERZEL = LINBEZ.KUERZEL "
    sSQL = sSQL & " , NVKAUSW.LINBEZEICH = LINBEZ.LINBEZEICH"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update NVKAUSW inner join LISRT on NVKAUSW.linr = LISRT.linr  "
    sSQL = sSQL & " Set NVKAUSW.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblx
    
    If Option1(2).Value = True Then
        reportbildschirm "", "aZEN133d"
    ElseIf Option1(3).Value = True Then
        reportbildschirm "", "aZEN133f"
    End If
    
    Screen.MousePointer = 0
    anzeige "normal", "", lblx
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AuswertungNeinVK"
    Fehler.gsFehlertext = "Im Programmteil Auswertung Nein Verkäufe ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub FlexGrid_Delete(oGrid As MSHFlexGrid)
On Error GoTo LOKAL_ERROR

  Dim nRow As Long
  Dim nCol As Long
  Dim nRowSel As Long
  Dim nColSel As Long
  Dim nDelRow As Long
  Dim lBig As Long
  

  With oGrid
    ' aktuelle Selektion merken
    
    nRow = .Row
    nCol = .Col
    nRowSel = .RowSel
    nColSel = .ColSel
    
    
    If nRow > nRowSel Then
        lBig = nRow
        nDelRow = nRowSel - 1
    Else
        lBig = nRowSel
        nDelRow = nRow - 1
    End If
    
    
    Do While nDelRow < lBig
    
        nDelRow = nDelRow + 1
        
        If nDelRow > 1 Then
            loescheausKUNDBEST nDelRow
        End If
    Loop
    
    
    
    
    
'    nRow = .Row
'    nCol = .Col
'    nRowSel = .RowSel
'    nColSel = .ColSel
'
'    nDelRow = nRow - 1
'    Do While nDelRow < nRowSel
'
'        nDelRow = nDelRow + 1
'        loescheausKUNDBEST nDelRow
'
'
'    Loop
    
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Delete"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 2
            If MSHFLEX1.RowSel > 1 Then
                FlexGrid_Update MSHFLEX1
                ZeigDaten " order by Filiale"
            Else
                
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FlexGrid_Update(oGrid As MSHFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    Dim sSQL As String
    
    Dim lartnr      As Long
    Dim czeit       As String
    
  With oGrid
    ' aktuelle Selektion merken
    nRow = .Row
    nCol = .Col
    nRowSel = .RowSel
    nColSel = .ColSel
    
    If nRow > nRowSel Then
        lBig = nRow
        nDelRow = nRowSel - 1
    Else
        lBig = nRowSel
        nDelRow = nRow - 1
    End If
    
    Do While nDelRow < lBig
        nDelRow = nDelRow + 1
        If nDelRow > 1 Then
            lartnr = -1
            
            lartnr = .TextMatrix(nDelRow, SpaltennummerArtnr)
            czeit = .TextMatrix(nDelRow, SpaltennummerBESTELLTUM)
            
            If lartnr <> -1 Then
                    
                sSQL = "Update NEINVK inner join KC" & srechnertab & " on "
                sSQL = sSQL & " NEINVK.ARTNR = KC" & srechnertab & ".ARTNR "
                sSQL = sSQL & " set NEINVK.anzei = 0 "
                sSQL = sSQL & " where NEINVK.ARTNR = " & lartnr
                sSQL = sSQL & " and NEINVK.AZEIT = '" & czeit & "'"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Delete from KC" & srechnertab & "  where ARTNR = " & lartnr
                sSQL = sSQL & " and AZEIT = '" & czeit & "'"
                gdBase.Execute sSQL, dbFailOnError
                
            End If
            
        End If
    Loop
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            DatumRauf Text1(0), "lang"
        Case 1
            DatumRunter Text1(0), "lang"
        Case 2
            DatumRauf Text1(1), "lang"
        Case 3
            DatumRunter Text1(1), "lang"
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    WKL94Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    LogtoStart Me
    
    fuelleBediener cboBed

    fuellecombo1
   
    zeigKnöpfe False
    
    anzeige "normal", "", lblanzeige
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermNVKErtragproMit(sBednu As String, iTage As Integer) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim sSQL1 As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    
    ermNVKErtragproMit = 0
    
    loeschNEW "Aktivi", gdBase
    CreateTableT2 "AKTIVI", gdBase
    
    sSQL = "Insert into Aktivi select artnr , menge from neinvk where bednu = " & sBednu
    If iTage = 1 Then
    'vor Monat
        If Month(Now) = 1 Then
            sSQL = sSQL & " and Month(adate) = 12 and Year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and Month(adate) = " & Month(Now) - 1 & " and Year(adate) = " & Year(Now)
        End If
    ElseIf iTage = 2 Then
    'akt Monat
        sSQL = sSQL & " and Month(adate) = " & Month(Now) & " and Year(adate) = " & Year(Now)
    ElseIf iTage = 3 Then
    'vor Monat
        If Month(Now) = 2 Then
            sSQL = sSQL & " and Month(adate) = 12 and Year(adate) = " & Year(Now) - 1
        ElseIf Month(Now) = 1 Then
            sSQL = sSQL & " and Month(adate) = 11 and Year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and Month(adate) = " & Month(Now) - 2 & " and Year(adate) = " & Year(Now)
        End If
    Else
        sSQL = sSQL & " and adate >= " & CLng(DateValue(Now) - 14)
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Aktivi inner join ARTIKEL on Aktivi.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " set Aktivi.LINR = Artikel.linr "
    sSQL = sSQL & " , Aktivi.KVKPR1 = Artikel.KVKPR1 "
    sSQL = sSQL & " , Aktivi.MWST = Artikel.MWST "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Aktivi inner join Artlief on Aktivi.linr = Artlief.linr "
    sSQL = sSQL & " and Aktivi.artnr = Artlief.artnr "
    sSQL = sSQL & " Set Aktivi.LEKPR = Artlief.LEKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Aktivi set ertrag = (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStV & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Aktivi set ertrag = (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStE & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Aktivi set ertrag = (((KVKPR1*MENGE) * 100)/(100 + " & gdMWStO & ")) - (LEKPR * MENGE)  "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
   
    sSQL = "select sum(ertrag) as maxi from AKTIVI"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!maxi) Then
            ermNVKErtragproMit = CDbl((rsrs!maxi))
        End If
    
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNVKErtragproMit"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermErtragproMit(sBednu As String, iTage As Integer) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim sSQL1 As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    
    ermErtragproMit = 0
    
    sSQL = "select sum(ertrag) as maxi from Kassbed where bediener = " & sBednu
    If iTage = 1 Then
    'vor Monat
        If Month(Now) = 1 Then
            sSQL = sSQL & " and Month(adate) = 12 and Year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and Month(adate) = " & Month(Now) - 1 & " and Year(adate) = " & Year(Now)
        End If
    ElseIf iTage = 2 Then
    'akt Monat
        sSQL = sSQL & " and Month(adate) = " & Month(Now) & " and Year(adate) = " & Year(Now)
    ElseIf iTage = 3 Then
    'vor Monat
        If Month(Now) = 2 Then
            sSQL = sSQL & " and Month(adate) = 12 and Year(adate) = " & Year(Now) - 1
        ElseIf Month(Now) = 1 Then
            sSQL = sSQL & " and Month(adate) = 11 and Year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and Month(adate) = " & Month(Now) - 2 & " and Year(adate) = " & Year(Now)
        End If
    Else
        sSQL = sSQL & " and adate >= " & CLng(DateValue(Now) - 14)
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!maxi) Then
            ermErtragproMit = CDbl((rsrs!maxi))
        End If
    
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermErtragproMit"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermNVKproMit(sBednu As String, iTage As Integer) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    
    ermNVKproMit = 0
    
    sSQL = "select sum(menge) as maxi from neinvk where bednu = " & sBednu
    
    If iTage = 1 Then
    'vor Monat
        If Month(Now) = 1 Then
            sSQL = sSQL & " and month(adate) = 12 and year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and month(adate) = " & Month(Now) - 1 & " and year(adate) = " & Year(Now)
        End If
    
    ElseIf iTage = 2 Then
    'akt Monat
        sSQL = sSQL & " and month(adate) = " & Month(Now) & " and year(adate) = " & Year(Now)
    ElseIf iTage = 3 Then
    'vorvor Monat
        If Month(Now) = 2 Then
            sSQL = sSQL & " and month(adate) = 12 and year(adate) = " & Year(Now) - 1
        ElseIf Month(Now) = 1 Then
            sSQL = sSQL & " and month(adate) = 11 and year(adate) = " & Year(Now) - 1
        Else
            sSQL = sSQL & " and month(adate) = " & Month(Now) - 2 & " and year(adate) = " & Year(Now)
        End If
    Else
        sSQL = sSQL & " and adate >= " & CLng(DateValue(Now) - 14)
    End If
   
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermNVKproMit = Val(rsrs!maxi)
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNVKproMit"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub zeigKnöpfe(bWie As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Command3(1).Enabled = bWie
    Frame2.Visible = bWie
    Command1(2).Visible = bWie
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigKnöpfe"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungTab()
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim iFil            As Integer
    Dim cART            As String
    Dim rsrs            As Recordset
    Dim datLVK          As Long
    Dim datERST         As Long
    Dim lUw             As Long
    Dim inBe            As Long
    
    Screen.MousePointer = 11
    
    loeschNEW "KC" & srechnertab, gdBase
    CreateTableT2 "KC" & srechnertab, gdBase
    
    sSQL = "Update NEINVK set anzei = 0 where adate < Datevalue(now) - 90 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KC" & srechnertab & " Select * from NEINVK where anzei = 1"
    sSQL = sSQL & " or anzei is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index artnr on KC" & srechnertab & " (artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KC" & srechnertab & " inner join artikel on  KC" & srechnertab & ".artnr = artikel.artnr"
    sSQL = sSQL & " set KC" & srechnertab & ".RKZ = artikel.RKZ "
    sSQL = sSQL & " , KC" & srechnertab & ".farbe = val(artikel.aWM) "
    sSQL = sSQL & " , KC" & srechnertab & ".farbnr = val(artikel.aWM) "
    sSQL = sSQL & " , KC" & srechnertab & ".Linr = artikel.linr "
    sSQL = sSQL & " , KC" & srechnertab & ".KVKPR1 = artikel.KVKPR1 "
    sSQL = sSQL & " , KC" & srechnertab & ".Libesnr = artikel.Libesnr "
    sSQL = sSQL & " , KC" & srechnertab & ".BESTAND = artikel.BESTAND "
    sSQL = sSQL & " , KC" & srechnertab & ".mb = artikel.minbest "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update KC" & srechnertab & " inner join zbestand on  KC" & srechnertab & ".artnr = zbestand.artnr"
'    sSQL = sSQL & " and KC" & srechnertab & ".FILIALE = zbestand.filialnr "
'    sSQL = sSQL & " set KC" & srechnertab & ".mb = zbestand.minbest "
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KC" & srechnertab & " inner join lisrt on  KC" & srechnertab & ".LINR = lisrt.linr"
    sSQL = sSQL & " set KC" & srechnertab & ".liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from KC" & srechnertab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                iFil = rsrs!FILIALE
                
                datLVK = ErmlzVKproFil(cART, iFil)
                inBe = erminBestell(cART)
                lUw = 0 'LeseUnterwegs(CLng(cART), CLng(iFil))
                datERST = ErmFirstZugang(cART)
                
                rsrs.Edit
                rsrs!ERSTDAT = datERST
                rsrs!lastvk = datLVK
                rsrs!inBe = inBe
                rsrs!UW = lUw
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    sSQL = "Update KC" & srechnertab & " inner join MBORDER on KC" & srechnertab & ".Artnr = MBORDER.Artnr "
    sSQL = sSQL & " set KC" & srechnertab & ".BLOCK = 'B' "
    sSQL = sSQL & " where MBORDER.FILIALE = KC" & srechnertab & ".FILIALE"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KC" & srechnertab & " set NEU = 'N' where erstdat > datevalue(now) - 90 "
    gdBase.Execute sSQL, dbFailOnError
 
    
    sSQL = "Update KC" & srechnertab & " set BESTANDO = BESTAND"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KC" & srechnertab & " set BESTANDO = 0 where BESTANDO is null "
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungTab"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EntwicklungNeinVK()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim i       As Integer
    
    Screen.MousePointer = 11

    For i = 0 To 6
        loeschNEW "AAT" & i, gdBase
        sSQL = "Select distinct adate , sum(Menge) as AMenge,Filiale into AAT" & i
        sSQL = sSQL & " from Neinvk "
        
        If Month(DateValue(Now)) - i = 0 Then
            sSQL = sSQL & " where month(adate) = 12 "
            
        ElseIf Month(DateValue(Now)) - i = -1 Then
            sSQL = sSQL & " where month(adate) = 11 "
            
        ElseIf Month(DateValue(Now)) - i = -2 Then
            sSQL = sSQL & " where month(adate) = 10 "
            
        ElseIf Month(DateValue(Now)) - i = -3 Then
            sSQL = sSQL & " where month(adate) = 9 "
        ElseIf Month(DateValue(Now)) - i = -4 Then
            sSQL = sSQL & " where month(adate) = 8 "
        ElseIf Month(DateValue(Now)) - i = -5 Then
            sSQL = sSQL & " where month(adate) = 7 "
        ElseIf Month(DateValue(Now)) - i = -6 Then
            sSQL = sSQL & " where month(adate) = 6 "
        Else
            sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - i
        End If
        
        If Month(DateValue(Now)) - i = 0 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -1 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -2 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -3 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -4 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -5 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -6 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        Else
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now))
        End If
        sSQL = sSQL & " group by adate, Filiale"
        gdBase.Execute sSQL, dbFailOnError
    Next i

    For i = 0 To 6
        loeschNEW "TOPNEINT" & i, gdBase

        sSQL = "Select sum(AMenge) as SMENGE ,Filiale into TOPNEINT" & i
        sSQL = sSQL & " from AAT" & i
        sSQL = sSQL & " group by Filiale"
        gdBase.Execute sSQL, dbFailOnError
    Next i


    For i = 0 To 6
        loeschNEW "TOPNEIN" & i, gdBase
        CreateTable "TOPNEIN" & i, gdBase

        sSQL = "Insert into TOPNEIN" & i & " Select SMENGE ,Filiale from TOPNEINT" & i
        gdBase.Execute sSQL, dbFailOnError
    Next i

    loeschNEW "NEIN5", gdBase
    CreateTable "NEIN5", gdBase

    For i = 1 To giAnzFil
        sSQL = "Insert into NEIN5 (Filiale) values ( " & i & ") "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 0 To 6
        sSQL = "Update NEIN5 inner join TOPNEIN" & i & " on NEIN5.Filiale = TOPNEIN" & i & ".Filiale "
        sSQL = sSQL & " SET  NEIN5.tMenge" & i & " = TOPNEIN" & i & ".sMenge "
        gdBase.Execute sSQL, dbFailOnError
    Next i

    sSQL = "Update NEIN5 inner join Filialen on NEIN5.Filiale = Filialen.Filialnr "
    sSQL = sSQL & " SET NEIN5.filNAME = Filialen.Filialname "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update NEIN5  SET KUCUTS = (tMenge1 + tMenge2 + tMenge3 + tMenge4 + tMenge5 + tMenge6)/6 "
    gdBase.Execute sSQL, dbFailOnError

    Dim cMontName

    For i = 0 To 6
        If Month(DateValue(Now)) - i = 0 Then
            cMontName = "Dezember"
        ElseIf Month(DateValue(Now)) - i = -1 Then
            cMontName = "November"
        ElseIf Month(DateValue(Now)) - i = -2 Then
            cMontName = "Oktober"
        ElseIf Month(DateValue(Now)) - i = -3 Then
            cMontName = "September"
        ElseIf Month(DateValue(Now)) - i = -4 Then
            cMontName = "August"
        ElseIf Month(DateValue(Now)) - i = -5 Then
            cMontName = "Juli"
        ElseIf Month(DateValue(Now)) - i = -6 Then
            cMontName = "Juni"
        Else
            cMontName = MonthName(Month(DateValue(Now)) - i)
        End If

        sSQL = "Update NEIN5  SET mont" & i & " = '" & cMontName & "'"
        gdBase.Execute sSQL, dbFailOnError
    Next i

    loeschNEW "NEIN6", gdBase
    sSQL = "select * into NEIN6 from NEIN5 order by KUCUTS desc"
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
    reportbildschirm "", "aZEN133c"
    anzeige "normal", "", lblanzeige
   
    loeschNEW "NEIN5", gdBase
    loeschNEW "NEIN6", gdBase

    For i = 0 To 6
        loeschNEW "ATT" & i, gdBase
        loeschNEW "TOPNEIN" & i, gdBase
        loeschNEW "TOPNEINT" & i, gdBase
    Next i

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EntwicklungNeinVK"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leeren()
    On Error GoTo LOKAL_ERROR
    
    cboBed.Text = "alle Bediener"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leeren"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub drucken(sAnzeige As String)
    On Error GoTo LOKAL_ERROR
    
    BringFarbeInsSpiel "DRUCK133", gdBase
    
    If Datendrin("DRUCK133", gdBase) Then
        anzeige "normal", "Druckvorschau wird vorbereitet...", lblanzeige
        If sAnzeige = "Packliste" Then
            reportbildschirm "", "aZEN133b"
        ElseIf sAnzeige = "Bestellvorschlag" Then
            reportbildschirm "", "aZEN133a"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub loescheausKUNDBEST(lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr          As String
    Dim cKundnr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
   
    cArtNr = MSHFLEX1.TextMatrix(lrow, SpaltennummerArtnr)
    cKundnr = MSHFLEX1.TextMatrix(lrow, SpaltennummerKUNDNR)
    cBestelltam = MSHFLEX1.TextMatrix(lrow, SpaltennummerBESTELLTAM)
    cBestelltum = MSHFLEX1.TextMatrix(lrow, SpaltennummerBESTELLTUM)
    
    If cArtNr <> "" Then
        If IsNumeric(cArtNr) Then
        
            cSQL = "Delete from KUNDBEST where KUNDNR = " & cKundnr
            cSQL = cSQL & " and ARTNR = " & cArtNr
            cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
            cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
            
            gdBase.Execute cSQL, dbFailOnError
                
            anzeige "normal", "erfolgreich gelöscht", lblanzeige
        
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheausKUNDBEST"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateKUNDBEST(lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr          As String
    Dim cKundnr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    Dim cSTATUS         As String
    
    cArtNr = MSHFLEX1.TextMatrix(lrow, SpaltennummerArtnr)
    cKundnr = MSHFLEX1.TextMatrix(lrow, SpaltennummerKUNDNR)
    cBestelltam = MSHFLEX1.TextMatrix(lrow, SpaltennummerBESTELLTAM)
    cBestelltum = MSHFLEX1.TextMatrix(lrow, SpaltennummerBESTELLTUM)
    
    If cArtNr <> "" Then
        If IsNumeric(cArtNr) Then
            Select Case Combo4.Text
                Case "noch nicht bestellt"
                    cSTATUS = "INBESTELLUNG"
                Case "ist bestellt"
                    cSTATUS = "BESTELLT"
                Case "geliefert"
                    cSTATUS = "GELIEFERT"
                Case "nicht geliefert"
                    cSTATUS = "NICHTGELIEFERT"
                Case Else
                    MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Zentrale Hinweis:"
                    Combo4.SetFocus
                    Exit Sub
                
            End Select
            
            cSQL = "Update KUNDBEST set STATUSARTIKEL = '" & cSTATUS & "'"
            cSQL = cSQL & " Where KUNDNR = " & cKundnr
            cSQL = cSQL & " and ARTNR = " & cArtNr
            cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
            cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
            gdBase.Execute cSQL, dbFailOnError
            
            MSHFLEX1.TextMatrix(lrow, SpaltennummerArtikelStatus) = Combo4.Text
        
        
            
            anzeige "normal", "Veränderung erfolgreich durchgeführt.", lblanzeige
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKUNDBEST"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub fuellecombo1()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    List1.Clear
    
    sSQL = "Select distinct(NEINART) from NEINVK order by NEINART"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!NEINART) Then
                cFeld = rsrs!NEINART
                List1.AddItem cFeld
            End If
            cFeld = ""
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo1"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub fuelleBediener(cboBed As ComboBox)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sTemp As String
    Dim cSatz As String
    Dim cFeld As String
    Dim counter As Long
    Dim lAnzahl As Long
    
    sSQL = "Select distinct(bedname.bednu),bedname.bedname  from bedname inner join NEINVK on"
    sSQL = sSQL & " bedname.bednu = NEINVK.Bednu  order by bedname.bedname"
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboBed.Clear
    cboBed.AddItem "alle Bediener"
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!BEDNU) Then
            
                cFeld = rs!BEDNU
                cSatz = cSatz & Space(4 - Len(cFeld)) & cFeld
                
                If Not IsNull(rs!bedname) Then
                    cFeld = rs!bedname
                    cSatz = cSatz & Space(2) & cFeld
                    cboBed.AddItem cSatz
                    cSatz = ""
                    cFeld = ""
                End If
            End If
        rs.MoveNext
        Loop
    End If
    rs.Close
    cboBed.Text = "alle Bediener"
    
    

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelleBediener"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Sub
Private Sub filcboBediener(cboBed As ComboBox)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sTemp As String
    Dim cSatz As String
    Dim cFeld As String
    Dim counter As Long
    Dim lAnzahl As Long
    
    sSQL = "Select bednu,bedname from bedname order by bedname"
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboBed.Clear
    cboBed.AddItem "alle Bediener"
    
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!BEDNU) Then
            
                cFeld = rs!BEDNU
                cSatz = cSatz & Space(4 - Len(cFeld)) & cFeld
                
                If Not IsNull(rs!bedname) Then
                    cFeld = rs!bedname
                    cSatz = cSatz & Space(2) & cFeld
                    cboBed.AddItem cSatz
                    cSatz = ""
                    cFeld = ""
                End If
            End If
        rs.MoveNext
        Loop
    End If
    rs.Close
    cboBed.Text = "alle Bediener"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "filcboBediener"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    Select Case KeyCode
        
        Case Is = vbKeyF2

            Screen.MousePointer = 11
            lrow = MSHFLEX1.Row
            lcol = MSHFLEX1.Col
            
            gsARTNR = MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerArtnr)
            anzeige "normal", "Die Artikeldaten(" & gsARTNR & ") werden angezeigt...", lblanzeige
            If gsARTNR <> "" Then
                
                frmWKL10.Show 1
                anzeige "normal", "", lblanzeige
                Me.Refresh
                Screen.MousePointer = 11
    
                MSHFLEX1.TopRow = lrow
                MSHFLEX1.Col = lcol
                MSHFLEX1.Row = lrow
                MSHFLEX1.SetFocus
                
                Screen.MousePointer = 0
            End If
            gsARTNR = ""
'        Case Is = vbKeyF3
'
'            lrow = MSHFLEX1.Row
'            lcol = MSHFLEX1.Col
'
'            gcArtNrFiliale = MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerArtnr)
'            anzeige "normal", "Die Artikeldaten(" & gcArtNrFiliale & ") werden angezeigt...", lblAnzeige
'            If gcArtNrFiliale <> "" Then
'                If gbBestinZ Then
'                    frmZENcg.Show 1
'                Else
'                    frmZENcf.Show 1
'                End If
'                anzeige "normal", "", lblAnzeige
'                Me.Refresh
'                Screen.MousePointer = 11
'
'                MSHFLEX1.TopRow = lrow
'                MSHFLEX1.Col = lcol
'                MSHFLEX1.Row = lrow
'                MSHFLEX1.SetFocus
'
'                Screen.MousePointer = 0
'            End If
    End Select



Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub WKL94Positionieren()
    On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Height = 6975
        .Left = 0
        .Top = 840
        .Width = 9495
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL94Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row > 1 Then
        
    Else
        If MSHFLEX1.Col = SpaltennummerBESTELLTAM Then
        
            If byteSortReihen = 1 Then
                byteSortReihen = 2
                ZeigDaten " order by adate desc"
            
            ElseIf byteSortReihen = 2 Then
                byteSortReihen = 1
                ZeigDaten " order by adate asc"
            End If
            
        Else
            sortierenHGrid MSHFLEX1
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub


