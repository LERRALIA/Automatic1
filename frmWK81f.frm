VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK81f 
   BackColor       =   &H00C0C000&
   Caption         =   "Termine - SMStext"
   ClientHeight    =   8910
   ClientLeft      =   1935
   ClientTop       =   2475
   ClientWidth     =   11910
   Icon            =   "frmWK81f.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   8520
      MaxLength       =   15
      ScrollBars      =   2  'Vertikal
      TabIndex        =   35
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   5400
      TabIndex        =   34
      ToolTipText     =   "Zeilenumbruch"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   5400
      TabIndex        =   33
      ToolTipText     =   "Zeilenumbruch"
      Top             =   6960
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   5400
      TabIndex        =   32
      ToolTipText     =   "Zeilenumbruch"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   5400
      TabIndex        =   31
      ToolTipText     =   "Zeilenumbruch"
      Top             =   6360
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   5400
      TabIndex        =   30
      ToolTipText     =   "Zeilenumbruch"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   5400
      TabIndex        =   29
      ToolTipText     =   "Zeilenumbruch"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   5400
      TabIndex        =   28
      ToolTipText     =   "Zeilenumbruch"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   5400
      TabIndex        =   27
      ToolTipText     =   "Zeilenumbruch"
      Top             =   4680
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   5400
      TabIndex        =   26
      ToolTipText     =   "Zeilenumbruch"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   5400
      TabIndex        =   25
      ToolTipText     =   "Zeilenumbruch"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   5400
      TabIndex        =   24
      ToolTipText     =   "Zeilenumbruch"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   5400
      TabIndex        =   23
      ToolTipText     =   "Zeilenumbruch"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5400
      TabIndex        =   22
      ToolTipText     =   "Zeilenumbruch"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5400
      TabIndex        =   21
      ToolTipText     =   "Zeilenumbruch"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5400
      TabIndex        =   20
      ToolTipText     =   "Zeilenumbruch"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "Zeilenumbruch"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5400
      TabIndex        =   18
      ToolTipText     =   "Zeilenumbruch"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   17
      ToolTipText     =   "Zeilenumbruch"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   6
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   16
      Top             =   6120
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   5
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   15
      Top             =   5280
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   7
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   14
      Top             =   6960
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   4
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   13
      Top             =   4440
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   3
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   11
      Top             =   3600
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   2
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   1
      Left            =   240
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   4
      Top             =   1320
      Width           =   5055
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   3240
      TabIndex        =   3
      Top             =   7680
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
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   8520
      TabIndex        =   2
      Top             =   6360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
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
      Caption         =   "Versende Test SMS"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   1
      Top             =   7080
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9960
      TabIndex        =   0
      Top             =   7680
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Die verwendete Mailadresse bei esendex.de muss die hinterlegte Firmenemailadresse in den Unternehmenseinstellungen sein."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   6240
      MouseIcon       =   "frmWK81f.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   39
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Konto bei esendex.de"
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
      Index           =   8
      Left            =   6240
      MouseIcon       =   "frmWK81f.frx":074C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   38
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Voraussetzungen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   37
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Handynummer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   8520
      TabIndex        =   36
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Anrede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Uhrzeit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Datum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "noch 32 Zeichen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Anrede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Beispieltext:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmWK81f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check29_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    zusammenbauen
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check29_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Select Case Index
        Case 0
            speichern_SMS_Text
        Case 1
            Unload frmWK81f
        Case 2
            'Test SMS
            Dim sTel As String
            Dim cAnEmailadresse As String
            Dim cBetreff As String
            Dim cMessagetext As String
            Dim cAbsenderEmail As String
            Dim sAttachment As String
            
            sAttachment = ""
            
            sTel = Trim(Text1(0).Text)
            
            sTel = SwapStr(sTel, "  ", "")
            sTel = SwapStr(sTel, " ", "")
            sTel = SwapStr(sTel, "/", "")
            sTel = SwapStr(sTel, "\", "")
            sTel = SwapStr(sTel, "-", "")
            
            cAbsenderEmail = ermFirmenMail
            If cAbsenderEmail = "" Then
                MsgBox "Bitte auch in den Unternehmensdaten eine Emaildadresse als Absendermailadresse hinterlegen (Service/Einstellungen/Unternehmens-Daten)", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If

            If sTel <> "" Then
                If IsNumeric(sTel) Then
                    cAnEmailadresse = sTel & "@echoemail.net"
                    'schicke Mail an die hinterlegte Adresse
                    
                    cBetreff = ""
                    cMessagetext = Label1(0).Caption
                    

                    
                    schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, "", cAnEmailadresse _
                    , cAbsenderEmail, gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, cBetreff, cMessagetext, sAttachment
            
            
                End If
            End If
        Case 5
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(5).Text = ""
            Text1(6).Text = ""
            Text1(7).Text = ""
            
            For i = 0 To 17
                Check29(i).Value = vbUnchecked
            Next i
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lade_Standard_SMS()
    On Error GoTo LOKAL_ERROR
    
    Dim cDaten As String
    Dim lcount As Long
    
    Label1(4).Caption = Format(DateValue(Now), "DD.MM.")
    Label1(5).Caption = Format(TimeValue(Now), "HH:MM:SS")
    
    If NewTableSuchenDBKombi("SMSTEXT", gdBase) = False Then
    
        Label1(3).Caption = "Liebe Frau Maier,": Check29(0).Value = vbChecked: Check29(1).Value = vbChecked
    
        Text1(1).Text = "wir freuen uns, Sie zu Ihrer Behandlung am "
        Text1(2).Text = " um "
        Text1(3).Text = " Uhr begrüßen zu dürfen.": Check29(8).Value = vbChecked: Check29(9).Value = vbChecked
        Text1(4).Text = "Dies ist eine computergenerierte SMS, die nicht beantwortet werden kann.": Check29(10).Value = vbChecked
        Text1(5).Text = "Bei Terminänderungen bitte 0511 9559 112, Kiss Kosmetik, anrufen.": Check29(12).Value = vbChecked: Check29(13).Value = vbChecked
        Text1(6).Text = "Liebe Grüße": Check29(14).Value = vbChecked
        Text1(7).Text = "Kiss Kosmetik": Check29(16).Value = vbChecked

        zusammenbauen

    Else
        Command1_Click 5 'leeren
        auslesen
        zusammenbauen
    End If
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lade_Standard_SMS"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub auslesen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    Label1(3).Caption = "Liebe Frau Maier,"
    
    sSQL = "Select * from SMSTEXT"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Zeile1) Then
            Text1(1).Text = rsrs!Zeile1
        End If
        
        If Not IsNull(rsrs!Zeile2) Then
            Text1(2).Text = rsrs!Zeile2
        End If
        
        If Not IsNull(rsrs!Zeile3) Then
            Text1(3).Text = rsrs!Zeile3
        End If
        
        If Not IsNull(rsrs!Zeile4) Then
            Text1(4).Text = rsrs!Zeile4
        End If
        
        If Not IsNull(rsrs!Zeile5) Then
            Text1(5).Text = rsrs!Zeile5
        End If
        
        If Not IsNull(rsrs!Zeile6) Then
            Text1(6).Text = rsrs!Zeile6
        End If
        
        If Not IsNull(rsrs!Zeile7) Then
            Text1(7).Text = rsrs!Zeile7
        End If
        
        If Not IsNull(rsrs!bo0) Then
            If rsrs!bo0 = -1 Then
                Check29(0).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo1) Then
            If rsrs!bo1 = -1 Then
                Check29(1).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo2) Then
            If rsrs!bo2 = -1 Then
                Check29(2).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo3) Then
            If rsrs!bo3 = -1 Then
                Check29(3).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo4) Then
            If rsrs!bo4 = -1 Then
                Check29(4).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo5) Then
            If rsrs!bo5 = -1 Then
                Check29(5).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo6) Then
            If rsrs!bo6 = -1 Then
                Check29(6).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo7) Then
            If rsrs!bo7 = -1 Then
                Check29(7).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo8) Then
            If rsrs!bo8 = -1 Then
                Check29(8).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo9) Then
            If rsrs!bo9 = -1 Then
                Check29(9).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo10) Then
            If rsrs!bo10 = -1 Then
                Check29(10).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo11) Then
            If rsrs!bo11 = -1 Then
                Check29(11).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo12) Then
            If rsrs!bo12 = -1 Then
                Check29(12).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo13) Then
            If rsrs!bo13 = -1 Then
                Check29(13).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo14) Then
            If rsrs!bo14 = -1 Then
                Check29(14).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo15) Then
            If rsrs!bo15 = -1 Then
                Check29(15).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo16) Then
            If rsrs!bo16 = -1 Then
                Check29(16).Value = vbChecked
            End If
        End If
        
        If Not IsNull(rsrs!bo17) Then
            If rsrs!bo17 = -1 Then
                Check29(17).Value = vbChecked
            End If
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "auslesen"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speichern_SMS_Text()
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "SMSTEXT", gdBase
    CreateTableT2 "SMSTEXT", gdBase
    
    Dim lcount As Long
    Dim cSQL As String
    
    cSQL = "Insert into SMSTEXT"
    cSQL = cSQL & " ( "
    cSQL = cSQL & " Anrede "
    cSQL = cSQL & ", Zeile1 "
    cSQL = cSQL & ", Zeile2 "
    cSQL = cSQL & ", Zeile3 "
    cSQL = cSQL & ", Zeile4 "
    cSQL = cSQL & ", Zeile5 "
    cSQL = cSQL & ", Zeile6 "
    cSQL = cSQL & ", Zeile7 "
    cSQL = cSQL & ") values ("
    
    cSQL = cSQL & "''"
    cSQL = cSQL & ",'" & Text1(1).Text & "'"
    cSQL = cSQL & ",'" & Text1(2).Text & "'"
    cSQL = cSQL & ",'" & Text1(3).Text & "'"
    cSQL = cSQL & ",'" & Text1(4).Text & "'"
    cSQL = cSQL & ",'" & Text1(5).Text & "'"
    cSQL = cSQL & ",'" & Text1(6).Text & "'"
    cSQL = cSQL & ",'" & Text1(7).Text & "'"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError
    
    For lcount = 0 To 17
    
        cSQL = "Update SMSTEXT set BO" & lcount & " = '" & Check29(lcount).Value & "' "
        gdBase.Execute cSQL, dbFailOnError
        
    Next lcount
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern_SMS_Text"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    lade_Standard_SMS
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(8).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        
        Case Is = 8
            URLGoTo Me.hwnd, "https://www.esendex.de/"
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    If Index = 8 Then
        Label1(8).ForeColor = glLink
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    zusammenbauen

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub zusammenbauen()
    On Error GoTo LOKAL_ERROR
    
    
    Dim cZeichen(18) As String
    
    Dim i As Integer
    For i = 0 To 17
        If Check29(i).Value = vbChecked Then
            cZeichen(i) = vbCrLf
        Else
            cZeichen(i) = ""
        End If
    Next i
    
    Label1(0).Caption = ""

    Label1(0).Caption = Label1(3).Caption & cZeichen(0) & cZeichen(1)
    Label1(0).Caption = Label1(0).Caption & Text1(1).Text & cZeichen(2) & cZeichen(3)
    Label1(0).Caption = Label1(0).Caption & Label1(4).Caption & cZeichen(4)
    Label1(0).Caption = Label1(0).Caption & Text1(2).Text & cZeichen(5) & cZeichen(6)
    Label1(0).Caption = Label1(0).Caption & Label1(5).Caption & cZeichen(7)
    Label1(0).Caption = Label1(0).Caption & Text1(3).Text & cZeichen(8) & cZeichen(9)
    Label1(0).Caption = Label1(0).Caption & Text1(4).Text & cZeichen(10) & cZeichen(11)
    Label1(0).Caption = Label1(0).Caption & Text1(5).Text & cZeichen(12) & cZeichen(13)
    Label1(0).Caption = Label1(0).Caption & Text1(6).Text & cZeichen(14) & cZeichen(15)
    Label1(0).Caption = Label1(0).Caption & Text1(7).Text & cZeichen(16) & cZeichen(17)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zusammenbauen"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command1_Click 3
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Termine SMS-Text ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub



