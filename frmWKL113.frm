VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL113 
   Caption         =   "Artikel löschen"
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   12
      Left            =   7800
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Zeigen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   11
      Left            =   9000
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Height          =   375
      Index           =   10
      Left            =   10200
      TabIndex        =   18
      ToolTipText     =   "rückgängig"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Rückgängig"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   8
      Left            =   10200
      TabIndex        =   16
      ToolTipText     =   "rückgängig"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Rückgängig"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   7
      Left            =   9000
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Height          =   375
      Index           =   6
      Left            =   7800
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Zeigen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Zeigen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   4
      Left            =   9000
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   10
      ToolTipText     =   "rückgängig"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Rückgängig"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   2
      Left            =   11280
      TabIndex        =   9
      ToolTipText     =   "Protokoll"
      Top             =   7320
      Width           =   375
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
      Caption         =   "P"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   8
      ToolTipText     =   "rückgängig"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Rückgängig"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   9600
      TabIndex        =   7
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   9195
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   9255
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   0
      Left            =   9000
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Height          =   375
      Index           =   9
      Left            =   7800
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Zeigen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
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
   Begin VB.Label Label1 
      Caption         =   "Artikel, die 'schwarz' eingefärbt sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die seit 3 Jahren nicht mehr verkauft wurden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die seit 2 Jahren nicht mehr verkauft wurden und geführt = 'N'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die seit 2 Jahren nicht mehr verkauft wurden und Linie = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel löschen"
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
      Top             =   0
      Width           =   9015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 3
            Unload frmWKL113
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "In Artikel Löschen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim cPfad As String
cPfad = gcDBPfad
If Right(cPfad, 1) <> "\" Then
    cPfad = cPfad & "\"
End If

Select Case Index
    Case 9
        TwoYearsNoVerkauftLPZ0
    Case 0
        DelThis
    Case 1
        Rück
    Case 5
        TwoYearsNoVerkauftGeführtN
    Case 4
        DelThis2
    Case 3
        Rück2
    Case 6
        ThreeYearsNoVerkauft
    Case 7
        DelThis3
    Case 8
        Rück3
    Case 2
        Screen.MousePointer = 11
        zeigeHilfe "LPROTOK", "ArtMenDel.txt", cPfad
        Screen.MousePointer = 0
    Case 12
        
        BLACKART
    Case 11
        DelThis4
    Case 10
        Rück4
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "In Artikel Löschen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    LogtoStart Me
    
    If NewTableSuchenDBKombi("ART55w", gdBase) Then
        If Datendrin("ART55w", gdBase) Then
            Command5(0).Visible = True
        Else
            Command5(0).Visible = False
        End If
    Else
        Command5(0).Visible = False
    End If
    
    If NewTableSuchenDBKombi("ART113", gdBase) Then
        If Datendrin("ART113", gdBase) Then
            Command5(1).Visible = True
        Else
            Command5(1).Visible = False
        End If
    Else
        Command5(1).Visible = False
    End If
    
    '***
    
    If NewTableSuchenDBKombi("ART55Y", gdBase) Then
        If Datendrin("ART55Y", gdBase) Then
            Command5(4).Visible = True
        Else
            Command5(4).Visible = False
        End If
    Else
        Command5(4).Visible = False
    End If
    
    If NewTableSuchenDBKombi("ART113b", gdBase) Then
        If Datendrin("ART113b", gdBase) Then
            Command5(3).Visible = True
        Else
            Command5(3).Visible = False
        End If
    Else
        Command5(3).Visible = False
    End If
    
    If NewTableSuchenDBKombi("ART55i", gdBase) Then
        If Datendrin("ART55i", gdBase) Then
            Command5(11).Visible = True
        Else
            Command5(11).Visible = False
        End If
    Else
        Command5(11).Visible = False
    End If
    
    If NewTableSuchenDBKombi("ART113i", gdBase) Then
        If Datendrin("ART113i", gdBase) Then
            Command5(10).Visible = True
        Else
            Command5(10).Visible = False
        End If
    Else
        Command5(10).Visible = False
    End If
    
    '****
    
    If NewTableSuchenDBKombi("ART55X", gdBase) Then
        If Datendrin("ART55X", gdBase) Then
            Command5(7).Visible = True
        Else
            Command5(7).Visible = False
        End If
    Else
        Command5(7).Visible = False
    End If
    
    If NewTableSuchenDBKombi("ART113c", gdBase) Then
        If Datendrin("ART113c", gdBase) Then
            Command5(8).Visible = True
        Else
            Command5(8).Visible = False
        End If
    Else
        Command5(8).Visible = False
    End If
    
   
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "In Artikel Löschen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    Dim bFound As Boolean
    
    bFound = False
    
    If NewTableSuchenDBKombi("ART113", gdBase) Then
        bFound = True
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART113b", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART113c", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART113i", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = True Then
        iRet = MsgBox("Möchten Sie auch die Sicherungsdaten, die für die 'Rückgängig-Funktion' bereitgehalten wurden, löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            loeschNEW "ART113i", gdBase
            loeschNEW "ARTL113i", gdBase
            loeschNEW "ART113c", gdBase
            loeschNEW "ARTL113c", gdBase
            loeschNEW "ART113b", gdBase
            loeschNEW "ARTL113b", gdBase
            loeschNEW "ART113", gdBase
            loeschNEW "ARTL113", gdBase
        End If
    End If
    
    
    
    bFound = False
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART55w", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART55Y", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART55i", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = False Then
        If NewTableSuchenDBKombi("ART55X", gdBase) Then
            bFound = True
        End If
    End If
    
    If bFound = True Then
        iRet = MsgBox("Möchten Sie auch die Sicherungsdaten, die für die 'Löschen-Funktion' bereitgehalten wurden, löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            loeschNEW "ART55X", gdBase
            loeschNEW "ART55i", gdBase
            loeschNEW "ART55Y", gdBase
            loeschNEW "ART55w", gdBase
        End If
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "In Artikel Löschen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub txtStatus_Change()
On Error GoTo LOKAL_ERROR

    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TwoYearsNoVerkauftLPZ0()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55W", gdBase
    CreateTable "ART55W", gdBase
    
    sSQL = "Update Artikel set bestand = 0 where bestand is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    sSQL = " Insert into ART55w select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 730
    sSQL = sSQL & " and LPZ = 0 "
    sSQL = sSQL & " and bestand <= 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    anzeige "normal", "das Kassenjournal wird importiert, bitte warten...", lbl1
    
    loeschNEW "KASSJOUR", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSJOUR"
    
    txtStatus.Text = 30
 
   
    sSQL = "Create index adate on Kassjour(adate) "
    gdApp.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "Create index artnr on Kassjour(artnr) "
    gdApp.Execute sSQL, dbFailOnError

    txtStatus.Text = 50
    
    anzeige "normal", "die letzten Verkäufe werden ermittelt...", lbl1

    Set rsrs = gdBase.OpenRecordset("ART55w")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)

            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                ldifferenz = 0
                rsrs.Edit
                datLVK = ErmlzVKausApp(cART)
                datLZU = ErmlzZugang(cART)

                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                
                Select Case ldifferenz
           
                    Case Is > 730
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 24 Monaten"
                        End If
                    
                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK
                rsrs!lastzu = datLZU
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

    txtStatus.Text = 10

    sSQL = "Delete from art55w where Monat = '' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from art55w where Monat is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update art55w inner join lisrt on art55w.linr = lisrt.linr "
    sSQL = sSQL & " Set art55w.liefbez = lisrt.liefbez "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 0
    picprogress.Visible = False

    If Datendrin("ART55w", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lbl1
        reportbildschirm "gfd", "aWKL113a"
        Command5(0).Visible = True
        anzeige "normal", "", lbl1
    Else
        anzeige "rot", "keine Artikel gefunden", lbl1
        Command5(0).Visible = False
    End If
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TwoYearsNoVerkauftLPZ0"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BLACKART()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55I", gdBase
    CreateTableT2 "ART55I", gdBase

    sSQL = " Insert into ART55I select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , 0 as BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where awm = '92'"
    sSQL = sSQL & " and BESTAND <= 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Update ART55i inner join lisrt on ART55i.linr = lisrt.linr "
    sSQL = sSQL & " Set ART55i.liefbez = lisrt.liefbez "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 0
    picprogress.Visible = False

    If Datendrin("ART55i", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lbl1
        reportbildschirm "gfd", "aWKL113d"
        Command5(11).Visible = True
        anzeige "normal", "", lbl1
    Else
        anzeige "rot", "keine Artikel gefunden", lbl1
        Command5(11).Visible = False
    End If
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BLACKART"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ThreeYearsNoVerkauft()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55X", gdBase
    CreateTable "ART55X", gdBase
    
    sSQL = "Update Artikel set bestand = 0 where bestand is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    sSQL = " Insert into ART55X select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 1095
    sSQL = sSQL & " and bestand <= 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    anzeige "normal", "das Kassenjournal wird importiert, bitte warten...", lbl1
    
    loeschNEW "KASSJOUR", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSJOUR"
    
    txtStatus.Text = 30
 
   
    sSQL = "Create index adate on Kassjour(adate) "
    gdApp.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "Create index artnr on Kassjour(artnr) "
    gdApp.Execute sSQL, dbFailOnError

    txtStatus.Text = 50
    
    anzeige "normal", "die letzten Verkäufe werden ermittelt...", lbl1

    Set rsrs = gdBase.OpenRecordset("ART55X")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)

            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                ldifferenz = 0
                rsrs.Edit
                datLVK = ErmlzVKausApp(cART)
                datLZU = ErmlzZugang(cART)

                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                
                Select Case ldifferenz
           
                    Case Is > 1095
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 36 Monaten"
                        End If
                    
                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK
                rsrs!lastzu = datLZU
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

    txtStatus.Text = 10

    sSQL = "Delete from ART55X where Monat = '' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from ART55X where Monat is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update ART55X inner join lisrt on ART55X.linr = lisrt.linr "
    sSQL = sSQL & " Set ART55X.liefbez = lisrt.liefbez "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 0
    picprogress.Visible = False

    If Datendrin("ART55X", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lbl1
        reportbildschirm "gfd", "aWKL113c"
        Command5(7).Visible = True
        anzeige "normal", "", lbl1
    Else
        anzeige "rot", "keine Artikel gefunden", lbl1
        Command5(7).Visible = False
    End If
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ThreeYearsNoVerkauft"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub TwoYearsNoVerkauftGeführtN()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55Y", gdBase
    CreateTable "ART55Y", gdBase
    
    sSQL = "Update Artikel set bestand = 0 where bestand is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    sSQL = " Insert into ART55Y select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 730
    sSQL = sSQL & " and gefuehrt ='N' "
    sSQL = sSQL & " and bestand <= 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    anzeige "normal", "das Kassenjournal wird importiert, bitte warten...", lbl1
    
    loeschNEW "KASSJOUR", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSJOUR"
    
    txtStatus.Text = 30
 
   
    sSQL = "Create index adate on Kassjour(adate) "
    gdApp.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "Create index artnr on Kassjour(artnr) "
    gdApp.Execute sSQL, dbFailOnError

    txtStatus.Text = 50
    
    anzeige "normal", "die letzten Verkäufe werden ermittelt...", lbl1

    Set rsrs = gdBase.OpenRecordset("ART55Y")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)

            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                ldifferenz = 0
                rsrs.Edit
                datLVK = ErmlzVKausApp(cART)
                datLZU = ErmlzZugang(cART)

                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                
                Select Case ldifferenz
           
                    Case Is > 730
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 24 Monaten"
                        End If
                    
                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK
                rsrs!lastzu = datLZU
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

    txtStatus.Text = 10

    sSQL = "Delete from ART55Y where Monat = '' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from ART55Y where Monat is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update ART55Y inner join lisrt on ART55Y.linr = lisrt.linr "
    sSQL = sSQL & " Set ART55Y.liefbez = lisrt.liefbez "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 0
    picprogress.Visible = False

    If Datendrin("ART55Y", gdBase) Then
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lbl1
        reportbildschirm "gfd", "aWKL113b"
        Command5(4).Visible = True
        anzeige "normal", "", lbl1
    Else
        anzeige "rot", "keine Artikel gefunden", lbl1
        Command5(4).Visible = False
    End If
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TwoYearsNoVerkauftGeführtN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub DelThis()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim lcount As Long
    Dim lWert As Long
    Dim cdatei As String
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DELART\"
    
'    lWert = DateValue(Now)
'    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "D" & Format$(TimeValue(Now), "HH:MM:SS")
'    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    If Datendrin("ART55w", gdBase) Then
        lcount = DatendrinWieviel("ART55w", gdBase)
        iRet = MsgBox("Möchten Sie wirklich die " & lcount & " Artikel löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    schreibeProtokollArtikelMengenLoeschen "Bediener: " & gcBedienerNr & " löscht " & lcount & " Artikel Dateiname:(" & cdatei & ")"
    
    Screen.MousePointer = 11
    
    loeschNEW "ART113", gdBase
    loeschNEW "ARTL113", gdBase
    
    anzeige "normal", "Artikel werden gelöscht, bitte warten...", lbl1
    
    Kill cPfad & cdatei & ".dbf"
    sSQL = "Select * into " & cdatei & " IN '" & cPfad & "' 'dbase IV;' from Artikel "
    sSQL = sSQL & "where artnr in (Select ARTNR from art55w)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ART113 from Artikel where artnr in (Select ARTNR from art55w)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ARTL113 from Artlief where artnr in (Select ARTNR from art55w)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from Artikel where artnr in (Select ARTNR from art55w)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from ARTLIEF where artnr in (Select ARTNR from art55w)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART55w", gdBase
    
    If NewTableSuchenDBKombi("ART113", gdBase) Then
        If Datendrin("ART113", gdBase) Then
            Command5(0).Visible = False
            Command5(1).Visible = True
        Else
            Command5(1).Visible = False
        End If
    Else
        Command5(1).Visible = False
    End If
    
    anzeige "normal", "Fertig! Artikel sind gelöscht", lbl1
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelThis"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Private Sub DelThis2()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim lcount As Long
    Dim lWert As Long
    Dim cdatei As String
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DELART\"
    
'    lWert = DateValue(Now)
'    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "D" & Format$(TimeValue(Now), "HH:MM:SS")
'    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    If Datendrin("ART55Y", gdBase) Then
        lcount = DatendrinWieviel("ART55Y", gdBase)
        iRet = MsgBox("Möchten Sie wirklich die " & lcount & " Artikel löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    schreibeProtokollArtikelMengenLoeschen "Bediener: " & gcBedienerNr & " löscht " & lcount & " Artikel Dateiname:(" & cdatei & ")"
    
    Screen.MousePointer = 11
    
    loeschNEW "ART113b", gdBase
    loeschNEW "ARTL113b", gdBase
    
    anzeige "normal", "Artikel werden gelöscht, bitte warten...", lbl1
    
    Kill cPfad & cdatei & ".dbf"
    sSQL = "Select * into " & cdatei & " IN '" & cPfad & "' 'dbase IV;' from Artikel "
    sSQL = sSQL & "where artnr in (Select ARTNR from ART55Y)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ART113b from Artikel where artnr in (Select ARTNR from ART55Y)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ARTL113b from Artlief where artnr in (Select ARTNR from ART55Y)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from Artikel where artnr in (Select ARTNR from ART55Y)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from ARTLIEF where artnr in (Select ARTNR from ART55Y)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART55Y", gdBase
    
    If NewTableSuchenDBKombi("ART113b", gdBase) Then
        If Datendrin("ART113b", gdBase) Then
            Command5(4).Visible = False
            Command5(3).Visible = True
        Else
            Command5(3).Visible = False
        End If
    Else
        Command5(3).Visible = False
    End If
    
    anzeige "normal", "Fertig! Artikel sind gelöscht", lbl1
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelThis2"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Private Sub DelThis3()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim lcount As Long
    Dim lWert As Long
    Dim cdatei As String
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DELART\"
    
'    lWert = DateValue(Now)
'    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "D" & Format$(TimeValue(Now), "HH:MM:SS")
'    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    If Datendrin("ART55X", gdBase) Then
        lcount = DatendrinWieviel("ART55X", gdBase)
        iRet = MsgBox("Möchten Sie wirklich die " & lcount & " Artikel löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    schreibeProtokollArtikelMengenLoeschen "Bediener: " & gcBedienerNr & " löscht " & lcount & " Artikel Dateiname:(" & cdatei & ")"
    
    Screen.MousePointer = 11
    
    loeschNEW "ART113c", gdBase
    loeschNEW "ARTL113c", gdBase
    
    anzeige "normal", "Artikel werden gelöscht, bitte warten...", lbl1
    
    Kill cPfad & cdatei & ".dbf"
    sSQL = "Select * into " & cdatei & " IN '" & cPfad & "' 'dbase IV;' from Artikel "
    sSQL = sSQL & "where artnr in (Select ARTNR from ART55X)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ART113c from Artikel where artnr in (Select ARTNR from ART55X)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ARTL113c from Artlief where artnr in (Select ARTNR from ART55X)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from Artikel where artnr in (Select ARTNR from ART55X)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from ARTLIEF where artnr in (Select ARTNR from ART55X)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART55X", gdBase
    
    If NewTableSuchenDBKombi("ART113c", gdBase) Then
        If Datendrin("ART113c", gdBase) Then
            Command5(7).Visible = False
            Command5(8).Visible = True
        Else
            Command5(8).Visible = False
        End If
    Else
        Command5(8).Visible = False
    End If
    
    anzeige "normal", "Fertig! Artikel sind gelöscht", lbl1
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelThis3"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Private Sub DelThis4()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim lcount As Long
    Dim lWert As Long
    Dim cdatei As String
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "DELART\"
    
'    lWert = DateValue(Now)
'    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "D" & Format$(TimeValue(Now), "HH:MM:SS")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    If Datendrin("ART55i", gdBase) Then
        lcount = DatendrinWieviel("ART55i", gdBase)
        iRet = MsgBox("Möchten Sie wirklich die " & lcount & " Artikel löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    schreibeProtokollArtikelMengenLoeschen "Bediener: " & gcBedienerNr & " löscht " & lcount & " Artikel Dateiname:(" & cdatei & ")"
    
    Screen.MousePointer = 11
    
    loeschNEW "ART113i", gdBase
    loeschNEW "ARTL113i", gdBase
    
    anzeige "normal", "Artikel werden gelöscht, bitte warten...", lbl1
    
    Kill cPfad & cdatei & ".dbf"
    sSQL = "Select * into " & cdatei & " IN '" & cPfad & "' 'dbase IV;' from Artikel "
    sSQL = sSQL & "where artnr in (Select ARTNR from ART55i)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ART113i from Artikel where artnr in (Select ARTNR from ART55i)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select  * into ARTL113i from Artlief where artnr in (Select ARTNR from ART55i)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from Artikel where artnr in (Select ARTNR from ART55i)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Delete from ARTLIEF where artnr in (Select ARTNR from ART55i)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART55i", gdBase
    
    If NewTableSuchenDBKombi("ART113i", gdBase) Then
        If Datendrin("ART113i", gdBase) Then
            Command5(11).Visible = False
            Command5(10).Visible = True
        Else
            Command5(10).Visible = False
        End If
    Else
        Command5(10).Visible = False
    End If
    
    anzeige "normal", "Fertig! Artikel sind gelöscht", lbl1
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelThis4"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Private Sub Rück()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    
    If Datendrin("ART113", gdBase) Then
        iRet = MsgBox("Möchten Sie wirklich die " & DatendrinWieviel("ART113", gdBase) & " Artikel wieder einfügen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden wieder einfügt, bitte warten...", lbl1
    
    Screen.MousePointer = 11
    sSQL = "Insert into Artikel select * from ART113 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artlief select * from ARTL113 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART113", gdBase
    loeschNEW "ARTL113", gdBase
    
    Command5(1).Visible = False
    
    anzeige "normal", "Fertig! Artikel sind wieder eingefügt", lbl1
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rück"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Rück2()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    
    If Datendrin("ART113b", gdBase) Then
        iRet = MsgBox("Möchten Sie wirklich die " & DatendrinWieviel("ART113b", gdBase) & " Artikel wieder einfügen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden wieder einfügt, bitte warten...", lbl1
    
    Screen.MousePointer = 11
    sSQL = "Insert into Artikel select * from ART113b "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artlief select * from ARTL113b "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART113b", gdBase
    loeschNEW "ARTL113b", gdBase
    
    Command5(3).Visible = False
    
    anzeige "normal", "Fertig! Artikel sind wieder eingefügt", lbl1
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rück2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Rück3()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    
    If Datendrin("ART113c", gdBase) Then
        iRet = MsgBox("Möchten Sie wirklich die " & DatendrinWieviel("ART113c", gdBase) & " Artikel wieder einfügen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden wieder einfügt, bitte warten...", lbl1
    
    Screen.MousePointer = 11
    sSQL = "Insert into Artikel select * from ART113c "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artlief select * from ARTL113c "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART113c", gdBase
    loeschNEW "ARTL113c", gdBase
    
    Command5(8).Visible = False
    
    anzeige "normal", "Fertig! Artikel sind wieder eingefügt", lbl1
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rück3"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Rück4()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    
    If Datendrin("ART113i", gdBase) Then
        iRet = MsgBox("Möchten Sie wirklich die " & DatendrinWieviel("ART113i", gdBase) & " Artikel wieder einfügen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden wieder einfügt, bitte warten...", lbl1
    
    Screen.MousePointer = 11
    sSQL = "Insert into Artikel select * from ART113i "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artlief select * from ARTL113i "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART113i", gdBase
    loeschNEW "ARTL113i", gdBase
    
    Command5(1).Visible = False
    
    anzeige "normal", "Fertig! Artikel sind wieder eingefügt", lbl1
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rück4"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



