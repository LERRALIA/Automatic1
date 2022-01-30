VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL02 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Berechnung von Aufschlag / Netto-Spanne "
   ClientHeight    =   4440
   ClientLeft      =   2385
   ClientTop       =   2370
   ClientWidth     =   6435
   ControlBox      =   0   'False
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
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4440
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   3015
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
      Caption         =   "Preise berechnen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtAbschlag 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtAufschlag 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   21
      Top             =   2040
      Width           =   1335
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   11
      Left            =   5760
      TabIndex        =   20
      Top             =   2280
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "C"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   10
      Left            =   5040
      TabIndex        =   19
      Top             =   2280
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   ","
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   9
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "0"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   8
      Left            =   5760
      TabIndex        =   17
      Top             =   1560
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "3"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   7
      Left            =   5040
      TabIndex        =   16
      Top             =   1560
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   6
      Left            =   4320
      TabIndex        =   15
      Top             =   1560
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "1"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   5
      Left            =   5760
      TabIndex        =   14
      Top             =   840
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "6"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   4
      Left            =   5040
      TabIndex        =   13
      Top             =   840
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "5"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   840
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "4"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   2
      Left            =   5760
      TabIndex        =   11
      Top             =   120
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "9"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   1
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "8"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   615
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   615
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
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
      Caption         =   "7"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3015
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
      Caption         =   "Preise ¸bernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   3015
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Nettospanne %"
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
      Left            =   360
      TabIndex        =   24
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Aufschlag %"
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
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4200
      X2              =   4200
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Kassenverkaufspreis"
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
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Verkaufspreis"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Einkaufspreis"
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
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmWKL02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iby As Byte
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
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim lFeld As Long
    Dim ctmp As String
    
    Select Case Index
        Case 0 To 9
            Select Case iby
                Case Is = 1
                    txtAufschlag.Text = txtAufschlag.Text & Command0(Index).Caption
                Case Is = 0
                    txtAbschlag.Text = txtAbschlag.Text & Command0(Index).Caption
                Case Is = 2
                    Text1(2).Text = Text1(2).Text & Command0(Index).Caption
            End Select
        Case Is = 10    'Komma
            Select Case iby
                Case Is = 1
                    txtAufschlag.Text = txtAufschlag.Text & Command0(Index).Caption
                Case Is = 0
                    txtAbschlag.Text = txtAbschlag.Text & Command0(Index).Caption
                Case Is = 2
                    Text1(2).Text = Text1(2).Text & Command0(Index).Caption
            End Select
        
        Case Is = 11    'Clear
            Select Case iby
                Case Is = 1
                    txtAufschlag.Text = ""
                    txtAufschlag.SetFocus
                Case Is = 0
                    txtAbschlag.Text = ""
                    txtAbschlag.SetFocus
                Case Is = 2
                    Text1(2).Text = ""
                    Text1(2).SetFocus
            End Select
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub aufschlagberechnen()
    On Error GoTo LOKAL_ERROR

    Dim ctmp As String
    Dim ctmp1 As String
    Dim dWert As Double
    Dim dWert1 As Double


    txtAbschlag.Text = ""
    Text1(2).Text = ""
    
    ctmp = Text1(0).Text
    dWert = CDbl(ctmp)
    ctmp1 = txtAufschlag.Text
    dWert1 = CDbl(ctmp1)
    dWert = (dWert * (1 + (dWert1 / 100)))
    dWert = Format$(dWert, "#####0.00")
    
    '//wenn hinter dem Komma "0" vorhanden
    ctmp = CStr(dWert)
    If InStr(1, ctmp, ",") > 0 Then
        If Mid$(ctmp, InStr(1, ctmp, ",") + 2, 1) <> "" Then
            dWert = dWert
            ctmp = CStr(dWert)
        Else
            ctmp = CStr(dWert)
            ctmp = ctmp & "0"
        End If
    Else '// kein Komma vorhanden
        dWert = dWert
        ctmp = CStr(dWert)
    End If
    Text1(2).Text = Format$(ctmp, "#####0.00")
    
    dWert = CDbl(ctmp)
    
    gdRechner(1) = Text1(0).Text
    gdRechner(2) = Text1(1).Text
    gdRechner(3) = Text1(2).Text
    

    '//Netto-Spanne f¸llen
    If gcMwSt = "V" Then
        dWert1 = gdRechner(3) / (100 + gdMWStV) * 100
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1

    ElseIf gcMwSt = "E" Then
        dWert1 = gdRechner(3) / (100 + gdMWStE) * 100
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1
    Else
        dWert1 = gdRechner(3)
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1
    End If
    
    txtAbschlag.Text = Format$(dWert, "#####0.00")
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "aufschlagberechnen"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
   
End Sub
Private Sub NSberechnen()
    On Error GoTo LOKAL_ERROR

    Dim ctmp As String
    Dim ctmp1 As String
    Dim dWert As Double
    Dim dWert1 As Double

    txtAufschlag.Text = ""
    Text1(2).Text = ""

    ctmp = Text1(0).Text
    dWert = CDbl(ctmp)
    ctmp1 = txtAbschlag.Text
    dWert1 = CDbl(ctmp1)
    
    gdRechner(1) = Text1(0).Text
    gdRechner(2) = Text1(1).Text

    If gcMwSt = "V" Then
     
        dWert = -((gdRechner(1) * 100) / (ctmp1 - 100))
        dWert1 = dWert * (100 + gdMWStV) / 100
        Text1(2).Text = Format$(dWert1, "#####0.00")
    
    ElseIf gcMwSt = "E" Then
        dWert = -((gdRechner(1) * 100) / (ctmp1 - 100))
        dWert1 = dWert * (100 + gdMWStE) / 100
        Text1(2).Text = Format$(dWert1, "#####0.00")
    Else
        dWert = -((gdRechner(1) * 100) / (ctmp1 - 100))
        dWert1 = dWert
        Text1(2).Text = Format$(dWert1, "#####0.00")
    End If
    
    gdRechner(3) = Text1(2).Text
        
    Dim dre As Double
    
    dre = IIf(gdRechner(1) = 0, 1, gdRechner(1))
    dWert = (((gdRechner(3) - gdRechner(1)) / dre * 100))
    txtAufschlag.Text = Format$(dWert, "#####0.00")
    
    aufschlagberechnen

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NSberechnen"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Berechnen()
    On Error GoTo LOKAL_ERROR

    Dim ctmp As String
    Dim ctmp1 As String
    Dim dWert As Double
    Dim dWert1 As Double
    Dim cZiel As String
    Dim cZeichen As String
    Dim cValid As String
    Dim lcount As Long
    
    
    cValid = "1234567890,"
            
    'KVKPR
    ctmp = Text1(2).Text
    cZiel = ""
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr(cValid, cZeichen) > 0 Then
            cZiel = cZiel & cZeichen
        End If
    Next lcount
    ctmp = cZiel
    ctmp = fnMoveComma2Point(ctmp)
    gdRechner(3) = Val(ctmp)
    
    
    If gdRechner(1) = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    dWert = (((gdRechner(3) - gdRechner(1)) / gdRechner(1)) * 100)
    txtAufschlag.Text = Format$(dWert, "#####0.00")
    
    If gcMwSt = "V" Then
        dWert1 = gdRechner(3) / (100 + gdMWStV) * 100
        
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1
        txtAbschlag.Text = Format$(dWert, "#####0.00")

    ElseIf gcMwSt = "E" Then
        dWert1 = gdRechner(3) / (100 + gdMWStE) * 100
        
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1
        txtAbschlag.Text = Format$(dWert, "#####0.00")

    Else
        dWert1 = gdRechner(3)
        
        If dWert1 = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        dWert = (dWert1 - gdRechner(1)) * 100 / dWert1
        txtAbschlag.Text = Format$(dWert, "#####0.00")

    End If

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "berechnen"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim ctmp As String
    Dim ctmp1 As String
    Dim dWert As Double
    Dim dWert1 As Double
    
    
    Select Case Index
        
        Case Is = 5
            
            Select Case iby
                Case Is = 1
                    If txtAufschlag.Text <> "" Then
                        aufschlagberechnen
                    Else
                        MsgBox "Bitte Aufschlag-Spanne in % eingeben ! ", vbOKOnly, "Winkiss Hinweis:"
                        txtAufschlag.SetFocus
                    End If
                Case Is = 0
                    If txtAbschlag.Text <> "" Then
                        NSberechnen
                    Else
                        MsgBox "Bitte Netto-Spanne in % eingeben ! ", vbOKOnly, "Winkiss Hinweis:"
                        txtAbschlag.SetFocus
                    End If
                Case Is = 2
                    If Text1(2).Text <> "" Then
                        Berechnen
                    Else
                        MsgBox "Kassenverkaufspreis eingeben ! ", vbOKOnly, "Winkiss Hinweis:"
                        Text1(2).SetFocus
                    End If
            End Select
        Case Is = 2     'Leer
            Text1(0).Text = CStr(gdRechner(1))
            Text1(1).Text = CStr(gdRechner(2))
            Text1(2).Text = ""
            txtAufschlag.Text = ""
            txtAbschlag.Text = ""
            txtAufschlag.SetFocus

        Case Is = 3     'Schlieﬂen
            Unload frmWKL02

        Case Is = 4     '‹bernehmen
            ctmp = Text1(0).Text
            ctmp = fnMoveComma2Point(ctmp)
            gdRechner(1) = Val(ctmp)

            ctmp = Text1(1).Text
            ctmp = fnMoveComma2Point(ctmp)
            gdRechner(2) = Val(ctmp)

            ctmp = Text1(2).Text
            ctmp = fnMoveComma2Point(ctmp)
            gdRechner(3) = Val(ctmp)
            
            ctmp = txtAbschlag.Text
            ctmp = fnMoveComma2Point(ctmp)
            gdRechner(4) = Val(ctmp)

            gdRechner(0) = 1
            Unload frmWKL02

    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim dWert       As Double
    Dim ctmp        As String
    Dim dWert1      As Double
    
    Screen.MousePointer = 11
    
    
    Text1(0).Text = Format$(gdRechner(1), "#####0.00")
    Text1(1).Text = Format$(gdRechner(2), "#####0.00")
    Text1(2).Text = Format$(gdRechner(3), "#####0.00")
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    
    If gdRechner(1) = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Berechnen
    
    
    
Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    iby = 2
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    iby = 2
    cValid = "1234567890," & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If cZeichen = "," Then
        If InStr(Text1(Index).Text, ",") <> 0 Then
            KeyAscii = 0
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub txtAbschlag_GotFocus()
    On Error GoTo LOKAL_ERROR

    txtAbschlag.BackColor = glSelBack1
    txtAbschlag.SelStart = 0
    txtAbschlag.SelLength = Len(txtAbschlag.Text)
    
    iby = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAbschlag_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub txtAbschlag_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String

    cValid = "1234567890," & Chr$(8)

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If

    If cZeichen = "," Then
        If InStr(txtAbschlag.Text, ",") <> 0 Then
            KeyAscii = 0
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAbschlag_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAbschlag_LostFocus()
    On Error GoTo LOKAL_ERROR

    txtAbschlag.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAbschlag_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub txtAufschlag_GotFocus()
    On Error GoTo LOKAL_ERROR

    txtAufschlag.BackColor = glSelBack1
    txtAufschlag.SelStart = 0
    txtAufschlag.SelLength = Len(txtAufschlag.Text)
    
    iby = 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAufschlag_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub txtAufschlag_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String

    cValid = "1234567890," & Chr$(8)

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAufschlag_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAufschlag_lostfocus()
    On Error GoTo LOKAL_ERROR

    txtAufschlag.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAufschlag_lostfocus"
    Fehler.gsFehlertext = "Im Programmteil Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub


