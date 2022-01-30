VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL55 
   Caption         =   "Diverse Artikellisten"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL55.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtxTage 
      Height          =   315
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   50
      Text            =   "365"
      Top             =   6720
      Width           =   615
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   18
      Left            =   10320
      TabIndex        =   42
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   17
      Left            =   10320
      TabIndex        =   40
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   16
      Left            =   10320
      TabIndex        =   37
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   15
      Left            =   10320
      TabIndex        =   36
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   14
      Left            =   10320
      TabIndex        =   34
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.OptionButton Option2 
      Caption         =   "alle Artikel"
      Height          =   255
      Left            =   2520
      TabIndex        =   33
      Top             =   6120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "nur Artikel mit Bestand"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Value           =   -1  'True
      Width           =   2295
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   13
      Left            =   10320
      TabIndex        =   30
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   28
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   11
      Left            =   10320
      TabIndex        =   26
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   22
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   20
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmWKL55.frx":0442
      Left            =   5880
      List            =   "frmWKL55.frx":0444
      TabIndex        =   15
      Text            =   "1 Monat"
      Top             =   1920
      Width           =   2535
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   5
      Left            =   10320
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   4
      Left            =   10320
      TabIndex        =   12
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   9195
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   9480
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
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
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   19
      Left            =   10320
      TabIndex        =   44
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   20
      Left            =   4680
      TabIndex        =   46
      Top             =   6480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   21
      Left            =   10320
      TabIndex        =   48
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   22
      Left            =   4680
      TabIndex        =   52
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   23
      Left            =   10320
      TabIndex        =   54
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   56
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin VB.Label Label1 
      Caption         =   "Artikel, die zur Zeit 'preis-geschützt' sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   57
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand > 0  und deren S-EK = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5880
      TabIndex        =   55
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, deren Handelsspanne < 30 % ist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   53
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "x Tage:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3840
      TabIndex        =   51
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand < 0  und deren EK = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5880
      TabIndex        =   49
      Top             =   6960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand > 0 und seit x Tagen nicht mehr verkauft worden sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   120
      TabIndex        =   47
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "erhöhte KassenVK (KVK > LVK)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   5880
      TabIndex        =   45
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand > 0  und deren L-EK = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5880
      TabIndex        =   43
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, MwSt = 'O'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5880
      TabIndex        =   41
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "reduzierte Kassenverkaufspreise (Kassenverkaufspreis < Listeneinkaufspreis)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   5880
      TabIndex        =   39
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "reduzierte KassenVK (KVK < LVK)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5880
      TabIndex        =   38
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "negative Erträge des letzten Monats"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5880
      TabIndex        =   35
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Lagerwerte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5880
      TabIndex        =   31
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand > 0  und geführt = N sind."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Lagerumschlag "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   27
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die nicht bonusfähig sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, nach Farbmerkmalen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die in den letzten 12 Monaten verkauft worden sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand > 0  und RKZ = J (geräumt) sind."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand = 0 sind."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Bestellwerte ermitteln Verkaufszeitraum zurückblickend von"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die seit 2 Jahren nicht mehr verkauft worden sind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die aus der Datenbank gelöscht wurden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die in den letzten 30 Tagen verkauft wurden und sich dem Bestand 0 nähern."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel, die im Bestand sind und deren Kassenverkaufspreis = 0 ist."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4095
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
      Caption         =   "Diverse Artikellisten"
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
      TabIndex        =   2
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label9 
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
Attribute VB_Name = "frmWKL55"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "ART55H", gdBase
    loeschNEW "ART38", gdBase
    loeschNEW "LAGPENN", gdBase
    loeschNEW "ART55", gdBase
    loeschNEW "RKZART", gdBase
    loeschNEW "GEFART", gdBase
    loeschNEW "ART58", gdBase
    loeschNEW "kasslvk", gdBase
    loeschNEW "NEGART", gdBase
    loeschNEW "ART56", gdBase
    loeschNEW "ART56A", gdBase
    loeschNEW "ART56B", gdBase
    loeschNEW "NEGERTRAGK", gdBase
    loeschNEW "NEGERTRAG", gdBase
    loeschNEW "NEGERTRAGPR", gdBase
    loeschNEW "REDART", gdBase
    loeschNEW "EKNULLART", gdBase
    loeschNEW "REDARTEAN", gdBase
    
    
    loeschNEW "NSPUNTER", gdBase
    loeschNEW "KL_ARTLIEF", gdBase
    loeschNEW "KL_LASTVK", gdBase
    
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
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim iTage As Integer
Dim iRet As Integer

Select Case Index

    Case 0
        Unload frmWKL55
    Case 1
        imbestandundKVKNULL
'        imbestandundNoVerkauft
    Case 2
        Renner
    Case 3 'Gelöschte Artikel
        zeigeHilfeDabapfad "LPROTOK", "geloeschteArtikel.txt"
    Case 4
        TwoYearsNoVerkauft
    Case 20
        XTage_NichtVerkauft txtxTage.Text
    Case 5
        Select Case Combo1.Text
        
            Case "6 Monate"
                iTage = 180
            Case "5 Monate"
                iTage = 153
            Case "4 Monate"
                iTage = 121
            Case "3 Monate"
                iTage = 91
            Case "2 Monate"
                iTage = 61
            Case "1 Monat"
                iTage = 30
            Case "2 Wochen"
                iTage = 14
            Case Else
                iTage = 30
        End Select
        Besteller iTage
    Case 6
        NULLBESTANDart
    Case 7
        BESTANDartAndRkzj
        
        If Datendrin("RKZART", gdBase) Then
            reportbildschirm "WKL024", "aWKL552"
        Else
            anzeigeNew "Rot", "Es wurden keine Artikel ermittelt.", Label9
        End If
    Case 8
        Renner12
    Case 9
        If Option1.Value = True Then
            Farbmerkmalsliste "nur mit Bestand"
        Else
            Farbmerkmalsliste "alle"
        End If
        
    Case 10
        ArtikelnichtBonus
    Case 11
        If alleLUGnachLief(txtStatus, picprogress, True) Then
        
            anzeige "normal", "Druckvorschau wird erstellt...", Label9
            reportbildschirm "", "aZEN08u"
            anzeige "normal", "", Label9
        End If
    Case 12
        BESTANDartAndgefuehrtN
        
        If Datendrin("GEFART", gdBase) Then
            reportbildschirm "WKL024", "aWKL553"
        Else
            anzeigeNew "Rot", "Es wurden keine Artikel ermittelt.", Label9
        End If
    Case 13
        LagerwertePennerwerte
    Case 14
        negativErtrag
    Case 15
        redart
    Case 16
        iRet = MsgBox("Nur Artikel mit Bestand anzeigen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            redart_LEK True
        Else
            redart_LEK False
        End If
        
    Case 17
        ArtikelohneMwst
    Case 18
        BESTANDartAndEKNull
        
        If Datendrin("EKNULLART", gdBase) Then
            reportbildschirm "WKL024", "aWKL554"
        Else
            anzeigeNew "Rot", "Es wurden keine Artikel ermittelt.", Label9
        End If
    Case 19
        Highart
     Case 21
        NegBESTANDartAndEKNull
        
        If Datendrin("EKNULLART", gdBase) Then
            reportbildschirm "WKL024", "aWKL555"
        Else
            anzeigeNew "Rot", "Es wurden keine Artikel ermittelt.", Label9
        End If
    Case 22 'NSP < 30
        NSPunter30
    Case 23
        BESTANDartAndSchnittEKNull
        
        If Datendrin("EKNULLART", gdBase) Then
            reportbildschirm "WKL024", "aWKL554a"
        Else
            anzeigeNew "Rot", "Es wurden keine Artikel ermittelt.", Label9
        End If
    Case 24 'Preisschutz
        PSart
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub redart()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "REDARTEAN", gdBase
    CreateTableT2 "REDARTEAN", gdBase
    
    cSQL = "Insert into REDARTEAN Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , EAN "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , 0 as LEK "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , VKPR "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where  vkpr > kvkpr1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "delete from REDARTEAN where  vkpr = kvkpr1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on REDARTEAN(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REDARTEAN inner join LISRT on REDARTEAN.Linr = LISRT.Linr "
    cSQL = cSQL & " set REDARTEAN.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label9

    reportbildschirm "WKL024", "aWKL40ba"
    

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "redart"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Highart()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "REDART", gdBase
    CreateTable "REDART", gdBase
    
    cSQL = "Insert into REDART Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , 0 as LEK "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , VKPR "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where  vkpr < kvkpr1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "delete from REDART where  vkpr = kvkpr1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on REDART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REDART inner join LISRT on REDART.Linr = LISRT.Linr "
    cSQL = cSQL & " set REDART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label9

    reportbildschirm "WKL024", "aWKL40bc"
    
'    anzeige "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Highart"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub PSart()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "PSART", gdBase
    CreateTableT3 "PSART", gdBase
    
    cSQL = "Insert into PSART Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , '' as LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , 0 as LEK "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , VKPR "
    cSQL = cSQL & " , 0 as LINR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where  Preisschu = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
'    cSQL = "Create index linr on REDART(linr) "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update REDART inner join LISRT on REDART.Linr = LISRT.Linr "
'    cSQL = cSQL & " set REDART.LIEFBEZ = LISRT.LIEFBEZ "
'    gdBase.Execute cSQL, dbFailOnError

    
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label9

    reportbildschirm "WKL024", "aWKL40be"
    
'    anzeige "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PSart"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub NSPunter30()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "NSPUNTER", gdBase
    CreateTableT3 "NSPUNTER", gdBase
    
    anzeige "normal", "kleinste EK...", Label9
    
    
    'kleinste lek Tabelle bauen
    loeschNEW "KL_ARTLIEF", gdBase
    
    cSQL = "Select Artnr, min(LEKPR) as MinLEKPR into KL_ARTLIEF"
    cSQL = cSQL & " from ARTLIEF where LEKPR > 0 group by Artnr  "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "letzter Verkauf...", Label9
    
    'last vk Tabelle bauen
    loeschNEW "KL_LASTVK", gdBase
    
    cSQL = "Select Artnr, max(adate) as lastvk into KL_LASTVK"
    cSQL = cSQL & " from Kassjour  group by Artnr  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Artikel zusammenstellen...", Label9
    
    
    cSQL = "Insert into NSPUNTER Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " ,'' as BEZEICH "
    cSQL = cSQL & " , 0 as BESTAND "
    cSQL = cSQL & " , 'V' as MWST "
    cSQL = cSQL & " , '' as LIBESNR "
    cSQL = cSQL & " , MinLEKPR as LEK "
    cSQL = cSQL & " , 0 as VKPR "
    cSQL = cSQL & " , 0 as KVKPR1 "
    cSQL = cSQL & " , 0 as LINR "
    cSQL = cSQL & " ,'' as  LIEFBEZ "
    cSQL = cSQL & " , 0 as NSP "
    cSQL = cSQL & " , null as lastvk "
    cSQL = cSQL & " from KL_ARTLIEF "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Index wird erstellt...", Label9
    
    cSQL = "Create index ARTNR on NSPUNTER(ARTNR) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index LEK on NSPUNTER(LEK) "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Lieferant wird aktualisiert...", Label9
    
    cSQL = "Update NSPUNTER inner join ARTLIEF on NSPUNTER.ARTNR = ARTLIEF.ARTNR "
    cSQL = cSQL & " and NSPUNTER.LEK = ARTLIEF.LEKPR "
    cSQL = cSQL & " set NSPUNTER.LINR = ARTLIEF.LINR "
    cSQL = cSQL & " , NSPUNTER.LIBESNR = ARTLIEF.LIBESNR "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    anzeige "normal", "Artikel wird aktualisiert...", Label9
    
    
    cSQL = "Update NSPUNTER inner join Artikel on NSPUNTER.ARTNR = Artikel.ARTNR "
    cSQL = cSQL & " set NSPUNTER.BEZEICH = Artikel.BEZEICH "
    cSQL = cSQL & " , NSPUNTER.BESTAND = Artikel.BESTAND "
    cSQL = cSQL & " , NSPUNTER.MWST = Artikel.MWST "
    cSQL = cSQL & " , NSPUNTER.KVKPR1 = Artikel.KVKPR1 "
    cSQL = cSQL & " , NSPUNTER.VKPR = Artikel.VKPR "
    gdBase.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Handelsspanne wird errechnet...", Label9
    
    
    cSQL = "Update NSPUNTER set NSP = "
    cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - lek) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
    cSQL = cSQL & "  "
    cSQL = cSQL & "  where MWST = 'V' and lek > 0 and KVKPR1 > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NSPUNTER set NSP = "
    cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - lek) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
    cSQL = cSQL & "  "
    cSQL = cSQL & "  where MWST = 'E' and lek > 0 and KVKPR1 > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NSPUNTER set NSP = "
    cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - lek) *100) / ((KVKPR1 * 100) / (100)) "
    cSQL = cSQL & "  "
    cSQL = cSQL & "  where MWST = 'O' and lek > 0 and KVKPR1 > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Artikel werden bereinigt...", Label9
    
    cSQL = "delete from NSPUNTER where  NSP >= 30 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "letzter Verkauf...", Label9
    
    cSQL = "Update NSPUNTER inner join KL_LASTVK on NSPUNTER.ARTNR = KL_LASTVK.ARTNR "
    cSQL = cSQL & " set NSPUNTER.LASTVK = KL_LASTVK.LASTVK "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Lieferantenbezeichnung...", Label9
    
    
    cSQL = "Create index linr on NSPUNTER(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NSPUNTER inner join LISRT on NSPUNTER.Linr = LISRT.Linr "
    cSQL = cSQL & " set NSPUNTER.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label9
    
    Screen.MousePointer = 0

    reportbildschirm "WKL024", "aWKL40bd"
    
'    anzeige "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NSPunter30"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub redart_LEK(bmitBestand As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "REDART", gdBase
    CreateTable "REDART", gdBase
    
    cSQL = "Insert into REDART Select "
    cSQL = cSQL & " ARTIKEL.ARTNR "
    cSQL = cSQL & " , ARTIKEL.BEZEICH "
    cSQL = cSQL & " , ARTIKEL.LIBESNR "
    cSQL = cSQL & " , ARTIKEL.BESTAND "
    cSQL = cSQL & " , min(Artlief.LEKPR) as LEK "
    cSQL = cSQL & " , ARTIKEL.KVKPR1 "
    cSQL = cSQL & " , ARTIKEL.VKPR "
    cSQL = cSQL & " , ARTIKEL.LINR "
    cSQL = cSQL & " from ARTIKEL inner join ARTLIEF on ARTIKEL.Artnr = Artlief.Artnr "
'    cSQL = cSQL & " where ARTLIEF.LEKPR > ARTIKEL.kvkpr1 "
    cSQL = cSQL & " group by  "
    cSQL = cSQL & " ARTIKEL.ARTNR "
    cSQL = cSQL & " , ARTIKEL.BEZEICH "
    cSQL = cSQL & " , ARTIKEL.LIBESNR "
    cSQL = cSQL & " , ARTIKEL.BESTAND "
    cSQL = cSQL & " , ARTIKEL.KVKPR1 "
    cSQL = cSQL & " , ARTIKEL.VKPR "
    cSQL = cSQL & " , ARTIKEL.LINR "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "TEMP", gdBase

    cSQL = "Select * into TEMP from REDART "
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "REDART", gdBase
    CreateTable "REDART", gdBase

    cSQL = "Insert into REDART Select * from TEMP "
    cSQL = cSQL & " where LEK > kvkpr1 "
    If bmitBestand = True Then
        cSQL = cSQL & " and Bestand > 0 "
    End If
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "TEMP", gdBase
    

    cSQL = "delete from REDART where kvkpr1 = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REDART inner join LISRT on REDART.Linr = LISRT.Linr "
    cSQL = cSQL & " set REDART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create index linr on REDART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label9

    reportbildschirm "", "aWKL40bb"
    
'    anzeige "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "redart_LEK"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Besteller(iTage As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim lBestand As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "tempZu", gdBase
    sSQL = "select max(adate)as maxdate, artnr  into tempZu from zugang where bewegung > 0 "
    sSQL = sSQL & " group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 8
    
    sSQL = "Create index  ARTNR on tempZu(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 10
    
    '*******
    
    loeschNEW "ART58", gdBase
    CreateTable "ART58", gdBase
    
    '1.Schritt alle Artikel auswählen

    sSQL = " Insert into ART58 select a.ARTNR"
    sSQL = sSQL & " , a.Bezeich "
    sSQL = sSQL & " , a.RKZ "
    sSQL = sSQL & " , a.LEKPR "
    sSQL = sSQL & " , a.KVKPR1 "
    sSQL = sSQL & " , a.LINR "
    sSQL = sSQL & " , a.LPZ "
    sSQL = sSQL & " , 0 as BESTAND "
    sSQL = sSQL & " , 0 as VKMENGE "
    sSQL = sSQL & " , 0 as INBEST "
    
    sSQL = sSQL & " , 0.0 as VKLEKWERT "
    sSQL = sSQL & " , 0.0 as BeLEKWERT "
    sSQL = sSQL & " , 0.0 as InBeLEKWERT "
    
    sSQL = sSQL & " , 0.0 as AWERT "
    sSQL = sSQL & " , 0.0 as MBESTWERT "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", a.AUFDAT  "
    sSQL = sSQL & ", a.EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", t.maxdate as LASTZU "
'    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , a.LIBESNR from Artikel a "
    sSQL = sSQL & " inner join tempZu t on a.artnr = t.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 11
    
    sSQL = "Update ART58 inner join ARTLIEF on "
    sSQL = sSQL & " ART58.ARTNR = ARTLIEF.ARTNR and ART58.LINR = ARTLIEF.LINR"
    sSQL = sSQL & " Set ART58.LEKPR = ARTLIEF.LEKPR "
    sSQL = sSQL & " , ART58.RKZ = ARTLIEF.RKZ "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 12
    
    '2.Schritt Lastvk schreiben
    txtStatus.Text = 15
    
    
    loeschNEW "kasslvk", gdBase
    
    sSQL = "select sum(menge)as vkmenge, artnr  ,max(adate)as addat into kasslvk from kassjour where adate > " & CLng(DateValue(Now) - iTage)
    sSQL = sSQL & " and menge > 0 "
    sSQL = sSQL & " group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20
    
    sSQL = " Create index  ARTNR on kasslvk(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = " Create index  addat on kasslvk(addat) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = " Create index  ARTNR on ART58(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    sSQL = "delete from ART58 where artnr not in(select artnr from kasslvk where kasslvk.artnr = art58.artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "update ART58 inner join kasslvk on art58.artnr = kasslvk.artnr set art58.vkmenge = kasslvk.vkmenge "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "update ART58 inner join kasslvk on art58.artnr = kasslvk.artnr set art58.lastvk = kasslvk.addat "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45
    
    sSQL = "Delete from art58 where vkmenge < 1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 50
    
    sSQL = "Delete from art58 where rkz = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 55
    
    sSQL = "update ART58 set VKLEKWERT = LEKPR * vkmenge "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 57
    
    sSQL = "update ART58 set BELEKWERT = LEKPR * BESTAND "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 59
    
    loeschNEW "InbestK", gdBase
    
    txtStatus.Text = 61
    
    sSQL = "select sum(bestvor)as inbestka,artnr  into InbestK from Bestrest "
    sSQL = sSQL & " group by artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 64
    
    sSQL = "update ART58 inner join InbestK on art58.artnr = InbestK.artnr set art58.inbest = InbestK.inbestka "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 66
    
    sSQL = "update ART58 set inbeLEKWERT = LEKPR * inbest "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 68
    
    sSQL = "update ART58 set MBESTWERT = VKLEKWERT -( inbeLEKWERT + beLEKWERT) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70
    
    sSQL = "update ART58 set monat =  " & iTage
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 73

    sSQL = "Update art58 inner join lisrt on art58.linr = lisrt.linr "
    sSQL = sSQL & " Set art58.liefbez = lisrt.liefbez "
    sSQL = sSQL & " , art58.awert = lisrt.awert "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 80
    
    sSQL = "update ART58 set awert = 0 where awert is null "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 100
    
    txtStatus.Text = 0
    picprogress.Visible = False
    loeschNEW "tempZu", gdBase

    Screen.MousePointer = 0
    reportbildschirm "gfd", "aWKL55d"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Besteller"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
   
End Sub
Private Sub TwoYearsNoVerkauft()
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
    
    loeschNEW "Kopf55", gdBase
    CreateTable "KOPF55", gdBase
    
    loeschNEW "ART55", gdBase
    CreateTable "ART55", gdBase

    sSQL = " Insert into ART55 select  ARTNR"
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
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 730
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    loeschNEW "KASSJOUR", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSJOUR"
    
    txtStatus.Text = 30
    
    sSQL = "Create index adate on Kassjour(adate) "
    gdApp.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "Create index artnr on Kassjour(artnr) "
    gdApp.Execute sSQL, dbFailOnError


    txtStatus.Text = 50

    Set rsrs = gdBase.OpenRecordset("ART55")
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

    sSQL = "Delete from art55 where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from art55 where Monat is null "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update art55 inner join lisrt on art55.linr = lisrt.linr "
    sSQL = sSQL & " Set art55.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update art55 inner join Artikel on art55.ARTNR = Artikel.Artnr "
    sSQL = sSQL & " set art55.BESTAND = ARTIKEL.BESTAND "
    gdBase.Execute sSQL, dbFailOnError
    

    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    reportbildschirm "gfd", "aWKL55c"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TwoYearsNoVerkauft"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub XTage_NichtVerkauft(sXtage As String)
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
    
    
    If sXtage = "" Then sXtage = "365"
    
    If IsNumeric(sXtage) = False Then sXtage = "365"
    
    
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "Kopf55", gdBase
    CreateTable "KOPF55", gdBase
    
    
    sSQL = "Insert into Kopf55 (Auswahl)"
    sSQL = sSQL & " values ( "
    sSQL = sSQL & sXtage
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ART55B", gdBase
    CreateTable "ART55B", gdBase

    sSQL = " Insert into ART55B select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , EKPR as SEKPR"
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", null as LASTVK "
    sSQL = sSQL & ", null as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel"
    sSQL = sSQL & " where bestand > 0  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 15
    
    
    loeschNEW "Kass55", gdBase
    
    sSQL = " select Artnr ,Max(adate) as MAXVKDAT into Kass55 from Kassjour "
    sSQL = sSQL & " group by artnr  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20
    
    loeschNEW "Zu55", gdBase
    
    sSQL = " select Artnr ,Max(adate) as MAXZUDAT into Zu55 from Zugang "
    sSQL = sSQL & " group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = "Update ART55B inner join Kass55 on ART55B.artnr = kass55.artnr  set ART55B.LASTVK = Kass55.MAXVKDAT   "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = "Update ART55B inner join Zu55 on ART55B.artnr = Zu55.artnr  set ART55B.LASTZU = Zu55.MAXZUDAT   "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    sSQL = "Delete from ART55B where LASTVK >= " & CLng(DateValue(Now)) - sXtage
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45

    sSQL = "Update ART55B inner join artlief on ART55B.artnr = artlief.artnr "
    sSQL = sSQL & " Set ART55B.lekpr = artlief.lekpr "
    sSQL = sSQL & " , ART55B.linr = artlief.linr "
    sSQL = sSQL & " , ART55B.RKZ = artlief.RKZ "
    sSQL = sSQL & " , ART55B.EXDAT = artlief.EXDAT "
    sSQL = sSQL & " , ART55B.LIBESNR = artlief.LIBESNR "
    gdBase.Execute sSQL, dbFailOnError
    
    
    


    txtStatus.Text = 60

    sSQL = "Update ART55B inner join lisrt on ART55B.linr = lisrt.linr "
    sSQL = sSQL & " Set ART55B.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    reportbildschirm "gfd", "aWKL55k"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "XTage_NichtVerkauft"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub NULLBESTANDart()

    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    loeschNEW "NEGART", gdBase
    CreateTable "NEGART", gdBase
    cSQL = "Insert into NEGART Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " , EAN "
    
    cSQL = cSQL & " from ARTIKEL where BESTAND = 0 "
    cSQL = cSQL & " and not Bestand is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on NEGART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NEGART inner join LISRT on NEGART.Linr = LISRT.Linr "
    cSQL = cSQL & " set NEGART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    reportbildschirm "WKL024", "aWKL55l"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NULLBESTANDart"
    Fehler.gsFehlertext = "Beim Ermitteln der negativen Bestände ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Renner()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim lBestand As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "Kopf55", gdBase
    CreateTable "KOPF55", gdBase
    
    loeschNEW "ART56", gdBase
    CreateTable "ART56", gdBase
    
    '1.Schritt alle Artikel auswählen

    sSQL = " Insert into ART56 select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "
    sSQL = sSQL & " , 0 as VKMENGE "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    '2.Schritt Lastvk schreiben
    txtStatus.Text = 0
    
    
    loeschNEW "kasslvk", gdBase
    
    sSQL = "select sum(menge)as vkmenge,artnr , adate  into kasslvk from kassjour where adate > " & CLng(DateValue(Now) - 30)
    sSQL = sSQL & " group by artnr,adate"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = " Create index  ARTNR on kasslvk(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = " Create index  adate on kasslvk(adate) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    sSQL = " Create index  ARTNR on ART56(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    
    sSQL = "delete from ART56 where artnr not in(select artnr from kasslvk where kasslvk.artnr = art56.artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 50
    

    Set rsrs = gdBase.OpenRecordset("ART56")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        
        Dim dakuta As Double
        Dim lVkMenge As Long
        
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)


            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                
                If Not IsNull(rsrs!BESTAND) Then
                    lBestand = rsrs!BESTAND
                Else
                    lBestand = 0
                End If
                
                rsrs.Edit
                datLVK = ErmlzVKFTemp(cART)
                datLZU = ErmlzZugang(cART)
                lVkMenge = ErmVKMFTemp(cART)


                If lVkMenge = 0 Then
                    ctmp = ""
                Else
                
                    dakuta = (100 * lBestand) / lVkMenge
                    
                    Select Case dakuta
                    
                    Case Is > 30
                        ctmp = "ausreichend vorhanden"
                    Case Is > 0
                        ctmp = "demnächst nachbestellen!"
                    Case Is = 0
                        ctmp = "unbedingt nachbestellen!"
                    Case Else
                        ctmp = ""
                End Select
                End If


                

                rsrs!Monat = ctmp
                rsrs!VKMENGE = lVkMenge
                rsrs!lastvk = datLVK
                rsrs!lastzu = datLZU
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    


    txtStatus.Text = 10

    sSQL = "Delete from art56 where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from art56 where Monat is null "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update art56 inner join lisrt on art56.linr = lisrt.linr "
    sSQL = sSQL & " Set art56.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    'jetzt noch Kopfdaten ermitteln
    
    Dim GesBestand      As Long
    Dim TeiBestand      As Long
    Dim SEKWertg        As Double
    Dim LEKWertg        As Double
    Dim KVKWertg        As Double
    Dim SEKWertT        As Double
    Dim LEKWertT        As Double
    Dim KVKWertT        As Double
    
    Dim sSEKWertg        As String
    Dim sLEKWertg        As String
    Dim sKVKWertg        As String
    Dim sSEKWertT        As String
    Dim sLEKWertT        As String
    Dim sKVKWertT        As String
    
    GesBestand = ermgesbestand()
    TeiBestand = ermTeibestand(txtStatus, picprogress, "ART56")
    
    GesBestand = GesBestand - TeiBestand

    SEKWertg = ermSEKWERT(txtStatus, picprogress)
    LEKWertg = ermLEKWERT(txtStatus, picprogress)
    KVKWertg = ermKVKWERT(txtStatus, picprogress)

    SEKWertT = ermSEKWERTT(txtStatus, picprogress, "ART56")
    LEKWertT = ermLEKWERTt(txtStatus, picprogress, "ART56")
    KVKWertT = ermKVKWERTt(txtStatus, picprogress, "ART56")
    
    SEKWertg = SEKWertg - SEKWertT
    LEKWertg = LEKWertg - LEKWertT
    KVKWertg = KVKWertg - KVKWertT
    
    sSEKWertg = CStr(SEKWertg)
    sSEKWertg = SwapStr(sSEKWertg, ",", ".")
    
    sLEKWertg = CStr(LEKWertg)
    sLEKWertg = SwapStr(sLEKWertg, ",", ".")
    
    sKVKWertg = CStr(KVKWertg)
    sKVKWertg = SwapStr(sKVKWertg, ",", ".")
    
    
    sSEKWertT = CStr(SEKWertT)
    sSEKWertT = SwapStr(sSEKWertT, ",", ".")
    
    sLEKWertT = CStr(LEKWertT)
    sLEKWertT = SwapStr(sLEKWertT, ",", ".")
    
    sKVKWertT = CStr(KVKWertT)
    sKVKWertT = SwapStr(sKVKWertT, ",", ".")

    sSQL = "Insert into Kopf55 (RestBestand,Auswahl,SEKWertg,SEKWertt,LEKWertg,LEKWertt,KVKWertg,KVKWertt )"
    sSQL = sSQL & " values ( "
    sSQL = sSQL & GesBestand
    sSQL = sSQL & "," & TeiBestand
    sSQL = sSQL & "," & sSEKWertg
    sSQL = sSQL & "," & sSEKWertT
    sSQL = sSQL & "," & sLEKWertg
    sSQL = sSQL & "," & sLEKWertT
    sSQL = sSQL & "," & sKVKWertg
    sSQL = sSQL & "," & sKVKWertT
    
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    

    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    reportbildschirm "gfd", "aWKL55b"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Renner"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Renner12()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim lBestand As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART56A", gdBase
    CreateTable "ART56A", gdBase
    
    '1.Schritt alle Artikel auswählen

    sSQL = " Insert into ART56A select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , EAN "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , VKPR "
    sSQL = sSQL & " , MWST "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "
    sSQL = sSQL & " , 0 as VKMENGE "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
'    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 90
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    '2.Schritt Lastvk schreiben
    txtStatus.Text = 0
    
    
    loeschNEW "kasslvk", gdBase
    
    sSQL = "select sum(menge)as vkmenge,artnr , adate  into kasslvk from kassjour where adate > " & CLng(DateValue(Now) - 365)
    sSQL = sSQL & " group by artnr,adate"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = " Create index  ARTNR on kasslvk(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = " Create index  adate on kasslvk(adate) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    sSQL = " Create index  ARTNR on ART56A(ARTNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    
    sSQL = "delete from ART56A where artnr not in(select artnr from kasslvk where kasslvk.artnr = ART56A.artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 50
  

    Set rsrs = gdBase.OpenRecordset("ART56A")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        
        Dim dakuta As Double
        Dim lVkMenge As Long
        
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)


            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                
                If Not IsNull(rsrs!BESTAND) Then
                    lBestand = rsrs!BESTAND
                Else
                    lBestand = 0
                End If
                
                rsrs.Edit
                datLVK = ErmlzVKFTemp(cART)
                lVkMenge = ErmVKMFTemp(cART)


                If lVkMenge = 0 Then
                    ctmp = ""
                Else
                
                End If
                rsrs!VKMENGE = lVkMenge
                rsrs!lastvk = datLVK
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    


    txtStatus.Text = 30

    sSQL = "Update ART56A inner join lisrt on ART56A.linr = lisrt.linr "
    sSQL = sSQL & " Set ART56A.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART56B", gdBase
    CreateTable "ART56B", gdBase

    sSQL = " Insert into ART56B select  LINR"
    sSQL = sSQL & " , LIEFBez "
    sSQL = sSQL & " , sum(BESTAND) as Bestand1 "
    sSQL = sSQL & " , sum(VKMENGE) as VKMENGE1 "
    sSQL = sSQL & " , count(artnr) as anz "
    sSQL = sSQL & " from ART56A "
    sSQL = sSQL & " group by linr,liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    reportbildschirm "gfd", "aWKL55e"
    reportbildschirm "gfd", "aWKL55f"
    

        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Renner12"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function ErmlzVKFTemp(cART As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmlzVKFTemp = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Kasslvk where ARTNR = " & cART & " "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVKFTemp = rsINB!MaxDate
        
        End If
    
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmlzVKFTemp"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ErmVKMFTemp(cART As String) As Long
    On Error GoTo LOKAL_ERROR
    
    ErmVKMFTemp = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select sum(vkmenge) as maxdate from Kasslvk where ARTNR = " & cART & " "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmVKMFTemp = rsINB!MaxDate
        
        End If
    
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmVKMFTemp"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function



Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    
    Screen.MousePointer = 11
    
    PositionierenWKL55
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    
    Combo1.Clear
    Combo1.AddItem "6 Monate"
    Combo1.AddItem "5 Monate"
    Combo1.AddItem "4 Monate"
    Combo1.AddItem "3 Monate"
    Combo1.AddItem "2 Monate"
    Combo1.AddItem "1 Monat"
    Combo1.AddItem "2 Wochen"
    Combo1.Text = "2 Monate"
    
   
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
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
Private Sub PositionierenWKL55()
On Error GoTo LOKAL_ERROR
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL55"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LagerwertePennerwerte()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cART As String
    Dim lVkMenge As Long
    Dim datLZU As Date
    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim siAnzeige As Single
    Dim i As Long
    Dim dUmsatz As Double
    Dim dEKUmsatz As Double
    Dim dEkpr As Double

    Dim lVon As Long
    Dim lBis As Long
    
    Dim sVon As String
    Dim sBis As String

    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "LAGPENN", gdBase
    CreateTableT2 "LAGPENN", gdBase
    
    If Month(DateValue(Now)) = 1 Then
        sVon = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
        sBis = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
    Else
        Select Case Month(DateValue(Now)) - 1
            Case 1, 3, 5, 7, 8, 10, 12
                sBis = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Case 2
                sBis = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Case Else
                sBis = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
        End Select
        
        
        
        
        
        
        
        
        
        sVon = Format(DateValue(sBis) + 1, "DD.MM.YYYY")
        
        
        If sVon = "29.02.2020" Then
            sVon = "28.02.2020"
        End If
        
        
        
        
        
        
        
        sVon = Format(Day(DateValue(sVon)) & "." & Month(DateValue(sVon)) & "." & Year(DateValue(sVon)) - 1, "DD.MM.YYYY")
        
    End If

    lVon = CLng(DateValue(sVon))
    lBis = CLng(DateValue(sBis))
    
    lAnz = lBis - lVon
    txtStatus.Text = 0
    
    For i = lVon To lBis
        siAnzeige = siAnzeige + 1
        txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
    
        dUmsatz = ermgesUmsatzausZumsatz(CStr(i), CStr(i))
        dEkpr = ermgesEKausZumsatz(CStr(i), CStr(i))
        dEKUmsatz = ermgesEKZugang(CStr(i), CStr(i), 0)
        sSQL = " Insert into LAGPENN (Datum,UMSATZWERT,EKUMSATZWERT,UMSATZEKWERT) values (" & i & ",'" & dUmsatz & "','" & dEKUmsatz & "','" & dEkpr & "')"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = " Update LAGPENN set UMSATZWERT = null where UMSATZWERT = 0 "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = " Update LAGPENN set EKUMSATZWERT = null where EKUMSATZWERT = 0 "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = " Update LAGPENN set UMSATZEKWERT = null where UMSATZEKWERT = 0 "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    txtStatus.Text = 10
    
    sSQL = " Update LAGPENN inner join LAGERW on LAGPENN.Datum = Lagerw.Datum "
    sSQL = sSQL & " set LAGPENN.LAGSEK = LAGERW.SEK "
    sSQL = sSQL & " , LAGPENN.LAGBEST = LAGERW.BEST "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 15
    
    sSQL = " Update LAGPENN inner join PENLAGERW on LAGPENN.Datum = PENLAGERW.Datum "
    sSQL = sSQL & " set LAGPENN.PENNSEK = PENLAGERW.SEK "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20
   
    Screen.MousePointer = 0
    picprogress.Visible = False
    
    reportbildschirm "", "aZEN00a1"
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LagerwertePennerwerte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub negativErtrag()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iMonat As Integer
    Dim iJahr As Integer
    Dim sDatum As String
    
    Screen.MousePointer = 11
    
    iMonat = Month(Now)
    iJahr = Year(Now)
    
    If iMonat = 1 Then
        iMonat = 12
        iJahr = iJahr - 1
    Else
        iMonat = iMonat - 1
        iJahr = iJahr
    End If
    
'    iMonat = 8
'    iJahr = 2010
    
    sDatum = MonthName(CLng(iMonat)) & " " & iJahr
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "NEGERTRAG", gdBase
    CreateTableT2 "NEGERTRAG", gdBase
    
    sSQL = "Insert into NEGERTRAG Select artnr, bezeich ,menge as summenge ,preis as sumpreis ,mwst, EKPR "
    sSQL = sSQL & ", linr, lpz from kassjour where"
    sSQL = sSQL & " Month(adate) = " & iMonat
    sSQL = sSQL & " and Year(adate) = " & iJahr
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = "Update NEGERTRAG Set NUMSATZ = sumpreis * 100/ (100 + " & gdMWStE & ") where "
    sSQL = sSQL & " MWST = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 37
    
    sSQL = "Update NEGERTRAG Set NUMSATZ = sumpreis * 100/ (100 + " & gdMWStV & ") where "
    sSQL = sSQL & " MWST = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 43
    
    sSQL = "Update NEGERTRAG Set NUMSATZ = sumpreis * 100/ (100 + " & gdMWStO & ") where "
    sSQL = sSQL & " MWST = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 58
    
    sSQL = "Update NEGERTRAG Set Ertrag =  0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 61
    
    sSQL = "Update NEGERTRAG Set Ertrag =  NUMSATZ - (EKPR * summenge) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 65
    
    loeschNEW "NEGERTRAGPR", gdBase
    CreateTableT2 "NEGERTRAGPR", gdBase
    
    sSQL = "Insert into NEGERTRAGPR Select artnr, bezeich ,sum(ertrag) as sumertrag, sum(summenge) as menge ,sum(sumpreis) as preis ,mwst ,EKPR "
    sSQL = sSQL & ", linr, lpz from NEGERTRAG "
    sSQL = sSQL & " where ertrag <= 0"
    sSQL = sSQL & " group by artnr, bezeich, mwst ,EKPR "
    sSQL = sSQL & ", linr, lpz "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78
    
    sSQL = "Update NEGERTRAGPR inner join ARTLIEF on NEGERTRAGPR.Linr = ARTLIEF.Linr and NEGERTRAGPR.ARTNR = ARTLIEF.ARTNR "
    sSQL = sSQL & " set NEGERTRAGPR.LIBESNR = ARTLIEF.LIBESNR "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 87
        
    sSQL = "Update NEGERTRAGPR inner join lisrt on NEGERTRAGPR.linr = lisrt.linr "
    sSQL = sSQL & " Set NEGERTRAGPR.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 91
    
    sSQL = "Update NEGERTRAGPR inner join Artikel on NEGERTRAGPR.artnr = Artikel.artnr "
    sSQL = sSQL & " Set NEGERTRAGPR.KVKPR1 = Artikel.KVKPR1 "
    sSQL = sSQL & " , NEGERTRAGPR.FARBNR = val(Artikel.awm) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 95
    
    BringFarbeInsSpiel "NEGERTRAGPR", gdBase
    
    loeschNEW "NEGERTRAGK", gdBase
    CreateTableT2 "NEGERTRAGK", gdBase
    
    sSQL = "Insert into NEGERTRAGK (Datum) values ('" & sDatum & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("NEGERTRAGPR", gdBase) Then
        reportbildschirm "", "aWKL55g"
    Else
        anzeige "rot", "Es sind keine Daten vorhanden.", Label9
    End If
    
    txtStatus.Text = 100
   
    Screen.MousePointer = 0
    picprogress.Visible = False

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "negativErtrag"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Farbmerkmalsliste(sKrit As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "ART38", gdBase
    CreateTable "ART38", gdBase
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    sSQL = " Insert into ART38 select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "
    sSQL = sSQL & " , RABATT_OK "
    sSQL = sSQL & " , BONUS_OK "
    sSQL = sSQL & " , AWM "
    sSQL = sSQL & " , '' as FARBText "
    sSQL = sSQL & " , val(awm) as FARBNR "
    sSQL = sSQL & " , '' as liefbez "
    sSQL = sSQL & "   from artikel "
    
    If sKrit = "alle" Then
    
    Else
        sSQL = sSQL & "   where bestand > 0 "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 15
    
    sSQL = "Update art38 inner join lisrt on art38.linr = lisrt.linr "
    sSQL = sSQL & " Set art38.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 18
    
    sSQL = "Create index Farbnr on art38(farbnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 27
    
    sSQL = "Delete from art38 where Farbnr = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 43
    
    diefarbezeichnung "ART38", gdBase

    txtStatus.Text = 98
    
    Screen.MousePointer = 0
    reportbildschirm "", "aZEN00C"
    
    txtStatus.Text = 0
    picprogress.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Farbmerkmalsliste"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub imbestandundNoVerkauft()
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
    
    anzeigeNew "normal", "Artikeldaten werden ermittelt...", Label9
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "Kopf55", gdBase
    CreateTable "KOPF55", gdBase
    
    loeschNEW "ART55", gdBase
    CreateTable "ART55", gdBase

    sSQL = " Insert into ART55 select  ARTNR"
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
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 180
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 40

    

    sSQL = "Delete from art55 where bestand <= 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 80

    sSQL = "Delete from art55 where bestand is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 0




    Set rsrs = gdBase.OpenRecordset("ART55")


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
                datLVK = ErmlzVK(cART)
                datLZU = ErmlzZugang(cART)

                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                Select Case ldifferenz
                        
                    Case Is > 274
                    
                    
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 9 Monaten"
                        End If
                        
                    Case Is > 243
                        ctmp = "seit 8 Monaten"
                    Case Is > 213
                        ctmp = "seit 7 Monaten"
                    Case Is > 182

                        ctmp = "seit 6 Monaten"
                    Case Is > 152
                        ctmp = "seit 5 Monaten"
                    Case Is > 121
                        ctmp = "seit 4 Monaten"
                    Case Is > 91
                        ctmp = "seit 3 Monaten"
                    Case Is > 61
                        ctmp = "seit 2 Monaten"
                    Case Is > 30
                        ctmp = "seit 1 Monat"
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

    sSQL = "Delete from art55 where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from art55 where Monat is null "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update art55 inner join lisrt on art55.linr = lisrt.linr "
    sSQL = sSQL & " Set art55.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    'jetzt noch Kopfdaten ermitteln
    
    Dim GesBestand      As Long
    Dim TeiBestand      As Long
    Dim SEKWertg        As Double
    Dim LEKWertg        As Double
    Dim KVKWertg        As Double
    Dim SEKWertT        As Double
    Dim LEKWertT        As Double
    Dim KVKWertT        As Double
    
    Dim sSEKWertg        As String
    Dim sLEKWertg        As String
    Dim sKVKWertg        As String
    Dim sSEKWertT        As String
    Dim sLEKWertT        As String
    Dim sKVKWertT        As String
    
    GesBestand = ermgesbestand()
    TeiBestand = ermTeibestand(txtStatus, picprogress, "ART55")
    
    GesBestand = GesBestand - TeiBestand

    SEKWertg = ermSEKWERT(txtStatus, picprogress)
    LEKWertg = ermLEKWERT(txtStatus, picprogress)
    KVKWertg = ermKVKWERT(txtStatus, picprogress)

    SEKWertT = ermSEKWERTT(txtStatus, picprogress, "ART55")
    LEKWertT = ermLEKWERTt(txtStatus, picprogress, "ART55")
    KVKWertT = ermKVKWERTt(txtStatus, picprogress, "ART55")
    
    SEKWertg = SEKWertg - SEKWertT
    LEKWertg = LEKWertg - LEKWertT
    KVKWertg = KVKWertg - KVKWertT
    
    sSEKWertg = CStr(SEKWertg)
    sSEKWertg = SwapStr(sSEKWertg, ",", ".")
    
    sLEKWertg = CStr(LEKWertg)
    sLEKWertg = SwapStr(sLEKWertg, ",", ".")
    
    sKVKWertg = CStr(KVKWertg)
    sKVKWertg = SwapStr(sKVKWertg, ",", ".")
    
    
    sSEKWertT = CStr(SEKWertT)
    sSEKWertT = SwapStr(sSEKWertT, ",", ".")
    
    sLEKWertT = CStr(LEKWertT)
    sLEKWertT = SwapStr(sLEKWertT, ",", ".")
    
    sKVKWertT = CStr(KVKWertT)
    sKVKWertT = SwapStr(sKVKWertT, ",", ".")

    
    sSQL = "Insert into Kopf55 (RestBestand,Auswahl,SEKWertg,SEKWertt,LEKWertg,LEKWertt,KVKWertg,KVKWertt )"
    sSQL = sSQL & " values ( "
    sSQL = sSQL & GesBestand
    sSQL = sSQL & "," & TeiBestand
    sSQL = sSQL & "," & sSEKWertg
    sSQL = sSQL & "," & sSEKWertT
    sSQL = sSQL & "," & sLEKWertg
    sSQL = sSQL & "," & sLEKWertT
    sSQL = sSQL & "," & sKVKWertg
    sSQL = sSQL & "," & sKVKWertT
    
    sSQL = sSQL & " ) "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    
    txtStatus.Text = 0
    picprogress.Visible = False

    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9

    Screen.MousePointer = 0
    reportbildschirm "sdsd", "aWKL55a"
    
    anzeigeNew "normal", "", Label9
    
    
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "imbestandundNoVerkauft"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub imbestandundKVKNULL()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    anzeigeNew "normal", "Artikeldaten werden ermittelt...", Label9
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55H", gdBase
    CreateTableT2 "ART55H", gdBase
    
    cSQL = "Update Artikel set KVKPR1 = 0 where kvkpr1 is null"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = " Insert into ART55H select ARTNR"
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", LINR  "
    cSQL = cSQL & ", LPZ  "
    cSQL = cSQL & ", LIBESNR  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", '' as liefbez  "
    cSQL = cSQL & " from Artikel "
    cSQL = cSQL & " where BESTAND > 0 and KVKPR1 <= 0 "
    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = 40

    cSQL = "Update art55h inner join lisrt on art55h.linr = lisrt.linr "
    cSQL = cSQL & " Set art55h.liefbez = lisrt.liefbez "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9

    Screen.MousePointer = 0
    reportbildschirm "sdsd", "aWKL55h"
    
    anzeigeNew "normal", "", Label9
    
    
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "imbestandundKVKNULL"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub ArtikelnichtBonus()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    anzeigeNew "normal", "Artikeldaten werden ermittelt...", Label9
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55H", gdBase
    CreateTableT2 "ART55H", gdBase

    cSQL = " Insert into ART55H select ARTNR"
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", LINR  "
    cSQL = cSQL & ", LPZ  "
    cSQL = cSQL & ", LIBESNR  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", '' as liefbez  "
    cSQL = cSQL & " from Artikel "
    cSQL = cSQL & " where Bonus_OK ='N' "
    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = 40

    cSQL = "Update art55h inner join lisrt on art55h.linr = lisrt.linr "
    cSQL = cSQL & " Set art55h.liefbez = lisrt.liefbez "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9

    Screen.MousePointer = 0
    reportbildschirm "sdsd", "aWKL55i"
    
    anzeigeNew "normal", "", Label9
    
    
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelnichtBonus"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub ArtikelohneMwst()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    anzeigeNew "normal", "Artikeldaten werden ermittelt...", Label9
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55H", gdBase
    CreateTableT2 "ART55H", gdBase

    cSQL = " Insert into ART55H select ARTNR"
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", LINR  "
    cSQL = cSQL & ", LPZ  "
    cSQL = cSQL & ", LIBESNR  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", '' as liefbez  "
    cSQL = cSQL & " from Artikel "
    cSQL = cSQL & " where MWST ='O' "
    cSQL = cSQL & " order by BEZEICH "
    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = 40

    cSQL = "Update art55h inner join lisrt on art55h.linr = lisrt.linr "
    cSQL = cSQL & " Set art55h.liefbez = lisrt.liefbez "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 0
    picprogress.Visible = False

    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9

    Screen.MousePointer = 0
    reportbildschirm "sdsd", "aWKL55j"
    
    anzeigeNew "normal", "", Label9
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelohneMwst"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Function ErmlzVK(cART As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmlzVK = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Kassjour where ARTNR = " & cART & " "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVK = rsINB!MaxDate
        End If
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmlzVK"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermTeibestand(txtStatus As TextBox, picprogress As PictureBox, sTab As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    ermTeibestand = 0
    
    txtStatus.Text = 40
    
    Set rsrs = gdBase.OpenRecordset("select sum(Bestand) as Maxi from " & sTab & " ")
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermTeibestand = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermTeibestand"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermSEKWERTT(txtStatus As TextBox, picprogress As PictureBox, sTab As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermSEKWERTT = 0

    loeschNEW "Lieflw6", gdBase
    CreateTable "LIEFLW6", gdBase
    loeschNEW "ArtTemp4", gdBase
    
    txtStatus.Text = 71

    sSQL = "select * into ArtTemp4 from " & sTab
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 72
    

    
    sSQL = "Update ArtTemp4 inner join artikel on ArtTemp4.artnr = artikel.artnr "
    sSQL = sSQL & " set ArtTemp4.lekpr = artikel.lekpr where ArtTemp4.lekpr = 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78

    sSQL = "INSERT into LIEFLW6 Select LINR, Sum(ArtTemp4.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(KVKPR1* ArtTemp4.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(lEKPR* ArtTemp4.BESTAND) as LagerEK"
    sSQL = sSQL & " from ArtTemp4 "
    sSQL = sSQL & " Where ArtTemp4.Bestand > 0  "
    sSQL = sSQL & " group BY ArtTemp4.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ArtTemp4", gdBase
    
    txtStatus.Text = 80
    
    sSQL = "select sum(lagerEK) as maxi from LIEFLW6 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
        
            ermSEKWERTT = rsrs!maxi
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermSEKWERTT"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function
Private Function ermSEKWERT(txtStatus As TextBox, picprogress As PictureBox) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermSEKWERT = 0

    loeschNEW "Lieflw1", gdBase
    CreateTable "LIEFLW1", gdBase
    loeschNEW "ArtTemp5", gdBase

    txtStatus.Text = 45

    sSQL = "select * into ArtTemp5 from artikel "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 50
    

    
    sSQL = "Update ArtTemp5 inner join artikel on ArtTemp5.artnr = artikel.artnr "
    sSQL = sSQL & " set ArtTemp5.ekpr = artikel.lekpr where ArtTemp5.ekpr = 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 55


    
    sSQL = "INSERT into LIEFLW1 Select LINR, Sum(ArtTemp5.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(KVKPR1* ArtTemp5.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(EKPR* ArtTemp5.BESTAND) as LagerEK"
    sSQL = sSQL & " from ArtTemp5 "
    sSQL = sSQL & " Where ArtTemp5.Bestand > 0  "
    sSQL = sSQL & " group BY ArtTemp5.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ArtTemp5", gdBase
    
    txtStatus.Text = 58
    
    sSQL = "select sum(lagerEK) as maxi from LIEFLW1 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermSEKWERT = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "Lieflw1", gdBase

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermSEKWERT"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function
Private Function ermKVKWERT(txtStatus As TextBox, picprogress As PictureBox) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermKVKWERT = 0

    loeschNEW "Lieflw2", gdBase
    CreateTable "LIEFLW2", gdBase
    loeschNEW "ArtTemp6", gdBase
    
    txtStatus.Text = 67

    sSQL = "select * into ArtTemp6 from artikel "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    
    txtStatus.Text = 68
    
    sSQL = "Update ArtTemp6 inner join artikel on ArtTemp6.artnr = artikel.artnr "
    sSQL = sSQL & " set ArtTemp6.ekpr = artikel.lekpr where ArtTemp6.ekpr = 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 69

    sSQL = "INSERT into LIEFLW2 Select LINR, Sum(ArtTemp6.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(KVKPR1* ArtTemp6.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(EKPR* ArtTemp6.BESTAND) as LagerEK"
    sSQL = sSQL & " from ArtTemp6 "
    sSQL = sSQL & " Where ArtTemp6.Bestand > 0  "
    sSQL = sSQL & " group BY ArtTemp6.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70
    
    loeschNEW "ArtTemp6", gdBase
    
    sSQL = "select sum(LagerVK) as maxi from LIEFLW2 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKVKWERT = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKVKWERT"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function
Private Function ermLEKWERT(txtStatus As TextBox, picprogress As PictureBox) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermLEKWERT = 0

    loeschNEW "Lieflw3", gdBase
    CreateTable "LIEFLW3", gdBase
    loeschNEW "ArtTemp7", gdBase
    
    txtStatus.Text = 60
    
    sSQL = "select * into ArtTemp7 from artikel "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 61
    


    sSQL = "INSERT into LIEFLW3 Select a.LINR, Sum(a.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(a.KVKPR1* a.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(b.lEKPR* a.BESTAND) as LagerEK"
    sSQL = sSQL & " from  ArtTemp7 A inner join artlief B on B.artnr = A.artnr "
    sSQL = sSQL & " and a.linr = b.linr "
    sSQL = sSQL & " Where a.Bestand > 0  "
    sSQL = sSQL & " group BY a.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 62

    sSQL = "select sum(lagerEK) as maxi from LIEFLW3 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
        
            ermLEKWERT = rsrs!maxi
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermLEKWERT"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function
Private Function ermKVKWERTt(txtStatus As TextBox, picprogress As PictureBox, sTab As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermKVKWERTt = 0

    loeschNEW "Lieflw4", gdBase
    CreateTable "LIEFLW4", gdBase
    loeschNEW "ArtTemp8", gdBase

    txtStatus.Text = 93

    sSQL = "select * into ArtTemp8 from " & sTab
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 97
    
    sSQL = "Update ArtTemp8 inner join artikel on ArtTemp8.artnr = artikel.artnr "
    sSQL = sSQL & " set ArtTemp8.lekpr = artikel.lekpr where ArtTemp8.lekpr = 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 98

    sSQL = "INSERT into LIEFLW4 Select LINR, Sum(ArtTemp8.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(KVKPR1* ArtTemp8.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(lEKPR* ArtTemp8.BESTAND) as LagerEK"
    sSQL = sSQL & " from ArtTemp8 "
    sSQL = sSQL & " Where ArtTemp8.Bestand > 0  "
    sSQL = sSQL & " group BY ArtTemp8.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ArtTemp8", gdBase
    
    txtStatus.Text = 99
    
    sSQL = "select sum(LagerVK) as maxi from LIEFLW4 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
        
            ermKVKWERTt = rsrs!maxi
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKVKWERTt"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function
Private Function ermLEKWERTt(txtStatus As TextBox, picprogress As PictureBox, sTab As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermLEKWERTt = 0

    loeschNEW "Lieflw5", gdBase
    CreateTable "LIEFLW5", gdBase
    loeschNEW "ArtTemp9", gdBase
    
    txtStatus.Text = 83
    
    sSQL = "select * into ArtTemp9 from " & sTab
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 89

    sSQL = "INSERT into LIEFLW5 Select a.LINR, Sum(a.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(a.KVKPR1* a.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(b.lEKPR* a.BESTAND) as LagerEK"
    sSQL = sSQL & " from  ArtTemp9 A inner join artlief B on B.artnr = A.artnr "
    sSQL = sSQL & " and a.linr = b.linr "
    sSQL = sSQL & " Where a.Bestand > 0  "
    sSQL = sSQL & " group BY a.LINR "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 91

    sSQL = "select sum(lagerEK) as maxi from LIEFLW5 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
        
            ermLEKWERTt = rsrs!maxi
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermLEKWERTt"
    Fehler.gsFehlertext = "Im Programmteil Diverse Artikellisten ist ein Fehler aufgetreten."
    Fehlermeldung1
End Function

