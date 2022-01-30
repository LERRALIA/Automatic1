VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL127 
   Caption         =   "Artikel Export"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL127.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
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
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   37
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Index           =   6
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   36
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   32
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   6840
      MaxLength       =   5
      TabIndex        =   26
      Text            =   "0"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   23
      Text            =   "0"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   21
      Top             =   5040
      Width           =   1215
   End
   Begin sevCommand3.Command Command5 
      Height          =   855
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "alle ""Shop"" -Artikel in Excel exportieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   355
      Index           =   4
      Left            =   10560
      TabIndex        =   19
      Top             =   3600
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      ButtonStyle     =   2
      Caption         =   "L"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   7560
      TabIndex        =   17
      Top             =   1440
      Width           =   3375
   End
   Begin sevCommand3.Command Command5 
      Height          =   355
      Index           =   7
      Left            =   10560
      TabIndex        =   15
      Top             =   960
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      ButtonStyle     =   2
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   9480
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   9195
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   9255
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "Pennerartikel in Excel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "alle Artikel als CSV exportieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   10
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
   Begin VB.TextBox Text1 
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
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "0"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin sevCommand3.Command Command5 
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "alle Artikel mit Bestand in Excel exportieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "alle Artikel in Excel exportieren"
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "alle EX - Artikel in Excel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   29
      Top             =   4320
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "neue Artikel (Artnr,EAN)"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   360
      Index           =   10
      Left            =   5880
      TabIndex        =   30
      ToolTipText     =   "Kalender"
      Top             =   4440
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
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   360
      Index           =   12
      Left            =   5880
      TabIndex        =   34
      Top             =   3840
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      ButtonStyle     =   2
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   360
      Index           =   13
      Left            =   5880
      TabIndex        =   35
      Top             =   5040
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      ButtonStyle     =   2
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   14
      Left            =   120
      TabIndex        =   38
      Top             =   6600
      Width           =   2655
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
      ButtonStyle     =   2
      Caption         =   "Artikel in Excel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   360
      Index           =   15
      Left            =   5880
      TabIndex        =   39
      Top             =   6720
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
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
      ButtonStyle     =   2
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label10 
      Caption         =   "Lieferant"
      Height          =   255
      Left            =   3000
      TabIndex        =   40
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Lieferant"
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "angelegt seit:"
      Height          =   255
      Left            =   3000
      TabIndex        =   31
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Abschlag der Kassenpreise in %"
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   27
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Zentriert
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "nur mit Bestand"
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Lieferant"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Anzahl Lieferanten"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Lieferanten ausschließen"
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Abschlag der Kassenpreise in %"
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "nur mit Bestand"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   495
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
      Caption         =   "Artikel Export"
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
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL127
        Case 1
            nurmitBestand Text1(0).Text, Text1(1).Text
        Case 2 'csv Export alle Artikel
            CSVExportalleArt
        Case 3
            ExcelExportallePennerArt
        Case 4
            List3.Clear
            Label4(2).Visible = False
            Command5(4).Visible = False
        Case 5
            alleShopArtikel Text1(3).Text, Text1(4).Text, Text1(2).Text
        Case 6 'excel Export alle Artikel
            ExcelExportalleArt
        Case 7
            Liefauswahl
        Case 11
            gsHelpstring = "Artikel Export"
            frmWKL110.Show 1
        Case 8 'excel Export EX Artikel
            ExcelExportEX_Art Text1(5).Text
        Case 9 'csv Export Artikel, ab angelegt seit
            CSVExportNeueArt Text1(6).Text
        Case 10
            Text1(6).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
        Case 12
            Text1_KeyUp 5, vbKeyF2, 0
        Case 13
            Text1_KeyUp 2, vbKeyF2, 0
        Case 15
            Text1_KeyUp 7, vbKeyF2, 0
        Case 14 'excel Export avenue
            ExcelExportAvenue Text1(7).Text
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub CSVExportalleArt()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim iRet        As Integer
    Dim rsrs        As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
   
    Dim lPos            As Long
    Dim cSatz           As String
    Dim sSpalte(17) As String
    Dim i               As Integer
    
    frmWKL69.Show 1
        
    Select Case Month(DateValue(Now))
        Case 1
            If gsZeitPass <> "Schlümpfe" Then
                Exit Sub
            End If
        Case 2
            If gsZeitPass <> "Single" Then
                Exit Sub
            End If
        Case 3
            If gsZeitPass <> "Ölschock" Then
                Exit Sub
            End If
        Case 4
            If gsZeitPass <> "Zweierkiste" Then
                Exit Sub
            End If
        Case 5
            If gsZeitPass <> "Mitte" Then
                Exit Sub
            End If
        Case 6
            If gsZeitPass <> "Wende" Then
                Exit Sub
            End If
        Case 7
            If gsZeitPass <> "Havarie" Then
                Exit Sub
            End If
        Case 8
            If gsZeitPass <> "Waldsterben" Then
                Exit Sub
            End If
        Case 9
            If gsZeitPass <> "Molkepulver" Then
                Exit Sub
            End If
        Case 10
            If gsZeitPass <> "Tiefflug" Then
                Exit Sub
            End If
        Case 11
            If gsZeitPass <> "Realo" Then
                Exit Sub
            End If
        Case 12
            If gsZeitPass <> "Eurogeld" Then
                Exit Sub
            End If
    End Select
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
            
    sSpalte(0) = "Artnr"
    sSpalte(1) = "Bezeich"
    sSpalte(2) = "Libesnr"
    sSpalte(3) = "EAN"
    sSpalte(4) = "EAN2"
    sSpalte(5) = "EAN3"
    sSpalte(6) = "BESTAND"
    sSpalte(7) = "LEKPR"
    sSpalte(8) = "KVKPR1"
    sSpalte(9) = "VKPR"
    sSpalte(10) = "MWST"
    sSpalte(11) = "LINR"
    sSpalte(12) = "gefuehrt"
    sSpalte(13) = "RKZ"
    sSpalte(14) = "PREISSCHU"
    sSpalte(15) = "FARBTEXT"
    sSpalte(16) = "MARKE"
    sSpalte(17) = "LINBEZ"
    
    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select ARTIKEL.Artnr "
    sSQL = sSQL & " , ARTIKEL.Bezeich "
    sSQL = sSQL & " , ARTIKEL.Libesnr"
    sSQL = sSQL & " , ARTIKEL.EAN"
    sSQL = sSQL & " , ARTIKEL.EAN2"
    sSQL = sSQL & " , ARTIKEL.EAN3"
    sSQL = sSQL & " , ARTIKEL.BESTAND"
    sSQL = sSQL & " , ARTLIEF.LEKPR"
    sSQL = sSQL & " , ARTIKEL.KVKPR1"
    sSQL = sSQL & " , ARTIKEL.VKPR"
    sSQL = sSQL & " , ARTIKEL.MWST"
    sSQL = sSQL & " , ARTLIEF.LINR"
    sSQL = sSQL & " , ARTIKEL.LPZ"
    sSQL = sSQL & " , ARTIKEL.gefuehrt"
    sSQL = sSQL & " , ARTIKEL.RKZ"
    sSQL = sSQL & " , ARTIKEL.PREISSCHU"
    sSQL = sSQL & " , val(ARTIKEL.AWM) as FARBNR "
    sSQL = sSQL & " , '' as  FARBTEXT "
    sSQL = sSQL & " , '' as  LINBEZ "
    sSQL = sSQL & " , '' as  Marke "
    sSQL = sSQL & " ,0 as FARBwert "
    sSQL = sSQL & " ,0 as FARBwertS "
    sSQL = sSQL & " into ArtExcA from ARTIKEL inner join Artlief on ARTIKEL.artnr = ARTLIEF.ARTNR and ARTIKEL.linr = ARTLIEF.linr "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update ArtExcA"
'    sSQL = sSQL & " set EAN = '' "
'    sSQL = sSQL & " , EAN2 = '' "
'    sSQL = sSQL & " , EAN3 = '' "
'    gdBase.Execute sSQL, dbFailOnError
    
    If List3.ListCount > 0 Then
        For i = 0 To List3.ListCount - 1
            sSQL = " Delete from ArtExcA where linr = " & Val(Left(List3.list(i), 6))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If
    
    BringFarbeInsSpiel "ArtExcA", gdBase
    
    Markenabgleich "ArtExcA", gdBase

    
    sSQL = " Select "
    For i = 0 To 17
        If i > 0 Then
            sSQL = sSQL & "," & sSpalte(i)
        Else
            sSQL = sSQL & sSpalte(i)
        End If

    Next i
    sSQL = sSQL & " from ArtExcA "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        sAusgabedatname = "alleArtikel" & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = ""
        For i = 0 To 17
            If i > 0 Then
                cSatz = cSatz & ";" & sSpalte(i)
            Else
                cSatz = cSatz & sSpalte(i)
            End If
    
        Next i
        cSatz = cSatz & Chr$(13) & Chr$(10)
        

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 17
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        cSatz = cSatz & ";" & rsrs.Fields(i)
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        gcBestellEmail.Attachment1 = cdatei
        Screen.MousePointer = 0
        frmWKL129.Show 1
    Else
        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSVExportalleArt"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Liefauswahl()
On Error GoTo LOKAL_ERROR
    
    Dim sAuswahlfeld As String
    Dim ctmp As String
    Dim lcount As Long
    
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = True
    
    
    gF2Prompt.cFeld = "LINR"
    If gF2Prompt.cFeld <> "" Then
        frmWK00a.Show 1
    End If
    
    List3.Visible = False
    List3.Clear
    For lcount = 0 To 100
        If gF2Prompt.cArray(lcount) <> "" Then
            List3.Visible = True
            If gF2Prompt.cArray(lcount) <> "" Then
                List3.AddItem gF2Prompt.cArray(lcount)
            End If
        End If
    Next lcount
            
    If List3.Visible = True Then
        Label4(2).Visible = True
        Label4(2).Caption = List3.ListCount & " Lieferanten"
        Label4(2).Refresh
        Command5(4).Visible = True
    Else
        Label4(2).Visible = False
        Command5(4).Visible = False
    End If
            
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Liefauswahl"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExcelExportalleArt()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim cDatname    As String
    Dim i           As Integer
    
    cDatname = "alleArtikel" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    frmWKL69.Show 1
        
    Select Case Month(DateValue(Now))
        Case 1
            If gsZeitPass <> "Schlümpfe" Then
                Exit Sub
            End If
        Case 2
            If gsZeitPass <> "Single" Then
                Exit Sub
            End If
        Case 3
            If gsZeitPass <> "Ölschock" Then
                Exit Sub
            End If
        Case 4
            If gsZeitPass <> "Zweierkiste" Then
                Exit Sub
            End If
        Case 5
            If gsZeitPass <> "Mitte" Then
                Exit Sub
            End If
        Case 6
            If gsZeitPass <> "Wende" Then
                Exit Sub
            End If
        Case 7
            If gsZeitPass <> "Havarie" Then
                Exit Sub
            End If
        Case 8
            If gsZeitPass <> "Waldsterben" Then
                Exit Sub
            End If
        Case 9
            If gsZeitPass <> "Molkepulver" Then
                Exit Sub
            End If
        Case 10
            If gsZeitPass <> "Tiefflug" Then
                Exit Sub
            End If
        Case 11
            If gsZeitPass <> "Realo" Then
                Exit Sub
            End If
        Case 12
            If gsZeitPass <> "Eurogeld" Then
                Exit Sub
            End If
    End Select
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select ARTIKEL.Artnr "
    sSQL = sSQL & " , ARTIKEL.Bezeich "
    sSQL = sSQL & " , ARTIKEL.Libesnr"
    sSQL = sSQL & " , ARTIKEL.EAN"
    sSQL = sSQL & " , ARTIKEL.EAN2"
    sSQL = sSQL & " , ARTIKEL.EAN3"
    sSQL = sSQL & " , ARTIKEL.BESTAND"
    sSQL = sSQL & " , ARTLIEF.LEKPR"
    sSQL = sSQL & " , ARTIKEL.KVKPR1"
    sSQL = sSQL & " , ARTIKEL.VKPR"
    sSQL = sSQL & " , ARTIKEL.MWST"
    sSQL = sSQL & " , ARTLIEF.LINR"
    sSQL = sSQL & " , ARTIKEL.gefuehrt"
    sSQL = sSQL & " , ARTIKEL.RKZ"
    sSQL = sSQL & " , '' as Artikelstatus "
    sSQL = sSQL & " into ArtExcA from ARTIKEL inner join Artlief on ARTIKEL.artnr = ARTLIEF.ARTNR and ARTIKEL.linr = ARTLIEF.linr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA"
    sSQL = sSQL & " set EAN = '' "
    sSQL = sSQL & " , EAN2 = '' "
    sSQL = sSQL & " , EAN3 = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    If List3.ListCount > 0 Then
        For i = 0 To List3.ListCount - 1
            sSQL = " Delete from ArtExcA where linr = " & Val(Left(List3.list(i), 6))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If

    cdatei = cPfad1 & "BOX\" & cDatname
    cPfad = cPfad1 & "BOX"

    sSQL = "Select * into ArtExcA IN '" & cdatei & "' 'Excel 8.0;' from ArtExcA "
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
    loeschNEW "ArtExcA", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExportalleArt"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExcelExportEX_Art(sLiefnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim cDatname    As String
    Dim i           As Integer
    
    cDatname = "ExArtikel" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select ARTIKEL.Artnr "
    sSQL = sSQL & " , ARTIKEL.Bezeich "
    sSQL = sSQL & " , ARTIKEL.Libesnr"
    sSQL = sSQL & " , ARTIKEL.EAN"
    sSQL = sSQL & " , ARTIKEL.EAN2"
    sSQL = sSQL & " , ARTIKEL.EAN3"
    sSQL = sSQL & " , ARTIKEL.BESTAND"
    sSQL = sSQL & " , ARTLIEF.LEKPR"
    sSQL = sSQL & " , ARTIKEL.KVKPR1"
    sSQL = sSQL & " , ARTIKEL.VKPR"
    sSQL = sSQL & " , ARTIKEL.MWST"
    sSQL = sSQL & " , ARTLIEF.LINR"
    sSQL = sSQL & " , ARTIKEL.gefuehrt"
    sSQL = sSQL & " , ARTIKEL.RKZ"
    sSQL = sSQL & " , '' as Artikelstatus "
    sSQL = sSQL & " into ArtExcA from ARTIKEL inner join Artlief on ARTIKEL.artnr = ARTLIEF.ARTNR and ARTIKEL.linr = ARTLIEF.linr "
    sSQL = sSQL & " where ARTLIEF.RKZ = 'J' "
    
    If sLiefnr <> "" Then
        sSQL = sSQL & " and  ARTLIEF.linr = " & sLiefnr & " "
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    


    cdatei = cPfad1 & "BOX\" & cDatname
    cPfad = cPfad1 & "BOX"

    sSQL = "Select * into ArtExcA IN '" & cdatei & "' 'Excel 8.0;' from ArtExcA "
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
    loeschNEW "ArtExcA", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExportEX_Art"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExcelExportAvenue(sLiefnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim cDatname    As String
    Dim i           As Integer
    
    If sLiefnr = "" Then
        Exit Sub
    End If
    
    cDatname = "Produkte_" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select ARTIKEL.Artnr as ID "
    sSQL = sSQL & " , ARTIKEL.EAN"
    sSQL = sSQL & " , ARTIKEL.Bezeich as Produktname"
    sSQL = sSQL & " , ARTIKEL.Bezeich as [Produkt-Gruppe] "
    sSQL = sSQL & " , '19' as [Steuersatz in %] "
    sSQL = sSQL & " , ARTIKEL.Bezeich as Hersteller "
    sSQL = sSQL & " , Artikel.Bestand as Menge "
    sSQL = sSQL & " , Round(Artikel.KVKPR1,2) as [Preis in EUR]"
    sSQL = sSQL & " , Round(ARTLIEF.LEKPR,2) as [Einkaufspreis netto in EUR]"
    
    
    sSQL = sSQL & " , ARTIKEL.PGN"
    sSQL = sSQL & " , ARTIKEL.MWST"
    sSQL = sSQL & " , ARTLIEF.LINR"
    
    

    sSQL = sSQL & " into ArtExcA from ARTIKEL inner join Artlief on ARTIKEL.artnr = ARTLIEF.ARTNR and ARTIKEL.linr = ARTLIEF.linr "

    
    If sLiefnr <> "" Then
        sSQL = sSQL & " where  ARTLIEF.linr = " & sLiefnr & " "
        sSQL = sSQL & " and  ARTIKEL.gefuehrt = 'J' "
        sSQL = sSQL & " and  ARTIKEL.bestand > 0 "
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update ArtExcA set [Steuersatz in %] = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA set [Steuersatz in %] = '19' where mwst = 'V'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA set [Steuersatz in %] = '7' where mwst = 'E'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "Update ArtExcA set [Produkt-Gruppe] = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA inner join pgndbf on ArtExcA.PGN = PGNDBF.PGN "
    sSQL = sSQL & " Set  ArtExcA.[Produkt-Gruppe] = PGNDBF.PGNBEZEICH "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update ArtExcA set Hersteller = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA inner join lisrt on ArtExcA.linr = lisrt.linr "
    sSQL = sSQL & " Set ArtExcA.Hersteller = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    'Alter table
    sSQL = "alter table ArtExcA drop column LINR"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "alter table ArtExcA drop column PGN"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "alter table ArtExcA drop column MWST"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
'    sSQL = " Update Artikel inner join Artlief "
'    sSQL = sSQL & " on Artikel.artnr = Artlief.artnr "
'    sSQL = sSQL & " set Artikel.awm  = '" & cAWM & "' "
'    sSQL = sSQL & " where artlief.linr = " & lLinr
'    sSQL = sSQL & " and Round(Artikel.vkpr,2) <> Round(Artikel.KVKPR1,2) "
'    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    


    cdatei = cPfad1 & "BOX\" & cDatname
    cPfad = cPfad1 & "BOX"

    sSQL = "Select * into Produkte IN '" & cdatei & "' 'Excel 8.0;' from ArtExcA "
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
    loeschNEW "ArtExcA", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExportAvenue"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExcelExportallePennerArt()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim cDatname    As String
    Dim rsrs        As Recordset
    Dim cART        As String
    Dim ctmp        As String
    Dim datLVK      As Date
    Dim datLZU      As Date
    Dim lLastvk     As Long
    Dim lHeute      As Long
    Dim ldifferenz  As Long
    
    Dim lAnz        As Long
    Dim siAnzeige   As Single
    Dim i           As Integer
    
    lHeute = CLng(DateValue(Now))
    
    cDatname = "allePenner" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    anzeigeNew "normal", "Pennerartikel werden ermittelt...", Label1(4)
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART55C", gdBase
    CreateTable "ART55C", gdBase
    
    
    sSQL = "Update artikel set awm = 0 where awm is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into ART55C select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , BESTAND * KVKPR1 as KVKWERT "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "
    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
'    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 180
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 40
    
    
    sSQL = "Delete from ART55C where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    If List3.ListCount > 0 Then
        For i = 0 To List3.ListCount - 1
            sSQL = " Delete from ART55C where linr = " & Val(Left(List3.list(i), 6))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If
    
    txtStatus.Text = 45
    
    Set rsrs = gdBase.OpenRecordset("ART55C")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
            
                siAnzeige = siAnzeige + 1
                txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
            
                anzeige "normal", lcount & " Artikel noch...", Label1(4)
                lcount = lcount - 1
                
                rsrs.Edit
                rsrs!ERSTDAT = ErmFirstZugang(rsrs!artnr)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "neue Artikel entfernen", Label1(4)
    
    sSQL = " Delete from ART55C  where ERSTDAT > datevalue(now) - 180 "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 80

    sSQL = "Delete from ART55C where bestand is null "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Penner werden jetzt ermittelt...", Label1(4)
    txtStatus.Text = 0
    
    siAnzeige = 0

    Set rsrs = gdBase.OpenRecordset("ART55C")
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
'                datLZU = ErmlzZugang(cART)

                lLastvk = CLng(datLVK)
                ldifferenz = lHeute - lLastvk
                Select Case ldifferenz
                        
                    Case Is > 365
                    
                    
                        If ldifferenz = lHeute Then
                            ctmp = "(noch gar nicht)"
                        Else
                            ctmp = "seit 12 Monaten"
                        End If
                        
                    Case Else
                        ctmp = ""
                End Select

                rsrs!Monat = ctmp
                rsrs!lastvk = datLVK
'                rsrs!lastzu = datLZU
                rsrs.Update

            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Excelexport wird durchgeführt...", Label1(4)

    txtStatus.Text = 10

    sSQL = "Delete from ART55C where Monat = '' "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20

    sSQL = "Delete from ART55C where Monat is null "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 30

    sSQL = "Update ART55C inner join lisrt on ART55C.linr = lisrt.linr "
    sSQL = sSQL & " Set ART55C.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 0
    picprogress.Visible = False

    anzeigeNew "normal", "", Label1(4)

    Screen.MousePointer = 0
    

    cdatei = cPfad1 & "BOX\" & cDatname
    cPfad = cPfad1 & "BOX"

    sSQL = "Select * into ART55C IN '" & cdatei & "' 'Excel 8.0;' from ART55C "
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
    loeschNEW "ART55C", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExportallePennerArt"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub CSVExportNeueArt(sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim sSpalte(6) As String
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
            
    sSpalte(0) = "Artnr"
    sSpalte(1) = "EAN"
    sSpalte(2) = "EAN2"
    sSpalte(3) = "EAN3"

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select ARTIKEL.Artnr "
    sSQL = sSQL & " , ARTIKEL.EAN"
    sSQL = sSQL & " , ARTIKEL.EAN2"
    sSQL = sSQL & " , ARTIKEL.EAN3"
    sSQL = sSQL & " into ArtExcA from Artikel "
    
    If sdat <> "" Then
        sSQL = sSQL & " where aufdat > " & CLng(DateValue(sdat)) & " "
    End If
    gdBase.Execute sSQL, dbFailOnError
    

    sSQL = " Select "
    For i = 0 To 3
        If i > 0 Then
            sSQL = sSQL & "," & sSpalte(i)
        Else
            sSQL = sSQL & sSpalte(i)
        End If

    Next i
    sSQL = sSQL & " from ArtExcA "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        sAusgabedatname = "neueArtikel" & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "STAT\" & sAusgabedatname
        cPfad = cPfad1 & "STAT"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = ""
        For i = 0 To 3
            If i > 0 Then
                cSatz = cSatz & ";" & sSpalte(i)
            Else
                cSatz = cSatz & sSpalte(i)
            End If
    
        Next i
        cSatz = cSatz & Chr$(13) & Chr$(10)
        

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 3
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        cSatz = cSatz & ";" & rsrs.Fields(i)
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gbFtpYes Then
        giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
        frmWKL38.Show 1
    Else
        cPfad = gcDBPfad
        If Right$(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        cPfad = cPfad & "STAT\"

        gsAnzeigeText = "Die Datei '" & sAusgabedatname & "' ist unter: " & cPfad & " erstellt. Bitte übertragen Sie diese."
        frmWK21l.Show 1
    End If
    
'    iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
'    If iRet = vbYes Then
'        gcBestellEmail.Attachment1 = cdatei
'        Screen.MousePointer = 0
'        frmWKL129.Show 1
'    Else
'        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
'    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CSVExportNeueArt"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub nurmitBestand(sBestand As String, sAufAbschlag As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim iRet        As Integer
    Dim rsrs        As Recordset
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim lPos            As Long
    Dim cSatz           As String
    Dim cNettopreis     As String
    Dim cBruttopreis    As String
    Dim cRabattpreis    As String
    Dim cPreis          As String
    Dim cMarke          As String
    Dim cLinbez         As String
    Dim lBest           As Long
    Dim dAufAbschlag    As Double
    
    Screen.MousePointer = 11
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    lBest = 0
    If sBestand <> "" Then
        If IsNumeric(sBestand) Then
            lBest = CLng(sBestand)
        End If
    End If
    
    dAufAbschlag = 0
    If sAufAbschlag <> "" Then
        If IsNumeric(sAufAbschlag) Then
            dAufAbschlag = CDbl(sAufAbschlag)
        End If
    End If

    loeschNEW "ArtExcA", gdBase
    
    sSQL = "Select Artnr "
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , BESTAND"
    
    If dAufAbschlag <> 0 Then
        sSQL = sSQL & " , KVKPR1- (KVKPR1 * '" & dAufAbschlag & "'/100) as Rabattpreis"
    End If
    
    sSQL = sSQL & " , KVKPR1 as Bruttopreis"
    sSQL = sSQL & " , KVKPR1 as Nettopreis"
    sSQL = sSQL & " , MWST"
    sSQL = sSQL & " , LINR"
    sSQL = sSQL & " , '' as  MARKE"
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , '' as  LINBEZ"
    sSQL = sSQL & " , RKZ as geraeumt"
    sSQL = sSQL & " into ArtExcA from ARTIKEL "
    sSQL = sSQL & " where "
    sSQL = sSQL & " Bestand > " & lBest
    gdBase.Execute sSQL, dbFailOnError
    
    If List3.ListCount > 0 Then
        For i = 0 To List3.ListCount - 1
            sSQL = " Delete from ArtExcA where linr = " & Val(Left(List3.list(i), 6))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If
    
    sSQL = "Update ArtExcA set Nettopreis = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA set Nettopreis = (Bruttopreis*100)/(100 + " & gdMWStV & ") "
    sSQL = sSQL & " where MWST = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA set Nettopreis = (Bruttopreis*100)/(100 + " & gdMWStE & ") "
    sSQL = sSQL & " where MWST = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ArtExcA set Nettopreis = Bruttopreis "
    sSQL = sSQL & " where MWST = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    Markenabgleich "ArtExcA", gdBase
    
    Set rsrs = gdBase.OpenRecordset("Select * from ARTEXCA")
    If Not rsrs.EOF Then
        
        cdatei = cPfad1 & "BOX\products.csv"
        cPfad = cPfad1 & "BOX"
        Kill cdatei
        
        Dim iFileNr         As Integer
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBez = rsrs!BEZEICH
            End If
            
'            If Not IsNull(rsrs!nettopreis) Then
'                cNettopreis = rsrs!nettopreis
'            End If
            
            If dAufAbschlag <> 0 Then
                If Not IsNull(rsrs!Rabattpreis) Then
                    cPreis = Format(rsrs!Rabattpreis, "#####0.00")
                End If
            Else
                If Not IsNull(rsrs!bruttopreis) Then
                    cPreis = Format(rsrs!bruttopreis, "#####0.00")
                End If
            End If
            
            
            
            If Not IsNull(rsrs!MARKE) Then
                cMarke = rsrs!MARKE
            End If
            
            If Not IsNull(rsrs!linbez) Then
                cLinbez = rsrs!linbez
            End If
            
            cSatz = cArtNr & ";" & cBez & ";" & cPreis & ";" & cMarke & ";" & cLinbez
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
'    MsgBox "Datei 'products.csv' in " & cPfad & " erstellt!", vbInformation, "Winkiss Hinweis:"

    iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        gcBestellEmail.Attachment1 = cPfad1 & "BOX\products.csv"
        Screen.MousePointer = 0
        frmWKL129.Show 1
        
    Else
    
        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: products.csv abgespeichert", vbInformation, "Winkiss Information:"
    End If

    Screen.MousePointer = 0
    

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "nurmitBestand"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub alleShopArtikel(sBestand As String, sAufAbschlag As String, cUebergabeLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim cArtNr          As String
    Dim cBezeich        As String
    Dim lPos            As Long
    Dim cSatz           As String
    Dim cNettopreis     As String
    Dim cBruttopreis    As String
    Dim cRabattpreis    As String
    Dim cPreis          As String
    Dim cMarke          As String
    Dim cLinbez         As String
    Dim cBestand        As String
    Dim cEAN1           As String
    Dim cEAN2           As String
    Dim cEAN3           As String
    Dim cBildangabe     As String
    Dim cAgn            As String
    Dim cAGNBEZ         As String
    Dim cPGN            As String
    Dim cPGNBEZ         As String
    Dim cLinr           As String
    Dim cLINRBEZ        As String
    Dim cInhalt         As String
    Dim cInhaltBez      As String
    Dim cGRUNDPREIS     As String
    Dim cLiBesNr        As String
    Dim cLEKPR          As String
    Dim cSEK            As String
    Dim cRKZ            As String
    
    Dim cArtBez         As String
    Dim cINTERBEZ       As String
    Dim cBESCHREIB      As String
    Dim cSHOPKVK        As String
    Dim cKATEGORIE1     As String
    Dim cKATEGORIE2     As String
    Dim cMwst           As String
    
    Dim lBest           As Long
    Dim dAufAbschlag    As Double
    
    Dim sQuelle             As String
    Dim sZiel               As String
    
    Dim dGrundPreisDM       As Double
    Dim dGrundPreisEur      As Double
    Dim cGrundInhalt        As String
    
    Dim cGP                 As String
    Dim cGI                 As String
    
    Dim sQuellpfad          As String
    Dim sZielpfad           As String
    Dim lfail               As Long
    Dim lRet                As Long
    Dim lAnz                As Long
    
    If cUebergabeLinr = "" Then
        MsgBox "Bitte geben Sie einen Lieferanten an!", vbInformation, "Winkiss Information:"
        Exit Sub
    End If
    
    sQuellpfad = gcDBPfad
    sQuellpfad = ShortPath(sQuellpfad)
    If Right(sQuellpfad, 1) <> "\" Then
        sQuellpfad = sQuellpfad & "\"
    End If
    sQuellpfad = sQuellpfad & "PICTURE\ARTIKEL"
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cPfad = cPfad1 & "BOX"
    VerzVorhanden "Bilder", cPfad & "\"
    
    sZielpfad = gcDBPfad
    sZielpfad = ShortPath(sZielpfad)
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "BOX\Bilder"
    
    Kill sZielpfad & "\*.*"
    
    anzeige "normal", "Daten werden erstellt...", Label1(4)
    
    Screen.MousePointer = 11
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    Dim sPfad As String

    sPfad = gcDBPfad 'Bildpfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    lBest = 0
    If sBestand <> "" Then
        If IsNumeric(sBestand) Then
            lBest = CLng(sBestand)
        End If
    End If
    
    dAufAbschlag = 0
    If sAufAbschlag <> "" Then
        If IsNumeric(sAufAbschlag) Then
            dAufAbschlag = CDbl(sAufAbschlag)
        End If
    End If

    loeschNEW "ArtExcS", gdBase
    CreateTableT2 "ARTEXCS", gdBase
    
    sSQL = "Insert into ARTEXCS Select a.Artnr "
    sSQL = sSQL & " , a.Bezeich "
    
    If dAufAbschlag <> 0 Then
        sSQL = sSQL & " , a.KVKPR1- (a.KVKPR1 * '" & dAufAbschlag & "'/100) as Rabattpreis"
    End If
    
    sSQL = sSQL & " , a.KVKPR1 as Bruttopreis"
    sSQL = sSQL & " , a.KVKPR1 as Nettopreis"
    sSQL = sSQL & " , a.MWST"
    sSQL = sSQL & " , '' as  MARKE"
    sSQL = sSQL & " , a.LPZ "
    sSQL = sSQL & " , '' as  LINBEZ"
    sSQL = sSQL & " , a.RKZ as geraeumt"
    sSQL = sSQL & " , a.INHALT "
    sSQL = sSQL & " , a.INHALTBEZ "
    sSQL = sSQL & " , a.GRUNDPREIS "
    sSQL = sSQL & " , '' as EAN "
    sSQL = sSQL & " , '' as EAN2 "
    sSQL = sSQL & " , '' as EAN3 "
    sSQL = sSQL & " , a.AGN "
    sSQL = sSQL & " , '' as  AGNBEZ "
    sSQL = sSQL & " , a.PGN "
    sSQL = sSQL & " , '' as  PGNBEZ "
    sSQL = sSQL & " , a.BESTAND"
    sSQL = sSQL & " , b.LINR "
    sSQL = sSQL & " , '' as  LIEFBEZ "
    sSQL = sSQL & " , '' as  ARTBEZ "
    sSQL = sSQL & " , '' as  INTERBEZ "
    sSQL = sSQL & " , '' as  BESCHREIB "
    sSQL = sSQL & " , a.EKPR as  SHOPKVK "
    sSQL = sSQL & " , '' as  KATEGORIE1 "
    sSQL = sSQL & " , '' as  KATEGORIE2 "
    sSQL = sSQL & " , a.EKPR as  LEKPR "
    sSQL = sSQL & " , a.EKPR as  SEK "
    sSQL = sSQL & " , a.LIBESNR "
    sSQL = sSQL & "  from ARTIKEL   "
    
    sSQL = sSQL & "  a inner join Artlief B on A.artnr = B.artnr   "
    
    sSQL = sSQL & " where a.artnr in (Select artnr from interart)"
    sSQL = sSQL & " and a.Bestand > " & lBest
   
    sSQL = sSQL & " and b.LINR =  " & cUebergabeLinr
    
    gdBase.Execute sSQL, dbFailOnError

    
    sSQL = "Update ARTEXCS set Nettopreis = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEXCS set LEKPR = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEXCS set SHOPKVK = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEXCS set Nettopreis = (Bruttopreis*100)/(100 + " & gdMWStV & ") "
    sSQL = sSQL & " where MWST = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEXCS set Nettopreis = (Bruttopreis*100)/(100 + " & gdMWStE & ") "
    sSQL = sSQL & " where MWST = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEXCS set Nettopreis = Bruttopreis "
    sSQL = sSQL & " where MWST = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich EAN)...", Label1(4)
    
    'Abgleich EAN
    sSQL = "Update ARTEXCS inner join ARTIKEL on ARTEXCS.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " Set  ARTEXCS.EAN = ARTIKEL.EAN "
    sSQL = sSQL & " ,  ARTEXCS.EAN2 = ARTIKEL.EAN2 "
    sSQL = sSQL & " ,  ARTEXCS.EAN3 = ARTIKEL.EAN3 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich Artlief)...", Label1(4)
    
    'Abgleich artlief
    sSQL = "Update ARTEXCS inner join artlief on ARTEXCS.artnr = artlief.artnr "
    sSQL = sSQL & " Set  ARTEXCS.LEKPR = artlief.LEKPR "
    sSQL = sSQL & " ,  ARTEXCS.LIBESNR = artlief.LIBESNR "
    sSQL = sSQL & " where artlief.linr = " & cUebergabeLinr
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich AGN)...", Label1(4)
    
    'Abgleich agntext
    sSQL = "Update ARTEXCS inner join agndbf on ARTEXCS.AGN = AGNDBF.AGN "
    sSQL = sSQL & " Set  ARTEXCS.AGNBEZ = AGNDBF.AGTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich PGN)...", Label1(4)
    'Abgleich pgntext
    sSQL = "Update ARTEXCS inner join pgndbf on ARTEXCS.PGN = PGNDBF.PGN "
    sSQL = sSQL & " Set  ARTEXCS.PGNBEZ = PGNDBF.PGNBEZEICH "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich LINR)...", Label1(4)
    'Abgleich Liefbez
    sSQL = "Update ARTEXCS inner join LISRT on ARTEXCS.LINR = LISRT.LINR "
    sSQL = sSQL & " Set  ARTEXCS.Liefbez = LISRT.Liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich WEB)...", Label1(4)
    'Abgleich Interart
    sSQL = "Update ARTEXCS inner join INTERART on ARTEXCS.ARTNR = INTERART.ARTNR "
    sSQL = sSQL & " Set  ARTEXCS.ARTBEZ = INTERART.ARTBEZ "
    sSQL = sSQL & " ,  ARTEXCS.INTERBEZ = INTERART.INTERBEZ "
    sSQL = sSQL & " ,  ARTEXCS.BESCHREIB = INTERART.BESCHREIB "
    sSQL = sSQL & " ,  ARTEXCS.SHOPKVK = INTERART.SHOPKVK "
    sSQL = sSQL & " ,  ARTEXCS.KATEGORIE1 = INTERART.KATEGORIE "
    sSQL = sSQL & " ,  ARTEXCS.KATEGORIE2 = INTERART.KATEGORIE2 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Daten werden erstellt(Abgleich MARKE)...", Label1(4)
    Markenabgleich "ARTEXCS", gdBase
    
    Set rsrs = gdBase.OpenRecordset("Select * from ARTEXCS")
    If Not rsrs.EOF Then
        
        cdatei = cPfad1 & "BOX\products.csv"
        cPfad = cPfad1 & "BOX"
        
        
        Kill cdatei
        
        Dim iFileNr         As Integer
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "Artikelnummer;Artikelbezeichnung;Nettopreis;VKPreis;Marke;Produktlinie;Bestand;EAN1;EAN2;EAN3;Bild"
        cSatz = cSatz & ";Lieferantenbestellnummer;Liefinfo_geräumt;Listeneinkaufspreis;Schnitteinkaufspreis"
        cSatz = cSatz & ";Artikelgruppennummer;Artikelgruppenbezeichnung;Produktgruppennummer;Produktgruppenbezeichnung"
        cSatz = cSatz & ";Lieferantennummer;Lieferantenbezeichnung;Inhalt;Inhaltbezeichnung;Grundpreis"
        cSatz = cSatz & ";Grundpreis_Inhalt;Grundpreis_Preis"
        cSatz = cSatz & ";WEB_Bezeichnung;WEB_Kurztext;WEB_Beschreibung;WEB_Preis;WEB_KAT1;WEB_KAT2;MWST"
        
        cSatz = cSatz & Chr$(13) & Chr$(10)
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
    
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        
        
            cArtNr = ""
            cBez = ""
            cNettopreis = ""
            cPreis = ""
            cMarke = ""
            cLinbez = ""
            cBestand = ""
            cEAN1 = ""
            cEAN2 = ""
            cEAN3 = ""
            cBildangabe = ""
            cAgn = ""
            cAGNBEZ = ""
            cPGN = ""
            cPGNBEZ = ""
            cLinr = ""
            cLINRBEZ = ""
            cInhalt = ""
            cInhaltBez = ""
            cGRUNDPREIS = ""
            cGI = ""
            cGP = ""
            cArtBez = ""
            cINTERBEZ = ""
            cBESCHREIB = ""
            cSHOPKVK = ""
            cKATEGORIE1 = ""
            cKATEGORIE2 = ""
            cMwst = ""
            
            cLiBesNr = ""
            cLEKPR = ""
            cSEK = ""
            cRKZ = ""
        
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBez = rsrs!BEZEICH
                cBez = SwapStr(cBez, ";", " ")
            End If
            
            If Not IsNull(rsrs!nettopreis) Then
                cNettopreis = rsrs!nettopreis
                cNettopreis = Format(cNettopreis, "#####0.00")
            End If
            
            If Not IsNull(rsrs!lekpr) Then
                cLEKPR = rsrs!lekpr
                cLEKPR = Format(cLEKPR, "#####0.00")
            End If
            
            If Not IsNull(rsrs!sEK) Then
                cSEK = rsrs!sEK
                cSEK = Format(cSEK, "#####0.00")
            End If
            
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = rsrs!LIBESNR
            End If
            
            If Not IsNull(rsrs!geraeumt) Then
                cRKZ = rsrs!geraeumt
            End If
            
            If dAufAbschlag <> 0 Then
                If Not IsNull(rsrs!Rabattpreis) Then
                    cPreis = Format(rsrs!Rabattpreis, "#####0.00")
                End If
            Else
                If Not IsNull(rsrs!bruttopreis) Then
                    cPreis = Format(rsrs!bruttopreis, "#####0.00")
                End If
            End If
            
            If Not IsNull(rsrs!MARKE) Then
                cMarke = rsrs!MARKE
                cMarke = SwapStr(cMarke, ";", " ")
            End If
            
            If Not IsNull(rsrs!linbez) Then
                cLinbez = rsrs!linbez
                cLinbez = SwapStr(cLinbez, ";", " ")
            End If
            
            If Not IsNull(rsrs!BESTAND) Then
                cBestand = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!EAN) Then
                cEAN1 = rsrs!EAN
            End If
            
            If Not IsNull(rsrs!EAN2) Then
                cEAN2 = rsrs!EAN2
            End If
            
            If Not IsNull(rsrs!EAN3) Then
                cEAN3 = rsrs!EAN3
            End If
            
            If FileExists(sPfad & "\" & cArtNr & ".jpg") Then
                cBildangabe = cArtNr & ".jpg"
                
                sQuelle = sQuellpfad & "\" & cBildangabe
                sZiel = sZielpfad & "\" & cBildangabe

                lRet = CopyFile(sQuelle, sZiel, lfail)
            Else
                If FileExists(sPfad & "\" & "keinBild.jpg") Then
                    cBildangabe = "keinBild.jpg"
                    
                    sQuelle = sQuellpfad & "\" & cBildangabe
                    sZiel = sZielpfad & "\" & cBildangabe
    
                    lRet = CopyFile(sQuelle, sZiel, lfail)
                End If
            End If
            
            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            End If
            
            If Not IsNull(rsrs!AGNBEZ) Then
                cAGNBEZ = rsrs!AGNBEZ
                cAGNBEZ = SwapStr(cAGNBEZ, ";", " ")
            End If
            
            If Not IsNull(rsrs!PGN) Then
                cPGN = rsrs!PGN
            End If
            
            If Not IsNull(rsrs!PGNBEZ) Then
                cPGNBEZ = rsrs!PGNBEZ
                cPGNBEZ = SwapStr(cPGNBEZ, ";", " ")
            End If
            
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            End If
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cLINRBEZ = rsrs!LIEFBEZ
                cLINRBEZ = SwapStr(cLINRBEZ, ";", " ")
            End If
            
            If Not IsNull(rsrs!INHALT) Then
                cInhalt = rsrs!INHALT
            End If
            
            If Not IsNull(rsrs!INHALTBEZ) Then
                cInhaltBez = rsrs!INHALTBEZ
            End If
            
            If Not IsNull(rsrs!GRUNDPREIS) Then
                cGRUNDPREIS = rsrs!GRUNDPREIS
            End If
            
            
            cGI = ""
            cGP = ""
            BerechneGrundPreis CDbl(cInhalt), cInhaltBez, CDbl(cPreis), cGrundInhalt, dGrundPreisDM, dGrundPreisEur
            If dGrundPreisDM > 0 Then
                cGI = cGrundInhalt
                cGP = dGrundPreisEur
                cGP = Format(cGP, "#####0.00")
            End If
                 
            If Not IsNull(rsrs!ARTBEZ) Then
                cArtBez = rsrs!ARTBEZ
                cArtBez = SwapStr(cArtBez, ";", " ")
            End If
            
            If Not IsNull(rsrs!INTERBEZ) Then
                cINTERBEZ = rsrs!INTERBEZ
                cINTERBEZ = SwapStr(cINTERBEZ, ";", " ")
            End If
            
            If Not IsNull(rsrs!BESCHREIB) Then
                cBESCHREIB = rsrs!BESCHREIB
                cBESCHREIB = SwapStr(cBESCHREIB, ";", " ")
                cBESCHREIB = SwapStr(cBESCHREIB, Chr(10), " ")
                cBESCHREIB = SwapStr(cBESCHREIB, Chr(13), " ")
                
'                MsgBox cBESCHREIB
            End If
            
            If Not IsNull(rsrs!SHOPKVK) Then
                cSHOPKVK = rsrs!SHOPKVK
                cSHOPKVK = Format(cSHOPKVK, "#####0.00")
            End If
            
            If Not IsNull(rsrs!KATEGORIE1) Then
                cKATEGORIE1 = rsrs!KATEGORIE1
            End If
            
            If Not IsNull(rsrs!KATEGORIE2) Then
                cKATEGORIE2 = rsrs!KATEGORIE2
            End If
            
            If Not IsNull(rsrs!MWST) Then
                cMwst = rsrs!MWST
            End If
            
            

            
            cSatz = cArtNr & ";" & cBez & ";" & cNettopreis & ";" & cPreis & ";" & cMarke & ";" & cLinbez & ";" & cBestand & ";" & cEAN1 & ";" & cEAN2 & ";" & cEAN3 & ";" & cBildangabe
            cSatz = cSatz & ";" & cLiBesNr & ";" & cRKZ & ";" & cLEKPR & ";" & cSEK & ""
            cSatz = cSatz & ";" & cAgn & ";" & cAGNBEZ & ";" & cPGN & ";" & cPGNBEZ & ";" & cLinr & ";" & cLINRBEZ
            cSatz = cSatz & ";" & cInhalt & ";" & cInhaltBez & ";" & cGRUNDPREIS & ";" & cGI & ";" & cGP
            cSatz = cSatz & ";" & cArtBez & ";" & cINTERBEZ & ";" & cBESCHREIB & ";" & cSHOPKVK & ";" & cKATEGORIE1 & ";" & cKATEGORIE2 & ";" & cMwst
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    anzeige "normal", "Fertig! " & lAnz & " Artikel wurden bereitgestellt", Label1(4)
    
    
    
    
    iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
    
        picprogress.Visible = True
        Zip_Folder "", cPfad1 & "\BOX\Bilder", cPfad1 & "\BOX\artpic.zip", txtStatus
        picprogress.Visible = False
    
        gcBestellEmail.Attachment1 = cPfad1 & "BOX\products.csv"
        gcBestellEmail.Attachment2 = cPfad1 & "BOX\artpic.zip"
        Screen.MousePointer = 0
        frmWKL129.Show 1
        
    Else
    
        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: products.csv abgespeichert", vbInformation, "Winkiss Information:"
    End If

    Screen.MousePointer = 0
    

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "alleShopArtikel"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
         
    End If
End Sub
Private Sub ExcelExport()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim i           As Integer
    Dim cDatname    As String
    
    cDatname = "Artikel" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    If NewTableSuchenDBKombi("TOP" & srechnertab, gdBase) Then
    
        loeschNEW "ArtExc", gdBase
        
        gsZSpalte = ""
        gstab = "ARTEX"
        frmWKL36.Show 1
        
        'dannach Tablay auswerten
        Tabcheck "ARTEX"
        FormatGridOverTablay "ARTEX"
        
        If byAnzahlSpalten > 0 Then
            sSQL = "Select " & sSpaltenbez(0) & " "
            
            If byAnzahlSpalten > 1 Then
                For i = 1 To byAnzahlSpalten - 1
                    sSQL = sSQL & " , " & sSpaltenbez(i) & " "
                Next i
            End If
        Else
            Exit Sub
        End If
        
        sSQL = sSQL & " into ArtExc from TOP" & srechnertab
        gdBase.Execute sSQL, dbFailOnError
    
        cdatei = cPfad1 & "BOX\" & cDatname
        cPfad = cPfad1 & "BOX"
        
        sSQL = "Select * into ArtExc IN '" & cdatei & "' 'Excel 8.0;' from ArtExc "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    
        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
        loeschNEW "ArtExc", gdBase
    End If

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
  
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExcelExport"
        Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    If NewTableSuchenDBKombi("List3", gdBase) Then
        LadeList3
        Command5(4).Visible = True
        Label4(2).Visible = True
        Label4(2).Caption = List3.ListCount & " Lieferanten"
        Label4(2).Refresh
    End If

    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LadeList3()
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    List3.Clear
    List3.Visible = True
    
    cSQL = "Select * from LIST3 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!AuswahlTEXT) Then
                List3.AddItem rsrs!AuswahlTEXT
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeList3"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    Select Case Index
        Case 0, 3
            cValid = "-1234567890" & Chr$(8)
        Case 1, 4
            cValid = "1234567890," & Chr$(8)
        Case 2, 5
            cValid = "1234567890" & Chr$(8)
    End Select
    
    If InStr(cValid, UCase$(Chr$(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
    
    End If
    
    If KeyCode = vbKeyEscape Then
        Command5_Click 0
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 5
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
'                    Label2(0).Caption = gF2Prompt.cWert
                End If
            Case Is = 7
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
'                    Label2(0).Caption = gF2Prompt.cWert
                End If
            Case Is = 2
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
'                    Label2(9).Caption = gF2Prompt.cWert
                End If
        End Select
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    If List3.ListCount > 0 Then
        iRet = MsgBox("Möchten Sie die ausgewählten Lieferanten für die nächste Ermittlung abspeichern", vbYesNo + vbQuestion, "Winkiss Frage:")
        If iRet = vbYes Then
            SpeicherList3
        End If
    End If
    
    loeschNEW "ArtExcS", gdBase
    loeschNEW "ArtExc", gdBase
    loeschNEW "ArtExcA", gdBase
    loeschNEW "ART55C", gdBase
    
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
Private Sub SpeicherList3()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim i           As Integer
    
    loeschNEW "LIST3", gdBase
    
    sSQL = "Create Table List3 (AuswahlTEXT Text(100)) "
    gdBase.Execute sSQL, dbFailOnError

    cLBSatz = ""
    For i = 0 To List3.ListCount - 1
        cLBSatz = List3.list(i)
        sSQL = "Insert into List3 (AuswahlTEXT) "
        sSQL = sSQL & " Values (  '" & cLBSatz & "')"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherList3"
    Fehler.gsFehlertext = "Im Programmteil Artikelexport ist ein Fehler aufgetreten."
    
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
