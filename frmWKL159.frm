VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL159 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Rewe Warengruppen"
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
   Begin VB.CheckBox Check3 
      Caption         =   "Auslistung"
      Height          =   255
      Left            =   9480
      TabIndex        =   43
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Neulistung"
      Height          =   255
      Left            =   9480
      TabIndex        =   42
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "VK/EK/ALL Änderung"
      Height          =   255
      Left            =   9480
      TabIndex        =   40
      Top             =   4320
      Width           =   2175
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
      Height          =   345
      Index           =   2
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   27
      Top             =   6000
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Lieferantenkürzel vor die Artikelbezeichnung stellen"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   5640
      Value           =   1  'Aktiviert
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   23
      Top             =   4800
      Width           =   975
   End
   Begin sevCommand3.Command Command4 
      Height          =   355
      Index           =   12
      Left            =   6720
      TabIndex        =   20
      Top             =   3720
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin sevCommand3.Command Command4 
      Height          =   355
      Index           =   10
      Left            =   6720
      TabIndex        =   16
      Top             =   2640
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   9
      Left            =   8400
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
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
      Caption         =   "S auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   8
      Left            =   6960
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
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
      Caption         =   "L auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   7
      Left            =   8400
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
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
      Caption         =   "S entfernen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
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
      Caption         =   "L entfernen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   355
      Index           =   16
      Left            =   6720
      TabIndex        =   9
      Top             =   4800
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   5
      Left            =   5520
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
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
      Caption         =   "alle auswählen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
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
      Caption         =   "alle entfernen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   2
      Left            =   9840
      TabIndex        =   6
      Top             =   7200
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
      Caption         =   "Protokolle"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   1
      Left            =   9840
      TabIndex        =   5
      Top             =   7800
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
      Caption         =   "zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   3
      Top             =   240
      Width           =   345
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
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   3
      Left            =   7920
      TabIndex        =   0
      Top             =   7200
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
      Caption         =   "Weiter"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   18
      FixedCols       =   2
      ForeColorSel    =   8454143
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   0
      Left            =   10800
      TabIndex        =   44
      Top             =   240
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
      Picture         =   "frmWKL159.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Nur Artikel mit KZ"
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
      Index           =   18
      Left            =   9480
      TabIndex        =   41
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   17
      Left            =   6840
      TabIndex        =   39
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   16
      Left            =   6840
      TabIndex        =   38
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   37
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   14
      Left            =   1200
      TabIndex        =   36
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   13
      Left            =   4200
      TabIndex        =   35
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Anzahl"
      Height          =   255
      Index           =   12
      Left            =   4200
      TabIndex        =   34
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ausg. gesamt:"
      Height          =   255
      Index           =   11
      Left            =   5520
      TabIndex        =   33
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Artikel gesamt:"
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   32
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "ausg. Lager:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "ausg. Strecke:"
      Height          =   255
      Index           =   8
      Left            =   2880
      TabIndex        =   30
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "insg. Lager:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   29
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "insg. Strecke:"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   28
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "AGN"
      Height          =   375
      Index           =   5
      Left            =   5520
      TabIndex        =   25
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "AGN"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   24
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Edeka Lager"
      Height          =   495
      Index           =   3
      Left            =   5520
      TabIndex        =   22
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Lieferant Lager"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Edeka Strecke"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   18
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Lieferant Strecke"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Welche Artikel möchten Sie übernehmen?"
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
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   9015
   End
   Begin VB.Label lblanzeige 
      BackColor       =   &H00C0C000&
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
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Edeka Stammdaten einlesen"
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
      Width           =   9615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL159"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerWGN As Byte
Dim SpaltennummerAUSGEWAEHLT_L As Byte
Dim SpaltennummerAUSGEWAEHLT_S As Byte
Dim SpaltennummerANZAHL_L As Byte
Dim SpaltennummerANZAHL_S As Byte
Dim SpaltennummerWGNTEXT As Byte

Private Sub flex(krit As String, sAuswahl_art As String)
On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
    
        If sAuswahl_art = "L" Or sAuswahl_art = "" Then
            MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_L
            MSFlexGrid1.Row = lcount
    
            Select Case krit
                Case "ausgewählt"
                    MSFlexGrid1.Text = "ausgewählt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbGreen
                    
                Case "entfernt"
                    MSFlexGrid1.Text = "entfernt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbRed
            End Select
        End If
        
        If sAuswahl_art = "S" Or sAuswahl_art = "" Then
            MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_S
            MSFlexGrid1.Row = lcount
    
            Select Case krit
                Case "ausgewählt"
                    MSFlexGrid1.Text = "ausgewählt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbGreen
                    
                Case "entfernt"
                    MSFlexGrid1.Text = "entfernt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbRed
            End Select
        End If
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
    zeige_Anzahlen
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "flex"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub zeige_Anzahlen()
On Error GoTo LOKAL_ERROR
    
    
    Label2(13).Caption = zeige_Anzahl_ausgewählt("S")
    Label2(15).Caption = zeige_Anzahl_ausgewählt("L")
    
    Label2(12).Caption = zeige_Anzahl_alle("S")
    Label2(14).Caption = zeige_Anzahl_alle("L")
    
    Label2(16).Caption = CLng(Label2(14).Caption) + CLng(Label2(12).Caption)
    Label2(17).Caption = CLng(Label2(15).Caption) + CLng(Label2(13).Caption)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Anzahlen"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function zeige_Anzahl_ausgewählt(sAuswahl_art As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim lAnzahl         As Long
    
    lAnzahl = 0
    zeige_Anzahl_ausgewählt = "0"
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
    
        If sAuswahl_art = "L" Then
            MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_L
            MSFlexGrid1.Row = lcount
            
            If MSFlexGrid1.Text = "ausgewählt" Then
                MSFlexGrid1.Col = SpaltennummerANZAHL_L
                lAnzahl = lAnzahl + Val(MSFlexGrid1.Text)
            End If
        ElseIf sAuswahl_art = "S" Then
            MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_S
            MSFlexGrid1.Row = lcount
            
            If MSFlexGrid1.Text = "ausgewählt" Then
                MSFlexGrid1.Col = SpaltennummerANZAHL_S
                lAnzahl = lAnzahl + Val(MSFlexGrid1.Text)
            End If
        End If

    Next lcount
    
    zeige_Anzahl_ausgewählt = lAnzahl
    
    MSFlexGrid1.Redraw = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Anzahl_ausgewählt"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function zeige_Anzahl_alle(sAuswahl_art As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim lAnzahl         As Long
    
    lAnzahl = 0
    zeige_Anzahl_alle = "0"
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
    
        MSFlexGrid1.Row = lcount
    
        If sAuswahl_art = "L" Then
            MSFlexGrid1.Col = SpaltennummerANZAHL_L
            lAnzahl = lAnzahl + Val(MSFlexGrid1.Text)
        ElseIf sAuswahl_art = "S" Then
            MSFlexGrid1.Col = SpaltennummerANZAHL_S
            lAnzahl = lAnzahl + Val(MSFlexGrid1.Text)
        End If

    Next lcount
    
    zeige_Anzahl_alle = lAnzahl
    
    MSFlexGrid1.Redraw = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Anzahl_alle"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub speicherDieWahl()
On Error GoTo LOKAL_ERROR
    
    Dim lcount              As Long
    Dim sWgnText            As String
    Dim sSQL                As String
    Dim lWGN                As Long
    Dim sAUSGEWAEHLT_L      As String
    Dim sAUSGEWAEHLT_S      As String
    Dim bo1                 As Integer
    Dim bo2                 As Integer
    Dim bo3                 As Integer
    Dim bo4                 As Integer
    
    anzeige "normal", "", lblanzeige
    Screen.MousePointer = 11
    
    loeschNEW "EDEK_WGN_HIST", gdBase
    CreateTableT2 "EDEK_WGN_HIST", gdBase
    
    
    sSQL = "Insert into EDEK_WGN_HIST select "
    sSQL = sSQL & " WGN "
    sSQL = sSQL & " ,WGNTEXT "
    sSQL = sSQL & " ,AUSGEWAEHLT_L "
    sSQL = sSQL & " ,AUSGEWAEHLT_S "
    sSQL = sSQL & " from EDEK_WGN "
    gdBase.Execute sSQL, dbFailOnError
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
    
        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = SpaltennummerWGN
        lWGN = Val(MSFlexGrid1.Text)
        
        MSFlexGrid1.Col = SpaltennummerWGNTEXT
        sWgnText = MSFlexGrid1.Text
        
        MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_L
        sAUSGEWAEHLT_L = MSFlexGrid1.Text
        
        MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_S
        sAUSGEWAEHLT_S = MSFlexGrid1.Text
        
        sSQL = "Update EDEK_WGN_HIST set WGNTEXT = '" & sWgnText & "'"
        sSQL = sSQL & " , AUSGEWAEHLT_L  = '" & sAUSGEWAEHLT_L & "'"
        sSQL = sSQL & " , AUSGEWAEHLT_S  = '" & sAUSGEWAEHLT_S & "'"
        sSQL = sSQL & " where WGN = " & lWGN
        gdBase.Execute sSQL, dbFailOnError
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
    If Check1.Value = vbChecked Then
        bo1 = -1
    Else
        bo1 = 0
    End If
    
    If Check2.Value = vbChecked Then
        bo2 = -1
    Else
        bo2 = 0
    End If
    
    If Check3.Value = vbChecked Then
        bo3 = -1
    Else
        bo3 = 0
    End If
    
    If Check4.Value = vbChecked Then
        bo4 = -1
    Else
        bo4 = 0
    End If
    
    sSQL = "Update EDEKE set LINR_L = " & Text1(1).Text & ""
    sSQL = sSQL & " , LINR_S  = " & Text1(0).Text & ""
    sSQL = sSQL & " , AGN  = " & Text1(4).Text & ""
    sSQL = sSQL & " , KUERZEL  = '" & Text1(2).Text & "'"
    sSQL = sSQL & " , BO1  = " & bo1 & ""
    sSQL = sSQL & " , BO2  = " & bo2 & ""
    sSQL = sSQL & " , BO3  = " & bo3 & ""
    sSQL = sSQL & " , BO4  = " & bo4 & ""
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    Unload frmWKL159
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDieWahl"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next

End Sub
Private Sub ZeigeEdeka_Voreinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim rsrs As Recordset
    Dim sSQL As String
    
    Text1(1).Text = ""
    Text1(0).Text = ""
    Text1(4).Text = ""
    Text1(2).Text = ""
    sSQL = "Select * from EDEKE"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LINR_L) Then
            Text1(1).Text = rsrs!LINR_L
        End If
        
        If Not IsNull(rsrs!LINR_S) Then
            Text1(0).Text = rsrs!LINR_S
        End If
        
        If Not IsNull(rsrs!AGN) Then
            Text1(4).Text = rsrs!AGN
        End If
        
        If Not IsNull(rsrs!Kuerzel) Then
            Text1(2).Text = rsrs!Kuerzel
        End If
        
        If rsrs!bo1 = True Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If
        
        If rsrs!bo2 = True Then
            Check2.Value = vbChecked
        Else
            Check2.Value = vbUnchecked
        End If
        
        If rsrs!bo3 = True Then
            Check3.Value = vbChecked
        Else
            Check3.Value = vbUnchecked
        End If
        
        If rsrs!bo4 = True Then
            Check4.Value = vbChecked
        Else
            Check4.Value = vbUnchecked
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeEdeka_Voreinstellungen"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        Case 0
            gsZSpalte = "WGN"
            gstab = "EDEKAGRUPPE"
            frmWKL36.Show 1
            'fertig
            
            ZeigeEdekagruppe
            If MSFlexGrid1.Visible = True Then
                MSFlexGrid1.Col = 1
                MSFlexGrid1.Row = 2
                MSFlexGrid1.SetFocus
            End If
        Case 1 'Zurück
            gb159 = False
            Unload frmWKL159
        Case 3
            speicherDieWahl
        Case 4
            flex "entfernt", ""
        Case 5
            flex "ausgewählt", ""
        Case 6
            flex "entfernt", "L"
        Case 7
            flex "entfernt", "S"
        Case 8
            flex "ausgewählt", "L"
        Case 9
            flex "ausgewählt", "S"
        Case 10
            Text1_KeyUp 0, vbKeyF2, 0
        Case 11
            gsHelpstring = "Edeka Warengruppen"
            frmWKL110.Show 1
        Case 12
            Text1_KeyUp 1, vbKeyF2, 0
        Case 16
            Text1_KeyUp 4, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    anzeige "normal", "", lblanzeige
    
'    ZeigeEdekagruppe
    
    If MSFlexGrid1.Visible = True Then
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = 2
        MSFlexGrid1.SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
         Case Is = SpaltennummerWGNTEXT
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß#"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                    
                End If
            End If
     End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_LeaveCell()
    On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    
    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSFlexGrid1.Rows Then
        lrow = lrow - 1
    End If
    
    If KeyCode = &H28 Or KeyCode = &H27 Or KeyCode = &H26 Or KeyCode = &H25 Or KeyCode = vbKeyF2 Then
        Exit Sub
    End If
    
    If iKeypress = 0 And KeyCode <> vbKeyBack Then
        
        If KeyCode <> 46 Then
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Text = ""
        End If
    End If
    iKeypress = iKeypress + 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "WGN"
                SpaltennummerWGN = i
            Case Is = "WGNTEXT"
                SpaltennummerWGNTEXT = i
            Case Is = "AUSGEWAEHLT_S"
                SpaltennummerAUSGEWAEHLT_S = i
            Case Is = "AUSGEWAEHLT_L"
                SpaltennummerAUSGEWAEHLT_L = i
            Case Is = "ANZ_S"
                SpaltennummerANZAHL_S = i
            Case Is = "ANZ_L"
                SpaltennummerANZAHL_L = i
        End Select
    Next i
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex159()
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
    Dim cSQL        As String
   
    cSQL = "Select * from EDEK_WGN order by WGN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
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
                            
                            Case Is = "auswählen L", "auswählen S"
                                
                                .Row = lrow
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                        sWert = rsrs(sSpaltenbez(i))
                                    Else
                                        sWert = ""
                                    End If
                                    .Row = lrow
                                    .Text = sWert
                                
                                    If sWert = "ausgewählt" Then
                                        .CellFontBold = True
                                        .CellForeColor = vbGreen
                                    Else
                                        .CellFontBold = True
                                        .CellForeColor = vbRed
                                    End If
                                End If
                                
                                
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        
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
    Fehler.gsFunktion = "FuellenMSFlex159"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeEdekagruppe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    If NewTableSuchenDBKombi("EDEK_WGN_HIST", gdBase) Then
        sSQL = "Update EDEK_WGN e inner join EDEK_WGN_HIST h on e.WGN = h.WGN "
        sSQL = sSQL & " Set  e.WGNTEXT = h.WGNTEXT "
        sSQL = sSQL & " ,e.AUSGEWAEHLT_L = h.AUSGEWAEHLT_L "
        sSQL = sSQL & " ,e.AUSGEWAEHLT_S = h.AUSGEWAEHLT_S "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    ZeigeEdekaTAB
    
    ZeigeEdeka_Voreinstellungen
    
    zeige_Anzahlen
    
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeEdekagruppe"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeEdekaTAB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    Set recAnz = gdBase.OpenRecordset("EDEK_WGN")
    If recAnz.EOF Then
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        
        anzeige "rot", "Keine Warengruppen gefunden!", lblanzeige
        recAnz.Close
        Exit Sub
    End If
    recAnz.Close
    
    Screen.MousePointer = 11

    Tabcheck "EDEKAGRUPPE"
    
    FormatGridOverTablay "EDEKAGRUPPE"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex159
        ermittlespalten
        
        .Redraw = False
    
        Tabellenbreiteanpassen MSFlexGrid1, 1 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeEdekaTAB"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
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
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR

    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    Else
        If MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_L Then
    
            Select Case MSFlexGrid1.Text()
                Case "entfernt"
                    MSFlexGrid1.Text = "ausgewählt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbGreen
                Case "ausgewählt"
                    MSFlexGrid1.Text = "entfernt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbRed
            End Select
            
            zeige_Anzahlen
        ElseIf MSFlexGrid1.Col = SpaltennummerAUSGEWAEHLT_S Then
    
            Select Case MSFlexGrid1.Text()
                Case "entfernt"
                    MSFlexGrid1.Text = "ausgewählt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbGreen
                Case "ausgewählt"
                    MSFlexGrid1.Text = "entfernt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbRed
            End Select
            
            zeige_Anzahlen
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index

    Case 0
        Label2(1).Caption = ermLiefBez(CLng(Text1(0).Text))
    Case 1
        Label2(3).Caption = ermLiefBez(CLng(Text1(1).Text))
    Case 4
        Label2(5).Caption = Ermittleagntext(Text1(4).Text)
    
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            
            Case 4   'AGN
                gF2Prompt.cFeld = "AGN"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
            Case 0
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label2(1).Caption = gF2Prompt.cWert
                    
                    Text1(2).Text = UCase(Trim(Left(gF2Prompt.cWert, 3)))
                    
                End If
            Case 1
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label2(3).Caption = gF2Prompt.cWert
                    
                    Text1(2).Text = UCase(Trim(Left(gF2Prompt.cWert, 3)))
                    
                End If
        End Select
        
    End If
    
    Text1(Index).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Edeka Warengruppen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
