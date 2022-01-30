VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL139 
   Caption         =   "Mindestbestand"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   3495
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
      Caption         =   "Mindestbestände festsetzen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   8640
      MaxLength       =   6
      TabIndex        =   26
      Top             =   6360
      Width           =   855
   End
   Begin sevCommand3.Command Command17 
      Height          =   495
      Left            =   7920
      TabIndex        =   25
      Top             =   6720
      Width           =   3735
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
      Caption         =   "Mindestbestände nur für diesen Artikel sofort rechnen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   3495
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
      Caption         =   "Mindestbestand überschritten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command13 
      Height          =   495
      Left            =   7920
      TabIndex        =   22
      Top             =   5520
      Width           =   3735
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
      Caption         =   "Mindestbestände sofort rechnen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7920
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
      Begin sevCommand3.Command Command12 
         Height          =   375
         Left            =   7920
         TabIndex        =   16
         Top             =   5520
         Width           =   1335
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
         Caption         =   "Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Text            =   "1,5"
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "4 Monate"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "5 Monate"
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "6 Monate"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "7 Monate"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "8 Monate"
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
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "9 Monate"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   36
         ToolTipText     =   "Kalender"
         Top             =   2400
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
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   1
         Left            =   1680
         TabIndex        =   37
         ToolTipText     =   "Kalender"
         Top             =   2760
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
      Begin VB.Label lbl6 
         Caption         =   "Berechnungszeitraum"
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
         Index           =   49
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lbl6 
         Caption         =   "von"
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
         Index           =   50
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lbl6 
         Caption         =   "bis"
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
         Index           =   51
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lbl6 
         Caption         =   "Bevorratung/ Faktor"
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
         Index           =   52
         Left            =   2160
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl6 
         Caption         =   "VK Zeitraum ausschließen"
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
         Index           =   53
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1695
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   3495
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
      Caption         =   "Mindestbestand unterschritten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   4
      Top             =   360
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   7920
      TabIndex        =   3
      Top             =   4920
      Width           =   3735
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
      Caption         =   "Voreinstellungen speichern"
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
   Begin VB.Label lbl6 
      Caption         =   "Anzeige"
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
      Index           =   7
      Left            =   6000
      TabIndex        =   34
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lbl6 
      Caption         =   "Gesamtschnitteinkaufswert (Mindestbestand überschritten):"
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
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Label lbl6 
      Caption         =   "Anzeige"
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
      Index           =   5
      Left            =   6000
      TabIndex        =   32
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lbl6 
      Caption         =   "Mindestbestand unterschritten (Artikelanzahl):"
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
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label lbl6 
      Caption         =   "Anzeige"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   30
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbl6 
      Caption         =   "Mindestbestand überschritten (Artikelanzahl):"
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
      TabIndex        =   29
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lbl6 
      Caption         =   "Anzeige"
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
      TabIndex        =   28
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lbl6 
      Caption         =   "Artnr"
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
      Index           =   69
      Left            =   7920
      TabIndex        =   27
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lbl6 
      Caption         =   "Voreinstellungen"
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
      Left            =   7920
      TabIndex        =   24
      Top             =   1080
      Width           =   2775
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
      Caption         =   "Mindestbestand"
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
      Width           =   6255
   End
End
Attribute VB_Name = "frmWKL139"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command13_Click()
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String

    Dim lVon    As Long
    Dim lBis    As Long
    
    Dim lDiff1  As Long
    Dim lDiff2  As Long
    Dim lDif    As Long
    
    Dim iTage As Integer
    
    Select Case MBDETAILMON
        Case 5 '9
            iTage = 272
        Case 4 '8
            iTage = 241
        Case 3 '7
            iTage = 211
        Case 2 '6
            iTage = 180
        Case 1 '5
            iTage = 150
        Case 0 '4
            iTage = 119
        Case Else
            iTage = 180
    End Select
    
    lVon = DateValue(Now) - iTage
    lBis = DateValue(Now)
    
    lDiff1 = lBis - lVon
    lDiff2 = MBDETAILBIS - MBDETAILVON
    
    If MBDETAILVON <= lBis And MBDETAILBIS <= lBis And MBDETAILVON >= lVon And MBDETAILBIS >= lVon Then
        'Fall1
'        MsgBox "Fall 1 optimal"
        lDif = lDiff1 - lDiff2
        
        
    ElseIf MBDETAILVON <= lBis And MBDETAILBIS > lBis And MBDETAILVON >= lVon Then
'        MsgBox "Fall 2 Überschneidungen Zeitraum teilweise größer"
        'Fall2
        lDif = MBDETAILVON - lVon
    ElseIf MBDETAILBIS <= lBis And MBDETAILVON < lVon And MBDETAILBIS >= lVon Then
'        MsgBox "Fall 3 Überschneidungen Zeitraum teilweise kleiner"
        'Fall3
        lDif = lBis - MBDETAILBIS
    ElseIf MBDETAILVON < lVon And MBDETAILBIS < lVon Then
'        MsgBox "Fall 4 Zeitraum komplett kleiner"
        'Fall4
        lDif = lDiff1
        
    ElseIf MBDETAILVON > lBis And MBDETAILBIS > lBis Then
'        MsgBox "Fall 5 Zeitraum komplett größer"
        'Fall5
        lDif = lDiff1
    End If
    
'    MsgBox lDif
    
    MBrechnen1 MBDETAILBVO, CInt(lDif), 1, lVon, lBis, Label1(4), MBDETAILVON, MBDETAILBIS
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command17_Click()
On Error GoTo LOKAL_ERROR

    Dim lVon    As Long
    Dim lBis    As Long
    
    Dim lDiff1  As Long
    Dim lDiff2  As Long
    Dim lDif    As Long
    
    Dim iTage As Integer
    
    Select Case MBDETAILMON
        Case 5 '9
            iTage = 272
        Case 4 '8
            iTage = 241
        Case 3 '7
            iTage = 211
        Case 2 '6
            iTage = 180
        Case 1 '5
            iTage = 150
        Case 0 '4
            iTage = 119
        Case Else
            iTage = 180
    End Select
    
    lVon = DateValue(Now) - iTage
    lBis = DateValue(Now)
    
    lDiff1 = lBis - lVon
    lDiff2 = MBDETAILBIS - MBDETAILVON
    
    If MBDETAILVON <= lBis And MBDETAILBIS <= lBis And MBDETAILVON >= lVon And MBDETAILBIS >= lVon Then
        'Fall1
'        MsgBox "Fall 1 optimal"
        lDif = lDiff1 - lDiff2
        
        
    ElseIf MBDETAILVON <= lBis And MBDETAILBIS > lBis And MBDETAILVON >= lVon Then
'        MsgBox "Fall 2 Überschneidungen Zeitraum teilweise größer"
        'Fall2
        lDif = MBDETAILVON - lVon
    ElseIf MBDETAILBIS <= lBis And MBDETAILVON < lVon And MBDETAILBIS >= lVon Then
'        MsgBox "Fall 3 Überschneidungen Zeitraum teilweise kleiner"
        'Fall3
        lDif = lBis - MBDETAILBIS
    ElseIf MBDETAILVON < lVon And MBDETAILBIS < lVon Then
'        MsgBox "Fall 4 Zeitraum komplett kleiner"
        'Fall4
        lDif = lDiff1
        
    ElseIf MBDETAILVON > lBis And MBDETAILBIS > lBis Then
'        MsgBox "Fall 5 Zeitraum komplett größer"
        'Fall5
        lDif = lDiff1
    End If
    
    MBrechnen1proArtnr MBDETAILBVO, CInt(lDif), 1, lVon, lBis, Label1(4), MBDETAILVON, MBDETAILBIS, CLng(Text3(6).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command17_Click"
    Fehler.gsFehlertext = "Im Programmteil Programmeinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub MBrechnen1proArtnr(dBVO As Double, iVKTage As Integer, iReserv As Integer, lVon As Long, lBis As Long, lblanzeige As Label, lvonNot As Long, lbisNot As Long, glartnr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim sTab    As String
    Dim rs      As Recordset
    Dim lartnr  As Long
    Dim j       As Integer
    Dim dnewMB  As Double
    Dim lnewMB  As Long
    Dim iFil    As Integer
    Dim lmbnew  As Long
    Dim ctmp    As String
    Dim lVkMenge As Long
    
    Dim dTeiler As Double
    
    Dim ifilvon As Integer
    Dim ifilbis As Integer
    

    dTeiler = iVKTage / 30
    
    loeschNEW "DRUMBAE1", gdBase
    CreateTable "DRUMBAE1", gdBase

    
    loeschNEW "DRUMBAE", gdBase
    CreateTable "DRUMBAE", gdBase

    sSQL = " Insert into DRUMBAE select " & gcFilNr & " as filiale , artnr from Artikel where artnr = " & glartnr
    gdBase.Execute sSQL, dbFailOnError
    
    sTab = "KASS" & gcFilNr
    loeschNEW sTab, gdBase
    sSQL = "Select artnr, menge into " & sTab & " from kassjour where  "
    sSQL = sSQL & "  ADATE between " & Trim$(Str$(lVon)) & " and " & Trim$(Str$(lBis)) & " "
    sSQL = sSQL & " and not ADATE between " & Trim$(Str$(lvonNot)) & " and " & Trim$(Str$(lbisNot)) & " "
    sSQL = sSQL & " and artnr = " & glartnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index ARTNR on " & sTab & " (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Create Index MENGE on " & sTab & " (MENGE)"
    gdBase.Execute sSQL, dbFailOnError
    
    
    Set rs = gdBase.OpenRecordset("DRUMBAE")
    If Not rs.EOF Then
        rs.MoveLast
        
        rs.MoveFirst
        Do While Not rs.EOF
            
            If Not IsNull(rs!artnr) Then
                lartnr = rs!artnr
            End If
            
            lVkMenge = Ermittlevk(CLng(gcFilNr), CLng(rs!artnr), gdBase, sTab)
            dnewMB = 0
            
            dnewMB = (dBVO / dTeiler) * lVkMenge
            
            lnewMB = 0
            If dnewMB > 0 Then
            
                lnewMB = Val(dnewMB)
                lnewMB = lnewMB + 1
                
                ctmp = "Der neue (aufgerundete) Mindestbestand: " & lnewMB
                ctmp = ctmp & vbCrLf
            
                MsgBox ctmp, vbInformation, "Winkiss Hinweis:"
                ctmp = ""
                
            End If
    
            rs.Edit
            rs!VKMENGE = lVkMenge
            rs!NEWMB = lnewMB
            rs.Update
            
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    sSQL = "Insert into DRUMBAE1 Select * from Drumbae "
    gdBase.Execute sSQL, dbFailOnError
        
    
    'Übernahme
    
    anzeige "normal", "Löschen nicht relevanter Daten...", lblanzeige
    
    sSQL = "Delete from  DRUMBAE1 where newmb = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Mindestbestände auf 0 setzen...", lblanzeige
    
    sSQL = "update ARTIKEL set minbest = 0 where artnr = " & glartnr
    gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("drumbae1")
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!artnr) Then
                lartnr = rs!artnr
                
                If Not IsNull(rs!NEWMB) Then
                    lmbnew = rs!NEWMB
                Else
                    lmbnew = 0
                End If
                schreibeNeuMBausNacht lartnr, lmbnew
                
            End If
        
        rs.MoveNext
        Loop
    End If
    rs.Close
    
    Dim rsrs1 As Recordset
    
    sSQL = "Select * from MBORDER where artnr =  " & glartnr
    Set rsrs1 = gdBase.OpenRecordset(sSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!artnr) Then
                lartnr = rsrs1!artnr
            End If
            
            If Not IsNull(rsrs1!MB) Then
                lmbnew = rsrs1!MB
            End If
            
            schreibeNeuMBausNacht lartnr, lmbnew
        
        
        rsrs1.MoveNext
        Loop
    End If
    rsrs1.Close

    ctmp = ctmp & "Wenn dieser Artikel einen festgesetzten MB enthält, dann wird er mit diesem überschrieben. "
    MsgBox ctmp, vbInformation, "Winkiss Hinweis:"
    
    sSQL = "Update Artikel set Artikel.Bestand = 0 "
    sSQL = sSQL & " where Artikel.Bestand is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Fertig", lblanzeige
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul6"
    Fehler.gsFunktion = "MBrechnen1proArtnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
    
        Case 11
            gsHelpstring = "Mindestbestand"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestbestand ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim lDate As Long
    
Select Case Index
    Case 0
        Unload frmWKL139
    Case 1
        speicherMBDetails
        leseMBDetails
    Case 2
        minbestunter
    Case 3
        minbestueber
    Case 4
        frmWKL140.Show 1
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestbestand ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub minbestunter()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cART As String
    Dim i As Integer
    
    loeschNEW "MBUNTER", gdBase
    CreateTableT2 "MBUNTER", gdBase
    
    Screen.MousePointer = 11
    
    sSQL = "Insert into MBUNTER  "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.Artnr "
    sSQL = sSQL & " , artikel.Bestand "
    sSQL = sSQL & " , artikel.MINBEST "
    sSQL = sSQL & " , artikel.bezeich "
    sSQL = sSQL & " , artikel.kvkpr1 "
    sSQL = sSQL & " , artikel.libesnr "
    sSQL = sSQL & " , artikel.linr "
    sSQL = sSQL & " , lisrt.liefbez "
    sSQL = sSQL & " from artikel,lisrt"
    sSQL = sSQL & " where  "
    sSQL = sSQL & " Artikel.gefuehrt = 'J' and Artikel.LPZ <> 0"
    sSQL = sSQL & " and artikel.bestand < artikel.minbest "
    sSQL = sSQL & " and artikel.linr = lisrt.linr "
    sSQL = sSQL & " and artikel.Minbest > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update MBUNTER inner join MBORDER on MBUNTER.Artnr = MBORDER.Artnr "
    sSQL = sSQL & " set MBUNTER.BLOCK = 'B' "
    gdBase.Execute sSQL, dbFailOnError
    

    
    Screen.MousePointer = 0

    reportbildschirm "", "aWKL00k"
    
    Pause (2)
    loeschNEW "MBUNTER", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "minbestunter"
    Fehler.gsFehlertext = "Bei unterschrittender Mindestbestand ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub minbestueber()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cART        As String
    Dim i           As Integer
    
    loeschNEW "MBUNTER", gdBase
    CreateTableT2 "MBUNTER", gdBase
    
    Screen.MousePointer = 11
    
    sSQL = "Insert into MBUNTER  "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.Artnr "
    sSQL = sSQL & " , artikel.Bestand "
    sSQL = sSQL & " , artikel.MINBEST "
    sSQL = sSQL & " , artikel.bezeich "
    sSQL = sSQL & " , artikel.kvkpr1 "
    sSQL = sSQL & " , artikel.libesnr "
    sSQL = sSQL & " , artikel.linr "
    sSQL = sSQL & " , lisrt.liefbez "
    sSQL = sSQL & " from artikel,lisrt"
    sSQL = sSQL & " where "
    sSQL = sSQL & " Artikel.gefuehrt = 'J' and Artikel.LPZ <> 0"
    sSQL = sSQL & " and artikel.bestand > artikel.minbest "
    sSQL = sSQL & " and artikel.linr = lisrt.linr "
    sSQL = sSQL & " and artikel.Minbest > 0 "
    
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update MBUNTER inner join MBORDER on MBUNTER.Artnr = MBORDER.Artnr "
    sSQL = sSQL & " set MBUNTER.BLOCK = 'B' "
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    reportbildschirm "", "aWkl00l"
    Pause (2)
    loeschNEW "MBUNTER", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "minbestueber"
    Fehler.gsFehlertext = "Bei überschrittender Mindestbestand ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub speicherMBDetails()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim dWert       As Double
    Dim i           As Integer
    
    loeschNEW "MBDETAIL", gdBase
    CreateTableT2 "MBDETAIL", gdBase
    
    
    'Bvo
    If Text3(2).Text = "" Then
        dWert = 1.5
        sSQL = "Insert into MBDETAIL (BVO) values ('1,5')"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILBVO = dWert
    Else
        dWert = CDbl(Text3(2).Text)
    
        sSQL = "Insert into MBDETAIL (BVO) values ('" & Text3(2).Text & "')"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILBVO = dWert
    End If
    
    'von
    If Text3(0).Text = "" Then
        Text3(0).Text = "01.01.2000"
        sSQL = "update MBDETAIL set von = '" & DateValue(Text3(0).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILVON = DateValue(Text3(0).Text)
    Else
        sSQL = "update MBDETAIL set von = '" & DateValue(Text3(0).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILVON = DateValue(Text3(0).Text)
    End If
    
    'bis
    If Text3(1).Text = "" Then
        Text3(1).Text = "01.01.2000"
        sSQL = "update MBDETAIL set bis = '" & DateValue(Text3(1).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILBIS = DateValue(Text3(1).Text)
    Else
        sSQL = "update MBDETAIL set bis = '" & DateValue(Text3(1).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        MBDETAILBIS = DateValue(Text3(1).Text)
    End If
    
    For i = 0 To 5
        If Option4(i).Value = True Then
            sSQL = "update MBDETAIL set optmon = " & i
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherMBDetails"
    Fehler.gsFehlertext = "Im Programmteil Mindestbestand ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Text3(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
        Case Is = 1
            Text3(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text3(2).Text = "1,5"

    leseMBDetails
    
    Text3(2).Text = MBDETAILBVO
    Text3(0).Text = Format$(MBDETAILVON, "DD.MM.YY")
    Text3(1).Text = Format$(MBDETAILBIS, "DD.MM.YY")
    Option4(MBDETAILMON).Value = True
    
    lbl6(3).Caption = ermgesMB(">")
    lbl6(3).Refresh
    
    lbl6(5).Caption = ermgesMB("<")
    lbl6(5).Refresh
    
    lbl6(7).Caption = Format(ermgesMBSchnittEK(">"), "######0.00" & " Euro")
    lbl6(7).Refresh
    
    
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Mindestbestand ist ein Fehler aufgetreten."
    
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


