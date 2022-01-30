VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL172 
   Caption         =   "Mailing Feedback"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL172.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Caption         =   "mindestens 2x gekauft"
      Height          =   375
      Left            =   9600
      TabIndex        =   32
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   36.248
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   36.248
      TabIndex        =   21
      Top             =   2280
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
      Height          =   405
      Index           =   3
      Left            =   8040
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   2
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
      Caption         =   "Suche"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3720
      TabIndex        =   12
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vormonat"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "aktueller Monat"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gestern"
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
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Heute"
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
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "aktuelles Jahr"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum Voreinstellung"
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
         Index           =   12
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
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
      Height          =   405
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   960
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
      Height          =   405
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Tag             =   "2"
      Top             =   1440
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
      Height          =   405
      Index           =   2
      Left            =   8040
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   6
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
      Index           =   0
      Left            =   9600
      TabIndex        =   3
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
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   20
      Left            =   3120
      TabIndex        =   33
      ToolTipText     =   "Kalender"
      Top             =   960
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
      Index           =   21
      Left            =   3120
      TabIndex        =   34
      ToolTipText     =   "Kalender"
      Top             =   1440
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
      Height          =   165
      Index           =   1
      Left            =   2760
      TabIndex        =   35
      Top             =   1200
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
   Begin sevCommand3.Command Command0 
      Height          =   165
      Index           =   0
      Left            =   2760
      TabIndex        =   36
      Top             =   960
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
      Height          =   165
      Index           =   2
      Left            =   2760
      TabIndex        =   37
      Top             =   1440
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
      Height          =   165
      Index           =   3
      Left            =   2760
      TabIndex        =   38
      Top             =   1680
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
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   39
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Verkaufsstückzahl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7440
      TabIndex        =   31
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Nettoertrag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   7440
      TabIndex        =   30
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
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
      Left            =   6840
      TabIndex        =   29
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2760
      TabIndex        =   28
      Top             =   6720
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Nettoertrag in Euro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7440
      TabIndex        =   27
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Nettoertrag in Euro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   26
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   2760
      TabIndex        =   25
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
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
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label6 
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
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
      Index           =   0
      Left            =   480
      MouseIcon       =   "frmWKL172.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   22
      Top             =   4440
      Width           =   6615
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
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Preis:"
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
      Index           =   3
      Left            =   6720
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
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
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
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
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Artnr:"
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
      Index           =   2
      Left            =   6720
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Mailing Feedback"
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
      TabIndex        =   4
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmWKL172"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glfarbeR(18) As Long
Dim bErstemal As Boolean
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    Select Case Index
        Case 0
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                lDat = lDat + 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 1
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                lDat = lDat - 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 2
            If IsDate(Text1(1).Text) = False Then
                Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(1).Text) = True Then
                    lDat = CLng(DateValue(Text1(1).Text))
                End If
                lDat = lDat + 1
                Text1(1).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 3
            If IsDate(Text1(1).Text) = False Then
                Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(1).Text) = True Then
                    lDat = CLng(DateValue(Text1(1).Text))
                End If
                lDat = lDat - 1
                Text1(1).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 20         ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(1).SetFocus
        Case 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
        End Select
        
        Command5_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL172
        Case 1
            zeigeRabattkreis
        Case 2
            DruckeArtikelKumuliert
        Case 11
            gsHelpstring = "Mailing Feedback"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeArtikelKumuliert()
    On Error GoTo LOKAL_ERROR
    
'    Kuant ist die Basistabelle
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim lDatum As Long
    Dim lBelegnr As Long
    
    Screen.MousePointer = 11
    
    loeschNEW "ARTIKELSAMMLUNG", gdBase
    CreateTableT2 "ARTIKELSAMMLUNG", gdBase
    
    sSQL = "select * from KUANT"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            lDatum = 0
            lBelegnr = 0
            If Not IsNull(rsrs!Adate) Then
                lDatum = rsrs!Adate
            End If
            
            If Not IsNull(rsrs!BELEGNR) Then
                lBelegnr = rsrs!BELEGNR
            End If
            
            If lDatum > 0 And lBelegnr > 0 Then
                InsertArtikelintoSammlung lDatum, lBelegnr, CLng(Text1(2).Text)
            End If

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    reportbildschirm "WKL017", "aWKL172"
    
    

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeArtikelKumuliert"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub zeigeRabattkreis()
    On Error GoTo LOKAL_ERROR
    
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim lartnr As Long
    Dim dKVkPr1 As Double
    Dim i As Integer
    
    If bErstemal Then
        bErstemal = False
        Exit Sub
    End If
    
    anzeige "normal", "", Label1(4)
    
    If Text1(2).Text = "" Then
        anzeige "rot2", "Bitte Artikelnummer eingeben!", Label1(4)
        Text1(2).SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    For i = 1 To 20
        Unload Label5(i)
        Unload Label6(i)
    Next i
    
    Label5(0).Visible = False
    Label6(0).Visible = False
    
    Picture1.FillStyle = 1
    Picture1.Refresh
    
    lDatVon = -1
    lDatBis = -1
    
    lDatVon = DateValue(Text1(0).Text)
    lDatBis = DateValue(Text1(1).Text)
    
    
    lartnr = Val(Text1(2).Text)
    
    If Text1(3).Text <> "" Then
        If IsNumeric(Text1(3).Text) Then
            dKVkPr1 = CDbl(Text1(3).Text)
        Else
            Text1(3).Text = Format(ermKVKPR1(CStr(lartnr)), "######0.00")
            dKVkPr1 = CDbl(Text1(3).Text)
        End If
    Else
        Text1(3).Text = Format(ermKVKPR1(CStr(lartnr)), "######0.00")
        dKVkPr1 = CDbl(Text1(3).Text)
        
    End If

    If lDatVon > -1 And lDatBis > -1 Then
        KundenAnteil lDatVon, lDatBis, lartnr, dKVkPr1
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 340 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigeRabattkreis"
        Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Sub KundenAnteil(lDatVon As Long, lDatBis As Long, lartnr As Long, dKVK As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL                As String
    Dim rsrs                As Recordset
    Dim lsumKunden          As Long
    Dim lDatum              As Long
    Dim lBelegnr            As Long
    Dim dBonPreis           As Double
    Dim lKUANZOver          As Long
    Dim lKUANZUnder         As Long
    Dim lKUNDNR             As Long
    Dim dStartKreis         As Double
    Dim dEndKreis           As Double
    Dim dNettoertragproBon  As Double
    Dim dsumNettoertrag     As Double
    Dim lsumVKmenge         As Long
    
    Dim j                   As Double
    Dim siRadius            As Single
    
    Dim sTemp               As String
    
    loeschNEW "KUANT", gdBase
    CreateTableT2 "KUANT", gdBase
    
    loeschNEW "KUNDENSTAN", gdBase
    CreateTableT2 "KUNDENSTAN", gdBase
    
    loeschNEW "KUNDENPLUS", gdBase
    CreateTableT2 "KUNDENPLUS", gdBase
    
    cSQL = "Insert into KUANT select distinct adate, BELEGNR  "
    cSQL = cSQL & " from Kassjour where ADATE >= " & Trim$(Str$(lDatVon))
    cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lDatBis))
    cSQL = cSQL & " and ARTNR = " & lartnr
    
    If Check1.Value = vbChecked Then
        cSQL = cSQL & " and Menge > 1 "
    End If
    
    cSQL = cSQL & " group by adate,BELEGNR "
    gdBase.Execute cSQL, dbFailOnError
    
    lsumKunden = 0
    cSQL = "select count(*) as maxi from KUANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lsumKunden = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Verkaufsmenge vom Aktionartikel
    lsumVKmenge = 0
    cSQL = "select sum(Menge) as maxi "
    cSQL = cSQL & " from Kassjour where ADATE >= " & Trim$(Str$(lDatVon))
    cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lDatBis))
    cSQL = cSQL & " and ARTNR = " & lartnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lsumVKmenge = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lKUANZOver = 0
    lKUANZUnder = 0
    dsumNettoertrag = 0
    
    cSQL = "select * from KUANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            lDatum = 0
            lBelegnr = 0
            If Not IsNull(rsrs!Adate) Then
                lDatum = rsrs!Adate
            End If
            
            If Not IsNull(rsrs!BELEGNR) Then
                lBelegnr = rsrs!BELEGNR
            End If
            
            If lDatum > 0 And lBelegnr > 0 Then
                dBonPreis = ermBonPreis(lDatum, lBelegnr)
                lKUNDNR = ermBonKundnr(lDatum, lBelegnr)
                dNettoertragproBon = ermNettoertragproBon(lDatum, lBelegnr, CLng(Text1(2).Text))
                dsumNettoertrag = dsumNettoertrag + dNettoertragproBon
            End If

            If Round(dBonPreis, 2) > Round(dKVK, 2) Then
                lKUANZOver = lKUANZOver + 1
                If lKUNDNR > 0 Then
                    insert_Kunden "KundenPlus", lKUNDNR
                End If
            Else
                lKUANZUnder = lKUANZUnder + 1
                
                If lKUNDNR > 0 Then
                    insert_Kunden "KundenStan", lKUNDNR
                End If
            End If

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sTemp = "Nettoertrag der Artikel, die zusätzlich mit Artikel " & Text1(2).Text & " " & Label2(8).Caption & " in einem Bon verkauft wurden. "
    sTemp = sTemp & "(Der Aktionsartikel " & Text1(2).Text & " " & Label2(8).Caption & " ist in der Nettoertragsberechnung nicht mit eingeschlossen.)"

    anzeige "normal", sTemp, Label1(5)
    anzeige "normal", Format(dsumNettoertrag, "######0.00") & " " & gcWaehrung, Label1(7)
    
    sTemp = "Verkaufsstückzahl des Aktionsartikels " & Text1(2).Text & " " & Label2(8).Caption
    
    anzeige "normal", sTemp, Label1(9)
    anzeige "normal", Format(lsumVKmenge, "######0") & " Stück", Label1(8)
    
    If lsumKunden = 0 Then
        anzeige "normal", "kein Kunde kaufte diesen Artikel", Label1(6)
        Exit Sub
    ElseIf lsumKunden = 1 Then
        anzeige "normal", "insgesamt kaufte " & lsumKunden & " Kunde diesen Artikel", Label1(6)
    Else
        anzeige "normal", "insgesamt kauften " & lsumKunden & " Kunden diesen Artikel", Label1(6)
    End If

    dStartKreis = 0.001
    dEndKreis = 0

    Picture1.ScaleMode = vbPixels
    
    If lsumKunden <> 0 Then
        dEndKreis = 100 * lKUANZOver / lsumKunden
    End If
    
    If dEndKreis = 100 Then
        dEndKreis = 99.9
    End If
    siRadius = CDbl(Picture1.Height) / 34
    For j = dStartKreis To dEndKreis
        j = j + 1
        If j > 99.9 Then j = 99.9
        
        Call DrawPiePiece(glfarbeR(17), dStartKreis, j, Picture1, siRadius)
        PauseSi (0.02)
    Next j
    
    
    Picture1.Refresh
    Call DrawPiePiece(glfarbeR(17), dStartKreis, dEndKreis, Picture1, siRadius)
    
    If dEndKreis > 0 Then
    dStartKreis = dEndKreis
    End If
    
    dEndKreis = 99.9
    
    If dStartKreis < 99.9 Then
        siRadius = CDbl(Picture1.Height) / 34
        For j = dStartKreis To dEndKreis
            j = j + 1
            If j > 99.9 Then j = 99.9
            Call DrawPiePiece(glfarbeR(18), dStartKreis, j, Picture1, siRadius)
    
            PauseSi (0.02)
        Next j
        Call DrawPiePiece(glfarbeR(18), dStartKreis, dEndKreis, Picture1, siRadius)
    End If
    
    If lKUANZOver > 0 Then
        Label6(0).BackColor = glfarbeR(17)
        Label6(0).Visible = True
        Label6(0).Refresh
    End If
    
    If lKUANZOver = 0 Then
        anzeige "normal", "", Label5(0)
    ElseIf lKUANZOver = 1 Then
        Label5(0).Visible = True
        Command5(2).Visible = True
        anzeige "normal", lKUANZOver & " Kunde kaufte mehr als nur diesen Artikel, davon " & ermKundenCount("KUNDENPLUS") & " registrierter Kunde", Label5(0)
    Else
        Label5(0).Visible = True
        Command5(2).Visible = True
        anzeige "normal", lKUANZOver & " Kunden kauften mehr als nur diesen Artikel, davon " & ermKundenCount("KUNDENPLUS") & " registrierte Kunden", Label5(0)
    End If
    
    If lKUANZUnder > 0 Then
        Load Label6(1)
        Label6(1).Top = Label6(0).Top
        Label6(1).Top = Label6(0).Top + Label6(0).Height * (1) + 60 * (1)
        Label6(1).BackColor = glfarbeR(18)
        Label6(1).Visible = True
    
        Load Label5(1)
        Label5(1).Top = Label5(0).Top
        Label5(1).Top = Label5(0).Top + Label5(0).Height * (1) + 60 * (1)
        
        If lKUANZUnder = 0 Then
            anzeige "normal", "", Label5(1)
        ElseIf lKUANZUnder = 1 Then
            Label5(1).Visible = True
            anzeige "normal", lKUANZUnder & " Kunde kaufte nur diesen Artikel, davon " & ermKundenCount("KUNDENSTAN") & " registrierter Kunde", Label5(1)
        Else
            Label5(1).Visible = True
            anzeige "normal", lKUANZUnder & " Kunden kauften nur diesen Artikel, davon " & ermKundenCount("KUNDENSTAN") & " registrierte Kunden", Label5(1)
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 360 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "KundenAnteil"
        Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Function ermBonPreis(lDat As Long, lBon As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    ermBonPreis = 0
    
    cSQL = "Select sum(preis) as maxi from Kassjour "
    cSQL = cSQL & " where ADATE = " & Trim$(Str$(lDat))
    cSQL = cSQL & " and Belegnr = " & lBon
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermBonPreis = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBonPreis"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Function ermNettoertragproBon(lDat As Long, lBon As Long, lartnr As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    ermNettoertragproBon = 0
    
    cSQL = "Select artnr ,preis/menge as maxi,ekpr,mwst from Kassjour "
    cSQL = cSQL & " where ADATE = " & Trim$(Str$(lDat))
    cSQL = cSQL & " and Belegnr = " & lBon
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                If CLng(rsrs!artnr) <> lartnr Then
                    If Not IsNull(rsrs!maxi) Then
                        ermNettoertragproBon = ermNettoertragproBon + CDbl(NettospanneInEuro(rsrs!maxi, rsrs!ekpr, rsrs!MWST))
                    End If
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNettoertragproBon"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Sub InsertArtikelintoSammlung(lDat As Long, lBon As Long, lartnr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    
    cSQL = "insert into ARTIKELSAMMLUNG Select artnr ,bezeich, preis, menge,ekpr,mwst  from Kassjour "
    cSQL = cSQL & " where ADATE = " & Trim$(Str$(lDat))
    cSQL = cSQL & " and Belegnr = " & lBon
    gdBase.Execute cSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertArtikelintoSammlung"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Function ermBonKundnr(lDat As Long, lBon As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    ermBonKundnr = 0
    
    cSQL = "Select kundnr from Kassjour "
    cSQL = cSQL & " where ADATE = " & Trim$(Str$(lDat))
    cSQL = cSQL & " and Belegnr = " & lBon
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Kundnr) Then
            ermBonKundnr = rsrs!Kundnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBonKundnr"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Function ermKundenCount(sTab As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    ermKundenCount = 0
    
    cSQL = "select count(*) as maxi from " & sTab & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKundenCount = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKundenCount"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function

Private Sub insert_Kunden(sTab As String, lKUNDNR As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    
    
    sSQL = "Delete from " & sTab & " where kundnr  = " & lKUNDNR
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into " & sTab & " (kundnr) values (" & lKUNDNR & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insert_Kunden"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Kuteilme_fuellen(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "KUTEILME", gdBase
    
    sSQL = " Select   "
    sSQL = sSQL & " Kunden.KUNDNR as Knummer"
    sSQL = sSQL & ", Kunden.KUERZEL "
    sSQL = sSQL & ", Kunden.FIRMA "
    sSQL = sSQL & ", Kunden.TITEL "
    sSQL = sSQL & ", Kunden.NAME "
    sSQL = sSQL & ", Kunden.VORNAME "
    sSQL = sSQL & ", Kunden.STRASSE "
    sSQL = sSQL & ", Kunden.PLZ "
    sSQL = sSQL & ", Kunden.STADT "
    sSQL = sSQL & ", Kunden.TEL "
    sSQL = sSQL & ", Kunden.FAXNR "
    sSQL = sSQL & ", Kunden.MERKMAL "
    sSQL = sSQL & ", Kunden.ANREDE "
    sSQL = sSQL & ", Kunden.MERKMAL2 "
    sSQL = sSQL & ", Kunden.FORMATDAT "
    sSQL = sSQL & ", Kunden.RECHNR "
    sSQL = sSQL & ", Kunden.KURZTEXT1 "
    sSQL = sSQL & ", Kunden.KURZTEXT2 "
    sSQL = sSQL & ", format(Kunden.DATUM1,'DD.MM.YY') as datum1 "
    sSQL = sSQL & ", format(Kunden.DATUM2,'DD.MM.YY') as datum2 "
    sSQL = sSQL & ", Kunden.UMSLJ "
    sSQL = sSQL & ", Kunden.UMSVJ "
    sSQL = sSQL & ", Kunden.OSUM "
    sSQL = sSQL & ", Kunden.KASSE "
    sSQL = sSQL & ", Kunden.RABATT "
    sSQL = sSQL & ", Kunden.FILIALNR "
    sSQL = sSQL & ", Kunden.GESCHLECHT "
    sSQL = sSQL & ", Kunden.ECIDENT "
    sSQL = sSQL & ", Kunden.GESPERRT "
    sSQL = sSQL & ", Kunden.KUNDKART "
    sSQL = sSQL & ", Kunden.BONUS "
    sSQL = sSQL & ", Kunden.PREISKZ "
    sSQL = sSQL & ", Kunden.Angelegt "
    sSQL = sSQL & ", Kunden.Aender "
    sSQL = sSQL & ", Kunden.Lastdate "
    sSQL = sSQL & ", Kunden.Lasttime "
    sSQL = sSQL & ", Kunden.EMAIL "
    sSQL = sSQL & ", Kunden.MOBILTEL "
    sSQL = sSQL & ", Kunden.awm "
    sSQL = sSQL & ", ' ' as OKredit "
    sSQL = sSQL & ", '" & sdatname & "' as Datname "
    sSQL = sSQL & ", '" & sErstelldat & "' as Daterstellung "
    sSQL = sSQL & ", 0.00 as Ertrag "
    sSQL = sSQL & ", 0.00 as Umsatz "
    sSQL = sSQL & " into KUTEILME from Kunden inner join " & sTab
    sSQL = sSQL & " on " & sTab & ".kundnr = kunden.kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kuteilme_fuellen"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    bErstemal = True
    
    Option1(5).Value = True
    
    glfarbeR(0) = &H8000000F
    glfarbeR(1) = &HFFFF&
    glfarbeR(2) = &HC000&
    glfarbeR(3) = &HFF&
    glfarbeR(4) = &HC0FFFF
    glfarbeR(5) = &H80FF&
    glfarbeR(6) = &HFF00FF
    glfarbeR(7) = &HFFFF00
    glfarbeR(8) = &HC0C0FF
    glfarbeR(9) = &HFFC0C0
    glfarbeR(10) = &H8080FF
    glfarbeR(11) = &HC0FFC0
    glfarbeR(12) = &HFF8080
    glfarbeR(13) = &H40C0&
    glfarbeR(14) = &H800080
    glfarbeR(15) = &H80&
    glfarbeR(16) = &H808000
    glfarbeR(17) = &HC0C0&
    glfarbeR(18) = &HFF80FF
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "KUANT", gdBase
    loeschNEW "KUNDENSTAN", gdBase
    loeschNEW "KUNDENPLUS", gdBase
    
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
Private Sub Label5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index

    Case 0
        Kuteilme_fuellen "KUNDENPLUS"
    Case 1
        Kuteilme_fuellen "KUNDENSTAN"
End Select

frmWKL173.Show 1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0     'vorjahr
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(1).Text = Format("31.12." & Year(Now) - 1, "DD.MM.YYYY")
            
        Case Is = 2    'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(1).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        Text1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
              
        Case Is = 5     'ak monat
            Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
        
        Case Is = 6     'gestern
            Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        
        Case Is = 7     'heute
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
        Case 12 'aktuelles Jahr
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    End Select
    
    Command5_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command5_Click 1
    End If
    If KeyCode = vbKeyEscape Then
        Command5_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

     Select Case Index
        Case 2
        
            If Len(Text1(Index).Text) > 5 Then
            
                If IsNumeric(Text1(Index).Text) Then
                
                    Label2(8).Caption = bezis(Text1(Index).Text)
                    Label2(8).Refresh
                Else
            
                    Label2(8).Caption = ""
                    Label2(8).Refresh
                    
                End If
            
            Else
            
                Label2(8).Caption = ""
                Label2(8).Refresh
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 2
            cValid = "1234567890" & Chr$(8)
        Case 3
            cValid = "1234567890," & Chr$(8)
    End Select

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)

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
    Fehler.gsFehlertext = "Im Programmteil Mailing Feedback ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
