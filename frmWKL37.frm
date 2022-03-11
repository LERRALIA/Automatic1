VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL37 
   Caption         =   "Protokoll der Bestandsveränderungen"
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
   Icon            =   "frmWKL37.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
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
      Height          =   375
      Index           =   0
      Left            =   9240
      MaxLength       =   6
      TabIndex        =   42
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   6960
      TabIndex        =   21
      Top             =   3840
      Width           =   4815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Shopverkauf"
         Height          =   255
         Index           =   21
         Left            =   2520
         TabIndex        =   50
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Personalkauf"
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   49
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Wundertüte"
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   48
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Retoure"
         Height          =   255
         Index           =   18
         Left            =   2520
         TabIndex        =   47
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Auflösen"
         Height          =   255
         Index           =   17
         Left            =   2520
         TabIndex        =   46
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle"
         Height          =   255
         Index           =   16
         Left            =   2520
         TabIndex        =   45
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Fehllieferung"
         Height          =   255
         Index           =   15
         Left            =   2520
         TabIndex        =   41
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Spende"
         Height          =   255
         Index           =   14
         Left            =   2520
         TabIndex        =   34
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Eigenbedarf"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Ladenbedarf"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   32
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unklare Bestandsdiff"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bruch"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   2055
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   6
         Left            =   2520
         TabIndex        =   27
         Top             =   3240
         Width           =   2175
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Diebstahl"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "reg. Warenentnahme"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bedienerfehler"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Verfallsdat erreicht"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bestandskorrektur"
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
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   """Rücknahme Kasse"" ausschließen"
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   2040
      Width           =   3255
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
      Left            =   1320
      TabIndex        =   14
      Top             =   1920
      Width           =   2415
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "aktuelles Jahr"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vormonat"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "aktueller Monat"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gestern"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Heute"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
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
         TabIndex        =   19
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
      Index           =   3
      Left            =   1440
      TabIndex        =   11
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
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Tag             =   "2"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "nur Stornobuchungen"
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Verkäufe einschließen / Filialtäusche"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   1680
      Width           =   3495
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
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
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
      Left            =   5640
      TabIndex        =   1
      Text            =   "alle"
      Top             =   2880
      Width           =   3495
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9480
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
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
      Left            =   9480
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
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
   Begin sevCommand3.Command Command1 
      Height          =   405
      Index           =   20
      Left            =   3120
      TabIndex        =   35
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
   Begin sevCommand3.Command Command1 
      Height          =   405
      Index           =   21
      Left            =   3120
      TabIndex        =   36
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
   Begin sevCommand3.Command Command1 
      Height          =   165
      Index           =   3
      Left            =   2760
      TabIndex        =   37
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
   Begin sevCommand3.Command Command1 
      Height          =   165
      Index           =   2
      Left            =   2760
      TabIndex        =   38
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
   Begin sevCommand3.Command Command1 
      Height          =   165
      Index           =   4
      Left            =   2760
      TabIndex        =   39
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
   Begin sevCommand3.Command Command1 
      Height          =   165
      Index           =   5
      Left            =   2760
      TabIndex        =   40
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
   Begin sevCommand3.Command Command1 
      Height          =   345
      Index           =   9
      Left            =   8280
      TabIndex        =   44
      Top             =   4200
      Width           =   360
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
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferant"
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
      Left            =   9240
      TabIndex        =   43
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Datum von:"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Datum bis:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Bediener"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   7800
      Width           =   10815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelnummer / EAN"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Protokoll der Bestandsveränderungen"
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
      TabIndex        =   4
      Top             =   120
      Width           =   11175
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
End
Attribute VB_Name = "frmWKL37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "bestpdru_kopf" & srechnertab, gdApp
    loeschNEW "bestpdru_kopf" & srechnertab, gdBase
    
    loeschNEW "bestpdru_kopf", gdBase
    loeschNEW "bestpdru_kopf", gdApp
    
    loeschNEW "BESTPDRU" & srechnertab, gdApp
    loeschNEW "BESTPDRU" & srechnertab, gdBase
    
    loeschNEW "BESTPDRU", gdBase
    loeschNEW "BESTPDRU", gdApp
    
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
Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check1.value = vbChecked Then
        Check2.Visible = True
    Else
        Check2.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    Dim sSQL As String
    
    Select Case index
        Case 0
        
            If gbOhnebestProt = True Then
            
                anzeige "rot", "Die Bestandsprotkollierung ist nicht aktiviert", lblanzeige
            Else
            
                If SucheDaten Then
                
                    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
                    
                    
                    loeschNEW "BESTPDRU" & srechnertab, gdApp
                    loeschNEW "BESTPDRU", gdApp
                    
                    TransferTab gdBase, App.Path & "\kissapp.mdb", "BESTPDRU" & srechnertab
                    
                    sSQL = "select * into BESTPDRU from BESTPDRU" & srechnertab
                    gdApp.Execute sSQL, dbFailOnError
                    
                    'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
                     'diese Funktion (GetRetoure) habe ich neu implementiert (die war nicht)
                     GetRetoure
                     
                     Dim rsrs As Recordset
                     Set rsrs = gdApp.OpenRecordset("BESTPDRU", dbOpenTable)
                      
                     If rsrs.EOF Then
                        lblanzeige.Caption = "Keine Daten ermittelt."
                        lblanzeige.Refresh
                        rsrs.Close: Set rsrs = Nothing
                        Screen.MousePointer = 0
                        Exit Sub
                     Else
                        anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
                        rsrs.Close: Set rsrs = Nothing
                     End If
                    
                    'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
                    
                    reportbildschirmApp "", "aWKL37"
                    anzeige "normal", "", lblanzeige
                
                End If
                
            End If
        Case 1
            Unload frmWKL37
        Case 2
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat + 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 3
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat - 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 4
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat + 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 5
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat - 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 6
        
            If Suchedaten_quick Then
                
                anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
                
                loeschNEW "BESTPDRU" & srechnertab, gdApp
                loeschNEW "BESTPDRU", gdApp
                TransferTab gdBase, App.Path & "\kissapp.mdb", "BESTPDRU" & srechnertab
                
                sSQL = "select * into BESTPDRU from BESTPDRU" & srechnertab
                gdApp.Execute sSQL, dbFailOnError
                
                loeschNEW "BESTPDRU_KOPF" & srechnertab, gdApp
                loeschNEW "BESTPDRU_KOPF", gdApp
                TransferTab gdBase, App.Path & "\kissapp.mdb", "BESTPDRU_KOPF" & srechnertab
                
                sSQL = "select * into BESTPDRU_KOPF from BESTPDRU_KOPF" & srechnertab
                gdApp.Execute sSQL, dbFailOnError
                
                reportbildschirmApp "", "aWKL37a"
                anzeige "normal", "", lblanzeige
                
            End If
        Case 9
            Text1_KeyUp 0, vbKeyF2, 0

        Case 20         ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(2).SetFocus
            
        Case 21         ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    WKL37Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    füllecboBediener cboBed
    
    Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermittleartnr(sSuchstring As String) As String
    On Error GoTo LOKAL_ERROR
    
    ermittleartnr = ""
    
    Dim cSQL As String
    Dim rsArt As Recordset
    
    cSQL = "select * from artikel where ean = '" & sSuchstring & "'"
    cSQL = cSQL & " or ean2 = '" & sSuchstring & "'"
    cSQL = cSQL & " or ean3 = '" & sSuchstring & "'"
    
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.RecordCount = 0 Then
        If Not IsNull(rsArt!artnr) Then
            ermittleartnr = CStr(rsArt!artnr)
        End If
    End If
    
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleartnr"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SucheDaten() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sWhere      As String
    Dim sTauwhere   As String
    Dim sKassjWhere As String
    Dim cVon        As String
    Dim lVon        As Long
    Dim cBis        As String
    Dim lBis        As Long
    Dim sSQL        As String
    Dim sArt        As String
    Dim sLinr       As String
    Dim sBedname    As String
    Dim ibednu      As Integer
    Dim rsrs        As Recordset
    
    SucheDaten = False
    
    lblanzeige.Caption = "Daten werden ermittelt..."
    lblanzeige.Refresh
    
'    Datenbankwechsel
    
    Screen.MousePointer = 11
    Me.Refresh
    
    If Text1(3).Text = "" Then Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    If Text1(2).Text = "" Then Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    sSQL = "Delete from Bestprot where aenart = 'Kassiervorgang'"
    gdBase.Execute sSQL, dbFailOnError
    
    If Text1(3).Text <> "" Then
        If IsDate(Text1(3).Text) Then
            cVon = Text1(3).Text
            lVon = DateValue(cVon)
            sWhere = " and Lastdate >= " & Trim$(Str$(lVon))
            sKassjWhere = " where adate >= " & Trim$(Str$(lVon))
        Else
            lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
            lblanzeige.Refresh
            Text1(3).SetFocus
            Exit Function
        End If
    End If
    
    If Text1(2).Text <> "" Then
        If IsDate(Text1(2).Text) Then
            cBis = Text1(2).Text
            lBis = DateValue(cBis)
            sWhere = sWhere & " and Lastdate <= " & Trim$(Str$(lBis))
            sKassjWhere = sKassjWhere & " and adate <= " & Trim$(Str$(lBis))
        Else
            lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
            lblanzeige.Refresh
            Text1(2).SetFocus
            Exit Function
        End If
    End If
    
    Text1(1).Text = Trim(Text1(1).Text)
    If Text1(1).Text <> "" Then
        
        If Len(Text1(1).Text) > 6 Then
            sArt = ermittleartnr(Text1(1).Text)
        Else
            sArt = Text1(1).Text
        End If
        
        If sArt <> "" Then
            If IsNumeric(sArt) Then
                If Trim(sWhere) = "" Then
                    sWhere = " and artnr = " & sArt
                Else
                    sWhere = sWhere & " and artnr = " & sArt
                End If
                
                If Trim(sKassjWhere) = "" Then
                    sKassjWhere = " where artnr = " & sArt
                Else
                    sKassjWhere = sKassjWhere & " and artnr = " & sArt
                End If
            End If
        End If
            
    End If
    
    
    
    
    

    sBedname = Trim(cboBed.Text)
    If sBedname = "alle" Then
        sTauwhere = sKassjWhere
    Else
        sSQL = " Select Bednu from bedname "
        sSQL = sSQL & " where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
        rsBed.MoveFirst
            If Not IsNull(rsBed!BEDNU) Then
                ibednu = rsBed!BEDNU
            Else
                ibednu = 0
            End If
        Else
            ibednu = 0
        End If
        rsBed.Close: Set rsBed = Nothing
        
        If ibednu <> 0 Then
            If Trim(sWhere) = "" Then
                sWhere = " and Bediener = " & ibednu
            Else
                sWhere = sWhere & " and Bediener = " & ibednu
            End If
            
            sTauwhere = sKassjWhere
            
            If Trim(sKassjWhere) = "" Then
                sKassjWhere = " where Bediener = " & ibednu
            Else
                sKassjWhere = sKassjWhere & " and Bediener = " & ibednu
            End If
        End If
    End If
    
    loeschNEW "BESTPDRU" & srechnertab, gdBase
    CreateTableT2 "BESTPDRU" & srechnertab, gdBase

    sSQL = "Insert into BESTPDRU" & srechnertab & " select * from BestProt where aenart <> 'Kassiervorgang' "
    sSQL = sSQL & sWhere
    gdBase.Execute sSQL, dbFailOnError
    
    If Check1.value = vbChecked Then
    
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener, 'Kassiervorgang' as aenart "
        sSQL = sSQL & ", Filiale "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", best1 as newbest "
        sSQL = sSQL & ", (best1 + menge) as oldbest "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & " from Kassjour  "
        sSQL = sSQL & sKassjWhere
        
        If Trim(sKassjWhere) = "" Then
            sSQL = sSQL & " where menge > 0 "
        Else
            sSQL = sSQL & " and menge > 0 "
        End If
        gdBase.Execute sSQL, dbFailOnError
        
        
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener, 'Kollegenverkauf' as aenart "
        sSQL = sSQL & ", Filiale "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", best1 as newbest "
        sSQL = sSQL & ", (best1 + menge) as oldbest "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & " from Kollverk  "
        sSQL = sSQL & sKassjWhere
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener,'Filialtausch' as aenart "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & ", Kasnum  as newbest "
        sSQL = sSQL & ", (Kasnum + menge) as oldbest "
        sSQL = sSQL & " from Tausch  "
        sSQL = sSQL & sTauwhere

        gdBase.Execute sSQL, dbFailOnError
        
        If Check2.value = vbChecked Then
            sSQL = "Delete from BESTPDRU" & srechnertab & " where aenart = 'Kassiervorgang' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    If Check3.value = vbChecked Then
        sSQL = "Delete from BESTPDRU" & srechnertab & " where aenart = 'Rücknahme Kasse' "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    'Lieferant
    sLinr = Trim(Text1(0).Text)
    If sLinr <> "" Then
        
        If sLinr <> "" Then
            If IsNumeric(sLinr) Then
            
                loeschNEW "sic_BESTPDRU" & srechnertab, gdBase
    
                sSQL = "Select BESTPDRU" & srechnertab & ".* into sic_BESTPDRU" & srechnertab & " from BESTPDRU" & srechnertab & " inner join artlief on"
                sSQL = sSQL & " BESTPDRU" & srechnertab & ".artnr = Artlief.artnr where Artlief.linr = " & sLinr
                gdBase.Execute sSQL, dbFailOnError
            
                loeschNEW "BESTPDRU" & srechnertab, gdBase
            
                sSQL = "Select * into BESTPDRU" & srechnertab & " from sic_BESTPDRU" & srechnertab
                gdBase.Execute sSQL, dbFailOnError
            
                loeschNEW "sic_BESTPDRU" & srechnertab, gdBase

            End If
        End If
            
    End If
    
    Set rsrs = gdBase.OpenRecordset("BESTPDRU" & srechnertab, dbOpenTable)
    If rsrs.EOF Then
        lblanzeige.Caption = "Keine Daten ermittelt."
        lblanzeige.Refresh
        'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
        'rsrs.Close: Set rsrs = Nothing
        'Exit Function
        'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update BESTPDRU" & srechnertab & " inner join bedname on BESTPDRU" & srechnertab & ".Bediener = bedname.bednu "
    sSQL = sSQL & " Set BESTPDRU" & srechnertab & ".bedname = bedname.bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTPDRU" & srechnertab & " inner join Artikel on BESTPDRU" & srechnertab & ".Artnr = Artikel.Artnr "
    sSQL = sSQL & " Set BESTPDRU" & srechnertab & ".bezeich = Artikel.bezeich"
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".FARBNR = VAL(Artikel.AWM)"
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".LINR = Artikel.LINR "
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".LPZ = Artikel.LPZ "
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".EKPR = Artikel.EKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    Markenabgleich "BESTPDRU" & srechnertab, gdBase
    
    BringFarbeInsSpiel "BESTPDRU" & srechnertab, gdBase
        
    SucheDaten = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next

End Function
Private Function Suchedaten_quick() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sWhere          As String
    Dim sTauwhere       As String
    Dim sKassjWhere     As String
    Dim cVon            As String
    Dim lVon            As Long
    Dim cBis            As String
    Dim lBis            As Long
    Dim sSQL            As String
    Dim sArt            As String
    Dim sLinr           As String
    Dim sBedname        As String
    Dim ibednu          As Integer
    Dim rsrs            As Recordset
    Dim caenderGRUND    As String
    
    Suchedaten_quick = False
    
    lblanzeige.Caption = "Daten werden ermittelt..."
    lblanzeige.Refresh
    
    Screen.MousePointer = 11
    Me.Refresh
    
    If Text1(3).Text = "" Then Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    If Text1(2).Text = "" Then Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    sSQL = "Delete from Bestprot where aenart = 'Kassiervorgang'"
    gdBase.Execute sSQL, dbFailOnError
    
    If Text1(3).Text <> "" Then
        If IsDate(Text1(3).Text) Then
            cVon = Text1(3).Text
            lVon = DateValue(cVon)
            sWhere = " and Lastdate >= " & Trim$(Str$(lVon))
            sKassjWhere = " where adate >= " & Trim$(Str$(lVon))
        Else
            lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
            lblanzeige.Refresh
            Text1(3).SetFocus
            Exit Function
        End If
    End If
    
    If Text1(2).Text <> "" Then
        If IsDate(Text1(2).Text) Then
            cBis = Text1(2).Text
            lBis = DateValue(cBis)
            sWhere = sWhere & " and Lastdate <= " & Trim$(Str$(lBis))
            sKassjWhere = sKassjWhere & " and adate <= " & Trim$(Str$(lBis))
        Else
            lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
            lblanzeige.Refresh
            Text1(2).SetFocus
            Exit Function
        End If
    End If
    
    Text1(1).Text = Trim(Text1(1).Text)
    If Text1(1).Text <> "" Then
        
        If Len(Text1(1).Text) > 6 Then
            sArt = ermittleartnr(Text1(1).Text)
        Else
            sArt = Text1(1).Text
        End If
        
        If sArt <> "" Then
            If IsNumeric(sArt) Then
                If Trim(sWhere) = "" Then
                    sWhere = " and artnr = " & sArt
                Else
                    sWhere = sWhere & " and artnr = " & sArt
                End If
                
                If Trim(sKassjWhere) = "" Then
                    sKassjWhere = " where artnr = " & sArt
                Else
                    sKassjWhere = sKassjWhere & " and artnr = " & sArt
                End If
            End If
        End If
            
    End If

    sBedname = Trim(cboBed.Text)
    If sBedname = "alle" Then
        sTauwhere = sKassjWhere
    Else
        sSQL = " Select Bednu from bedname "
        sSQL = sSQL & " where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
        rsBed.MoveFirst
            If Not IsNull(rsBed!BEDNU) Then
                ibednu = rsBed!BEDNU
            Else
                ibednu = 0
            End If
        Else
            ibednu = 0
        End If
        rsBed.Close: Set rsBed = Nothing
        
        If ibednu <> 0 Then
            If Trim(sWhere) = "" Then
                sWhere = " and Bediener = " & ibednu
            Else
                sWhere = sWhere & " and Bediener = " & ibednu
            End If
            
            sTauwhere = sKassjWhere
            
            If Trim(sKassjWhere) = "" Then
                sKassjWhere = " where Bediener = " & ibednu
            Else
                sKassjWhere = sKassjWhere & " and Bediener = " & ibednu
            End If
        End If
    End If
    
    loeschNEW "BESTPDRU" & srechnertab, gdBase
    CreateTableT2 "BESTPDRU" & srechnertab, gdBase
    

    sSQL = "Insert into BESTPDRU" & srechnertab & " select * from BestProt where aenart <> 'Kassiervorgang' "
    
    If Option1(4).value = True Then 'Diebstahl
        caenderGRUND = "Diebstahl"
    ElseIf Option1(3).value = True Then ' reg. Warenentnahme
        caenderGRUND = "reg. Warenentnahme"
    ElseIf Option1(1).value = True Then ' Bedienerfehler
        caenderGRUND = "Bedienerfehler"
    ElseIf Option1(0).value = True Then ' Verfallsdat erreicht
        caenderGRUND = "Verfallsdat erreicht"
    ElseIf Option1(8).value = True Then ' Bruch
        caenderGRUND = "Bruch"
    ElseIf Option1(11).value = True Then 'unklare Bestandsdiff
        caenderGRUND = "unklare Bestandsdiff"
    ElseIf Option1(12).value = True Then 'Ladenbedarf
        caenderGRUND = "Ladenbedarf"
    ElseIf Option1(13).value = True Then 'Eigenbedarf
        caenderGRUND = "Eigenbedarf"
    ElseIf Option1(14).value = True Then 'Spende
        caenderGRUND = "Spende"
    ElseIf Option1(15).value = True Then 'Fehllieferung
        caenderGRUND = "Fehllieferung"
        
    ElseIf Option1(17).value = True Then 'Auflösen
        caenderGRUND = "Auflösen"
    ElseIf Option1(18).value = True Then 'Retoure
        caenderGRUND = "Retoure"
    ElseIf Option1(19).value = True Then 'Wundertüte
        caenderGRUND = "Wundertüte"
    ElseIf Option1(20).value = True Then 'Personalkauf
        caenderGRUND = "Personalkauf"
    ElseIf Option1(21).value = True Then 'Shopverkauf
        caenderGRUND = "Shopverkauf"
    ElseIf Option1(16).value = True Then 'alle
        caenderGRUND = "alle"
    
    End If
    
    If Option1(16).value = True Then 'alle
        sSQL = sSQL & " and AENGRUND <> ''"
    Else
    
        sSQL = sSQL & " and AENGRUND = '" & caenderGRUND & "' "
    
    End If
                       
    sSQL = sSQL & sWhere
    
    gdBase.Execute sSQL, dbFailOnError
    
    If Check1.value = vbChecked Then
    
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener, 'Kassiervorgang' as aenart "
        sSQL = sSQL & ", Filiale "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", best1 as newbest "
        sSQL = sSQL & ", (best1 + menge) as oldbest "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & " from Kassjour  "
        sSQL = sSQL & sKassjWhere
        
        If Trim(sKassjWhere) = "" Then
            sSQL = sSQL & " where menge > 0 "
        Else
            sSQL = sSQL & " and menge > 0 "
        End If
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener, 'Kollegenverkauf' as aenart "
        sSQL = sSQL & ", Filiale "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", best1 as newbest "
        sSQL = sSQL & ", (best1 + menge) as oldbest "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & " from Kollverk  "
        sSQL = sSQL & sKassjWhere
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into BESTPDRU" & srechnertab & " select artnr, Bediener,'Filialtausch' as aenart "
        sSQL = sSQL & ", adate as Lastdate "
        sSQL = sSQL & ", azeit as lasttime "
        sSQL = sSQL & ", menge as umenge "
        sSQL = sSQL & ", Kasnum  as newbest "
        sSQL = sSQL & ", (Kasnum + menge) as oldbest "
        sSQL = sSQL & " from Tausch  "
        sSQL = sSQL & sTauwhere
        gdBase.Execute sSQL, dbFailOnError
        
        If Check2.value = vbChecked Then
            sSQL = "Delete from BESTPDRU" & srechnertab & " where aenart = 'Kassiervorgang' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    If Check3.value = vbChecked Then
        sSQL = "Delete from BESTPDRU" & srechnertab & " where aenart = 'Rücknahme Kasse' "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    'Lieferant
    sLinr = Trim(Text1(0).Text)
    If sLinr <> "" Then
        
        If sLinr <> "" Then
            If IsNumeric(sLinr) Then
            
                loeschNEW "sic_BESTPDRU" & srechnertab, gdBase
    
                sSQL = "Select BESTPDRU" & srechnertab & ".* into sic_BESTPDRU" & srechnertab & " from BESTPDRU" & srechnertab & " inner join artlief on"
                sSQL = sSQL & " BESTPDRU" & srechnertab & ".artnr = Artlief.artnr where Artlief.linr = " & sLinr
                gdBase.Execute sSQL, dbFailOnError
            
                loeschNEW "BESTPDRU" & srechnertab, gdBase
            
                sSQL = "Select * into BESTPDRU" & srechnertab & " from sic_BESTPDRU" & srechnertab
                gdBase.Execute sSQL, dbFailOnError
            
                loeschNEW "sic_BESTPDRU" & srechnertab, gdBase

            End If
        End If
            
    End If
    
    
    Set rsrs = gdBase.OpenRecordset("BESTPDRU" & srechnertab, dbOpenTable)
    If rsrs.EOF Then
        lblanzeige.Caption = "Keine Daten ermittelt."
        lblanzeige.Refresh
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update BESTPDRU" & srechnertab & " inner join bedname on BESTPDRU" & srechnertab & ".Bediener = bedname.bednu "
    sSQL = sSQL & " Set BESTPDRU" & srechnertab & ".bedname = bedname.bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BESTPDRU" & srechnertab & " inner join Artikel on BESTPDRU" & srechnertab & ".Artnr = Artikel.Artnr "
    sSQL = sSQL & " Set BESTPDRU" & srechnertab & ".bezeich = Artikel.bezeich"
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".FARBNR = VAL(Artikel.AWM)"
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".LINR = Artikel.LINR "
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".LPZ = Artikel.LPZ "
    sSQL = sSQL & " , BESTPDRU" & srechnertab & ".EKPR = Artikel.EKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    Markenabgleich "BESTPDRU" & srechnertab, gdBase
    
    BringFarbeInsSpiel "BESTPDRU" & srechnertab, gdBase
    
    loeschNEW "BESTPDRU_KOPF" & srechnertab, gdBase
    CreateTableT2 "BESTPDRU_KOPF" & srechnertab, gdBase
    
    sSQL = "Insert into BESTPDRU_KOPF" & srechnertab & " (VON,BIS,AENGRUND "
    sSQL = sSQL & " ) values ("
    sSQL = sSQL & " '" & cVon & "'  "
    sSQL = sSQL & ", '" & cBis & "'  "
    sSQL = sSQL & ", '" & caenderGRUND & "'  "
    sSQL = sSQL & "  ) "
    gdBase.Execute sSQL, dbFailOnError
    
    Suchedaten_quick = True
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Suchedaten_quick"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub WKL37Positionieren()
    On Error GoTo LOKAL_ERROR
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL37Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
    
        Case Is = 2    'vormonat
        
            If Month(DateValue(Now)) = 1 Then
                Text1(3).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(2).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(3).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(2).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(2).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(2).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        Text1(2).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
                
        Case Is = 5     'ak monat
            Text1(3).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
        
        Case Is = 6     'gestern
            Text1(3).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Text1(2).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        
        Case Is = 7     'heute
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 9 'akt Jahr
            Text1(3).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 10 'vorjahr
            Text1(3).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(2).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = glSelBack1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

Select Case index

    Case 0
    
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            
            
            gF2Prompt.bMultiple = False
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            If gF2Prompt.cWahl <> "" Then
                Text1(0).Text = gF2Prompt.cWahl
            End If
                    
                
            Text1(0).SetFocus
        End If
End Select

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtlief_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbobed_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    cboBed.SelStart = 0
    cboBed.SelLength = Len(cboBed.Text)
    cboBed.BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbobed_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbobed_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cboBed.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbobed_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
Private Sub GetRetoure()
On Error GoTo LOKAL_ERROR
     
Dim Scmd As String
Dim sArt As String
Dim sBedname As String
Dim ibednu As Integer
 

Scmd = "INSERT INTO [MS Access;Database=" & App.Path & "\kissapp.mdb].BESTPDRU(LASTDATE,LASTTIME,ARTNR,BEZEICH,Bediener,AENART,AENGRUND,UMENGE,NEWBEST,OLDBEST,SYNSTATUS,FILIALE,SENDOK,BEDNAME,FARBTEXT,FARBwert,FARBwertS,FARBNR,EKPR,LINR,LPZ,MARKE,LINBEZ)SELECT"
Scmd = Scmd & " ADATE as LASTDATE,AZEIT as LASTTIME,ARTNR,BEZEICH,BEDIENER,'Retoure' as AENART,'Retoure' as AENGRUND,BEST1 as UMENGE,'-' & MENGE as NEWBEST,(MENGE+BEST1) as OLDBEST,''as SYNSTATUS,FILIALE,SENDOK,'' as BEDNAME,''as FARBTEXT,null as FARBwert,null as FARBwertS , 0 as FARBNR,EKPR,LINR,LPZ,'' as MARKE,'' as LINBEZ FROM RETOURE"


If Text1(3).Text <> "" Then
 If IsDate(Text1(3).Text) Then
     
     Scmd = Scmd & " WHERE Datevalue(ADATE) >= CDate('" & Text1(3).Text & "')"
            
 Else
      lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
      lblanzeige.Refresh
      Text1(3).SetFocus
      Exit Sub
 End If
End If
    
If Text1(2).Text <> "" Then
 If IsDate(Text1(2).Text) Then
     
     Scmd = Scmd & " AND Datevalue(ADATE) <= CDate('" & Text1(2).Text & "')"
            
 Else
      lblanzeige.Caption = "Bitte geben Sie ein richtiges Datumsformat ein!"
      lblanzeige.Refresh
      Text1(2).SetFocus
      Exit Sub
 End If
End If
    
Text1(1).Text = Trim(Text1(1).Text)
If Text1(1).Text <> "" Then

        If Len(Text1(1).Text) > 6 Then
            sArt = ermittleartnr(Text1(1).Text)
        Else
            sArt = Text1(1).Text
        End If

        If sArt <> "" Then
            If IsNumeric(sArt) Then
               Scmd = Scmd & " AND ARTNR=" & sArt
            End If
        End If

End If

    sBedname = Trim(cboBed.Text)
    If sBedname <> "alle" Then
       
            Dim rsBed As Recordset
            Set rsBed = gdBase.OpenRecordset("SELECT BEDNU FROM BEDNAME WHERE BEDNAME='" & sBedname & "'")
            
            If Not rsBed.EOF Then
                rsBed.MoveFirst
                If Not IsNull(rsBed!BEDNU) Then
                    ibednu = rsBed!BEDNU
                Else
                    ibednu = 0
                End If
            Else
                ibednu = 0
            End If
            rsBed.Close: Set rsBed = Nothing
    
    Else
            ibednu = 0
      
    End If

If ibednu <> 0 Then
    Scmd = Scmd & " AND BEDIENER=" & ibednu
End If
    
gdBase.Execute Scmd, dbFailOnError
 
gdBase.Execute "UPDATE BEDNAME BN INNER JOIN [MS Access;Database=" & App.Path & "\kissapp.mdb].BESTPDRU B ON BN.BEDNU=B.BEDIENER SET B.BEDNAME=BN.BEDNAME", dbFailOnError
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GetRetoure"
    Fehler.gsFehlertext = "Im Programmteil Protokoll der Bestandsveränderungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE

