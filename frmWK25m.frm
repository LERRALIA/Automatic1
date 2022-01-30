VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWK25m 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Artikelverkäufe nach Zugang"
   ClientHeight    =   8625
   ClientLeft      =   1755
   ClientTop       =   1845
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWK25m.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
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
      Left            =   6000
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   11595
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Auswahlkriterien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11655
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
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   4455
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "Vorjahr"
            Height          =   255
            Index           =   14
            Left            =   2760
            TabIndex        =   27
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "aktuelles Jahr"
            Height          =   255
            Index           =   12
            Left            =   2760
            TabIndex        =   26
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "Heute"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "Gestern"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "aktueller Monat"
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   23
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C000&
            Caption         =   "Vormonat"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4680
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   315
         Index           =   1
         Left            =   10440
         TabIndex        =   4
         Top             =   960
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   6
         Left            =   6000
         TabIndex        =   10
         Top             =   480
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
      Begin sevCommand3.Command Command20 
         Height          =   360
         Index           =   21
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "Kalender"
         Top             =   480
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
      Begin sevCommand3.Command Command20 
         Height          =   360
         Index           =   20
         Left            =   1440
         TabIndex        =   13
         ToolTipText     =   "Kalender"
         Top             =   480
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
      Begin sevCommand3.Command Command20 
         Height          =   360
         Index           =   0
         Left            =   9600
         TabIndex        =   19
         ToolTipText     =   "Kalender"
         Top             =   480
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
      Begin sevCommand3.Command Command20 
         Height          =   360
         Index           =   1
         Left            =   7920
         TabIndex        =   20
         ToolTipText     =   "Kalender"
         Top             =   480
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Verkaufszeitraum"
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
         Left            =   6600
         TabIndex        =   29
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Zugangszeitraum"
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
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
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
         Index           =   4
         Left            =   6600
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
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
         Left            =   8400
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   15
      FixedCols       =   2
      ForeColorSel    =   8454143
      FocusRect       =   0
      HighLight       =   2
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
   End
   Begin sevCommand3.Command Command1 
      Height          =   315
      Index           =   0
      Left            =   10560
      TabIndex        =   9
      Top             =   7920
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   12
      Left            =   11400
      TabIndex        =   30
      Top             =   120
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
   Begin sevCommand3.Command Command11 
      Height          =   360
      Left            =   10920
      TabIndex        =   31
      Top             =   120
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
      Picture         =   "frmWK25m.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "76543210,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   37
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "76543210,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   36
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Nettoverkaufswert"
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
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Einkaufswert"
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
      Index           =   3
      Left            =   2880
      TabIndex        =   34
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie ein Suchkriterium ein!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   9135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelverkäufe nach Zugang"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmWK25m"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SpaltennummerArtnr          As Byte

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 1 'suchen
            ermittleVerkaufNZU Text1(4).Text
            ZeigeArtikelVERKAUFNZU
        Case 0
            Unload frmWK25m
        Case 6
            Text1_KeyUp 4, vbKeyF2, 0
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermittleVerkaufNZU(sLinr1 As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim siAnzeige As Single
    
    Dim lDatVonZU     As Long
    Dim lDatBisZU     As Long
    Dim cDatVonZU     As String
    Dim cDatBisZU     As String
   
    cDatVonZU = Trim$(Text1(2).Text)
    cDatBisZU = Trim$(Text1(3).Text)
    
    If cDatVonZU <> "" Then
        lDatVonZU = Fix(DateValue(cDatVonZU))
        cDatVonZU = Trim$(Str$(lDatVonZU))
    End If
    
    If cDatBisZU <> "" Then
        lDatBisZU = Fix(DateValue(cDatBisZU))
        cDatBisZU = Trim$(Str$(lDatBisZU))
    End If
    
    Dim lDatVonVK     As Long
    Dim lDatBisVK     As Long
    Dim cDatVonVK     As String
    Dim cDatBisVK     As String
   
    cDatVonVK = Trim$(Text1(1).Text)
    cDatBisVK = Trim$(Text1(0).Text)
    
    If cDatVonVK <> "" Then
        lDatVonVK = Fix(DateValue(cDatVonVK))
        cDatVonVK = Trim$(Str$(lDatVonVK))
    End If
    
    If cDatBisVK <> "" Then
        lDatBisVK = Fix(DateValue(cDatBisVK))
        cDatBisVK = Trim$(Str$(lDatBisVK))
    End If
    
    Label3(0).Visible = False
    Label3(1).Visible = False
    
    If sLinr1 = "" Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "VERKAUFNZU", gdBase
    CreateTableT2 "VERKAUFNZU", gdBase
    
    sSQL = "Create index artnr on VERKAUFNZU(artnr) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "die Artikel werden ermittelt...", Label6
    
    sSQL = " Insert into VERKAUFNZU select ARTNR , sum(Bewegung) as Zugangsmenge from ZUGANG where linr = " & sLinr1
    
    If cDatVonZU <> "" Then
        sSQL = sSQL & " and Zugang.ADATE >= " & cDatVonZU & " "
    End If
    
    If cDatBisZU <> "" Then
        sSQL = sSQL & " and Zugang.ADATE <= " & cDatBisZU & " "
    End If
    
    
    sSQL = sSQL & " group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    

    
    
    
    
    txtStatus.Text = 12
    
    loeschNEW "VERKAUFNZU_DAT2", gdBase

    sSQL = "select ARTNR,sum(menge) as verkaufsmenge,sum(preis) as verkaufspreis  into VERKAUFNZU_DAT2 "
    sSQL = sSQL & " from Kassjour "
    
    sSQL = sSQL & " where Artnr in (Select artnr from VERKAUFNZU) "
    
    If cDatVonVK <> "" Then
        sSQL = sSQL & " and Kassjour.ADATE >= " & cDatVonVK & " "
    End If
    
    If cDatBisVK <> "" Then
        sSQL = sSQL & " and Kassjour.ADATE <= " & cDatBisVK & " "
    End If
    
    sSQL = sSQL & " group by Artnr "
    gdBase.Execute sSQL, dbFailOnError


    txtStatus.Text = 25

    sSQL = " Update VERKAUFNZU set VERKAUFSMENGE = 0"
    sSQL = sSQL & ", VERKAUFSPREIS = 0 "
    sSQL = sSQL & ", LEKPR = 0 "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 40

    sSQL = " Update VERKAUFNZU inner join Artlief on VERKAUFNZU.artnr = Artlief.artnr "
    sSQL = sSQL & " set VERKAUFNZU.LEKPR = Artlief.lekpr where linr = " & sLinr1
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45

    sSQL = " Update VERKAUFNZU set ZUGANGSWERT = LEKPR * ZUGANGSMENGE"
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 75

    sSQL = " Update VERKAUFNZU inner join Artikel on VERKAUFNZU.artnr = Artikel.artnr "
    sSQL = sSQL & " set VERKAUFNZU.Bezeich = Artikel.Bezeich "
    sSQL = sSQL & " , VERKAUFNZU.KVKPR1 = Artikel.KVKPR1 "
    sSQL = sSQL & " , VERKAUFNZU.BESTAND = Artikel.BESTAND "
    sSQL = sSQL & " , VERKAUFNZU.LVKPR = Artikel.VKPR "
    sSQL = sSQL & " , VERKAUFNZU.MWST = Artikel.MWST "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 76
    
    
    sSQL = " Update VERKAUFNZU inner join VERKAUFNZU_DAT2 on VERKAUFNZU.artnr = VERKAUFNZU_DAT2.artnr "
    sSQL = sSQL & " set VERKAUFNZU.VERKAUFSMENGE = VERKAUFNZU_DAT2.VERKAUFSMENGE "
    sSQL = sSQL & " , VERKAUFNZU.VERKAUFSPREIS = VERKAUFNZU_DAT2.VERKAUFSPREIS "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update VERKAUFNZU set NettoVERKAUFSPREIS = (VERKAUFSPREIS * 100)/(100 + " & gdMWStV & ")  "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update VERKAUFNZU set NettoVERKAUFSPREIS = (VERKAUFSPREIS * 100)/(100 + " & gdMWStE & ")  "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update VERKAUFNZU set NettoVERKAUFSPREIS = (VERKAUFSPREIS * 100)/(100 + " & gdMWStO & ")  "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    Label3(0).Caption = VkwertErmittlung
    Label3(0).Visible = True
    Label3(1).Caption = EkwertErmittlung
    Label3(1).Visible = True
    
    
    
    '            cSQL = cSQL & "(ARTNR int "
'            cSQL = cSQL & ", BEZEICH varchar(35) "
'            cSQL = cSQL & ", LEKPR real "
'            cSQL = cSQL & ", LVKPR real "
'            cSQL = cSQL & ", KVKPR1 float "
'            cSQL = cSQL & ", BESTAND int "
'            cSQL = cSQL & ", ZUGANGSMENGE int "
'            cSQL = cSQL & ", VERKAUFSMENGE int "
'            cSQL = cSQL & ", MWST varchar(1) "
'            cSQL = cSQL & ", VERKAUFSPREIS real "
'            cSQL = cSQL & ", NettoVERKAUFSPREIS real "
'            cSQL = cSQL & ", ZUGANGSWERT real "
'            cSQL = cSQL & " ) "
    


    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
 
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleVerkaufNZU"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function VkwertErmittlung() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    anzeige "normal", "VK-Wert wird ermittelt...", Label6
    
    VkwertErmittlung = ""
    
    sSQL = "Select sum(NettoVERKAUFSPREIS)as sVKWERT from VERKAUFNZU"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        VkwertErmittlung = Format$(rsrs!sVKWERT, "#######0.00")
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "", Label6
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VkwertErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function EkwertErmittlung() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    anzeige "normal", "EK-Wert wird ermittelt...", Label6
    
    EkwertErmittlung = ""
    
    sSQL = "Select sum(ZUGANGSWERT)as sEKWERT from VERKAUFNZU"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        EkwertErmittlung = Format$(rsrs!sEKWERT, "#######0.00")
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "", Label6
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EkwertErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub ZeigeArtikelVERKAUFNZU()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    
    MSFlexGrid1.Clear
    
'    Command5(2).Enabled = False
    
    If Not NewTableSuchenDBKombi("VERKAUFNZU", gdBase) Then
        anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label6
        Exit Sub
    Else
        If Datendrin("VERKAUFNZU", gdBase) = False Then
            anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label6
            Exit Sub
        End If
    End If
    
'    Command5(2).Enabled = True
    
    anzeige "normal", "Artikel werden angezeigt, bitte warten...", Label6
    
    Screen.MousePointer = 11
    
    Tabcheck "VERKAUFNZU"
    FormatGridOverTablay "VERKAUFNZU"

    With MSFlexGrid1
        .Redraw = False
'        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) '* 1.8
        Next j
    End With
    
    ermittlespalten
    
    GridFuellen "Select * from VERKAUFNZU order by artnr "
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikelVerkaufNZU"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub GridFuellen(cSQL As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim iRet        As Integer
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim lMax        As Long
    Dim lAnz        As Long
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    picprogress.Visible = True
    With MSFlexGrid1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        lAnz = lMax
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            .Rows = lrow + 1
            .Col = 0
            
            txtStatus.Text = (lrow * 100) / lMax
            
            Select Case lMax
                Case Is > 5000
                
                    j = lAnz Mod 500
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label6
                    End If
                
                Case Is > 1000
                
                    j = lAnz Mod 100
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label6
                    End If
                
                Case Is <= 500
                
                    j = lAnz Mod 50
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label6
                    End If
        
            End Select
    
            lAnz = lAnz - 1
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case UCase(sSpaltenname(i))
                        Case Is = "LEK", "LVK", "KVK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            
                        
                        
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                    End Select
                    
                    If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                        aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                    End If
                    
                End If
            Next i
                                
            rsrs.MoveNext
        Loop
        
        
        
        MSFlexGrid1.Visible = True
        anzeige "normal", lMax & " Artikel", Label6
    
    Else
        MSFlexGrid1.Visible = False
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label6
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
    
    picprogress.Visible = False
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    .Redraw = True
'    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."

    Fehlermeldung1
 
End Sub
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gstab = "VERKAUFNZU"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command20_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 21        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 1        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 0        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 12
            gsHelpstring = "Artikelverkäufe nach Zugang"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
'    positionierenwkl25h
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 2    'vormonat
        
            If Month(DateValue(Now)) = 1 Then
                Text1(2).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(3).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(2).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(3).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(3).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                        
                        If Year(DateValue(Now)) = 2020 Then
                            Text1(3).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        Text1(3).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
                
        Case Is = 5     'ak monat
            Text1(2).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case Is = 6     'gestern
            Text1(2).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        Case Is = 7     'heute
            Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 12 'aktuelles Jahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 14 'Vorjahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(3).Text = Format("31.12." & Year(Now) - 1, "DD.MM.YYYY")
    End Select
    
    Text1(1).Text = Text1(2).Text
    Text1(0).Text = Text1(3).Text
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option3_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 4 Then
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            gF2Prompt.cFeld = "LINR"
                    
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(4).Text = gF2Prompt.cWahl
                End If
            End If
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Verkauf nach Zugang ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "txtStatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
