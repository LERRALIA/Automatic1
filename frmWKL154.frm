VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL154 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " - Bedienerbeteiligung"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   1  'Fenstermitte
   Begin sevCommand3.Command Command1 
      Height          =   345
      Index           =   6
      Left            =   10320
      TabIndex        =   73
      ToolTipText     =   "Leeren"
      Top             =   360
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
      Caption         =   "L"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame5"
      Height          =   1695
      Left            =   -360
      TabIndex        =   72
      Top             =   5520
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   5400
         TabIndex        =   88
         Top             =   3600
         Width           =   6375
         Begin VB.Frame Frame8 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'Kein
            Height          =   615
            Left            =   120
            TabIndex        =   94
            Top             =   1080
            Width           =   2295
            Begin VB.OptionButton Option2 
               Caption         =   "von/bis Zeitraum"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   96
               Top             =   360
               Width           =   1815
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Monat"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   95
               Top             =   120
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   16
            Left            =   4560
            TabIndex        =   93
            Top             =   1320
            Width           =   1695
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
            Caption         =   "Zeigen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Lieferant"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Artikelgruppe"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2880
            Style           =   2  'Dropdown-Liste
            TabIndex        =   89
            Top             =   1320
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker Text2 
            Height          =   375
            Index           =   3
            Left            =   2880
            TabIndex        =   97
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   368574465
            UpDown          =   -1  'True
            CurrentDate     =   38453
         End
         Begin MSComCtl2.DTPicker Text2 
            Height          =   375
            Index           =   2
            Left            =   2880
            TabIndex        =   98
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Format          =   368574465
            UpDown          =   -1  'True
            CurrentDate     =   38453
         End
         Begin sevCommand3.Command Command0 
            Height          =   360
            Index           =   9
            Left            =   4200
            TabIndex        =   103
            ToolTipText     =   "Kalender"
            Top             =   240
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
            Index           =   10
            Left            =   4200
            TabIndex        =   104
            ToolTipText     =   "Kalender"
            Top             =   720
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
         Begin VB.Label Label2 
            BackColor       =   &H80000001&
            Caption         =   "Zusammenfassung nach "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   2655
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   15
         Left            =   9960
         TabIndex        =   86
         Top             =   3120
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   14
         Left            =   9960
         TabIndex        =   84
         Top             =   2640
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   13
         Left            =   9960
         TabIndex        =   82
         Top             =   2160
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   12
         Left            =   9960
         TabIndex        =   80
         Top             =   1680
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   11
         Left            =   9960
         TabIndex        =   78
         Top             =   1200
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   10
         Left            =   9960
         TabIndex        =   75
         Top             =   720
         Width           =   1695
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   9
         Left            =   9960
         TabIndex        =   74
         Top             =   7080
         Width           =   1695
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   345
         Index           =   4
         Left            =   10800
         TabIndex        =   100
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
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
         Picture         =   "frmWKL154.frx":0000
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label lblanze 
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   6960
         Width           =   5295
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Provisionen rabattierfähiger Artikel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   87
         Top             =   3240
         Width           =   8055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Provisionen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   85
         Top             =   2760
         Width           =   8055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Ertrag pro Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   83
         Top             =   2280
         Width           =   8055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Umsatz pro Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   81
         Top             =   1800
         Width           =   8055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Entwicklung Verkauf pro Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   79
         Top             =   1320
         Width           =   8055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "vorgefertigte Auswertungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Verkauf pro Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   8055
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   8
      Left            =   9960
      TabIndex        =   71
      Top             =   7080
      Width           =   1695
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
      Caption         =   "Listen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   70
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Euro pro Stück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
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
      Height          =   330
      Index           =   9
      Left            =   6480
      MaxLength       =   5
      TabIndex        =   69
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   68
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
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   63
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
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
      Caption         =   "vom Rohertrag"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   62
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
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
      Caption         =   "vom Umsatz"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
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
      Height          =   330
      Index           =   8
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   61
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   54
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   40
      Top             =   7560
      Width           =   1695
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11775
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
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
         Height          =   330
         Index           =   6
         Left            =   5520
         MaxLength       =   13
         TabIndex        =   59
         Top             =   120
         Width           =   1575
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   58
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
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
         Height          =   330
         Index           =   1
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   56
         Top             =   960
         Width           =   855
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   1
         Left            =   4320
         TabIndex        =   55
         Top             =   600
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
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   53
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   0
         Left            =   9480
         TabIndex        =   51
         Top             =   600
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
         Height          =   330
         Index           =   0
         Left            =   9000
         MaxLength       =   6
         TabIndex        =   50
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "nur umsatzrelevante Artikelverkäufe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   49
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   2055
         Left            =   6120
         TabIndex        =   42
         Top             =   3000
         Width           =   3135
         Begin VB.OptionButton Option3 
            Caption         =   "Vorjahr Zeitraum"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            Caption         =   "aktuelles Jahr"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Vorjahr"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "aktueller Monat"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Vormonat"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Datum Voreinstellung"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ohne Gutscheine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   3120
         TabIndex        =   24
         Top             =   2280
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "nur Warengruppen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Tag             =   "Menge"
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ausschließen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Tag             =   "Preis"
            Top             =   720
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "einschließen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Tag             =   "Ertrag"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Warengruppen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Bedienername"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   66
            Tag             =   "Menge"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Bedienernummer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   41
            Tag             =   "Menge"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rohertrag"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Tag             =   "Ertrag"
            Top             =   360
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Umsatz"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Tag             =   "Preis"
            Top             =   720
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Stückzahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Tag             =   "Menge"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "sortiert nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   2175
         End
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
         Height          =   330
         Index           =   2
         Left            =   4680
         MaxLength       =   6
         TabIndex        =   18
         Top             =   960
         Width           =   1095
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
         Height          =   330
         Index           =   3
         Left            =   6720
         MaxLength       =   13
         TabIndex        =   17
         Top             =   960
         Width           =   1455
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
         Height          =   330
         Index           =   4
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   5
         Left            =   5400
         TabIndex        =   15
         Top             =   600
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
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   4
         Left            =   8640
         TabIndex        =   14
         Top             =   600
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
         Height          =   330
         Index           =   5
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   2
         Left            =   6360
         TabIndex        =   12
         Top             =   600
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
      Begin sevCommand3.Command Command0 
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   11
         Top             =   240
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
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
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
         Height          =   330
         Index           =   7
         Left            =   7920
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   6
         Left            =   9480
         TabIndex        =   8
         Top             =   120
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
      Begin MSComCtl2.DTPicker Text2 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130744321
         UpDown          =   -1  'True
         CurrentDate     =   38453
      End
      Begin MSComCtl2.DTPicker Text2 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   32
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130744321
         UpDown          =   -1  'True
         CurrentDate     =   38453
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   7
         Left            =   1440
         TabIndex        =   101
         ToolTipText     =   "Kalender"
         Top             =   960
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
         Index           =   8
         Left            =   3360
         TabIndex        =   102
         ToolTipText     =   "Kalender"
         Top             =   960
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
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000001&
         Caption         =   "ArtikelNr / EAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Bed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   57
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PGN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   9120
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Nr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Bestell-Nr."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "AGN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   37
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Linie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "alle Farben"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   32
         Left            =   1080
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000001&
         Caption         =   "Marke"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   9
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
   End
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
      Left            =   6360
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393216
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin sevCommand3.Command cmdEnd 
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   8040
      Width           =   1695
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
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
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
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Prozent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   14
      Left            =   7440
      TabIndex        =   67
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "vom Umsatz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   13
      Left            =   7440
      TabIndex        =   65
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Prozent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   64
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   7920
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bedienerauswertung"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmWKL154"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL154
        
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Text1_KeyUp 0, vbKeyF2, 0
        Case Is = 1
            Text1_KeyUp 1, vbKeyF2, 0
        Case Is = 4
            Text1_KeyUp 4, vbKeyF2, 0
        Case Is = 5
            Text1_KeyUp 2, vbKeyF2, 0
        Case Is = 2
            Text1_KeyUp 5, vbKeyF2, 0
        Case 3
            gsBackcolor = Label4(32).BackColor
            gsForecolor = Label4(32).ForeColor
            gsArtikelFarbe = Label4(32).Tag
            
            frmWKL49.Show 1
            
            Label4(32).BackColor = gsBackcolor
            Label4(32).ForeColor = gsForecolor
            Label4(32).Tag = gsArtikelFarbe
            
            If gsArtikelFarbe <> "" Then
                Label4(32).Caption = "Farbauswahl"
            Else
                Label4(32).Caption = "alle Farben"
            End If
        Case Is = 7
            Text2(0).Value = Format(Datumschreiben11a(5600, Text2(0).Left), "DD.MM.YY")
        Case Is = 8
            Text2(1).Value = Format(Datumschreiben11a(5600, Text2(1).Left), "DD.MM.YY")
        Case Is = 9
            Text2(2).Value = Format(Datumschreiben11a(5600, Text2(0).Left), "DD.MM.YY")
        Case Is = 10
            Text2(3).Value = Format(Datumschreiben11a(5600, Text2(1).Left), "DD.MM.YY")
            'fertig
        Case 6
            Text1_KeyUp 7, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub WKLatPositionieren()
    On Error GoTo LOKAL_ERROR
    
    MSHFLEX1.Height = 5895
    MSHFLEX1.Left = 240
    MSHFLEX1.Top = 960
    MSHFLEX1.Width = 11415
    
    With Frame1
        .Top = 960
        .Left = 0
        .Height = 5775
        .Width = 11895
        .BorderStyle = 0
    End With
    
    With Frame5
        .Top = 960
        .Left = 0
        .Height = 7575
        .Width = 11895
        .BorderStyle = 0
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKLatPositionieren"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sOrder  As String
    Dim sSQL    As String
    Dim i       As Integer
    Dim iMonat  As Integer
    Dim iJahr   As Integer
    Dim cVon    As String
    Dim cBis    As String
    Dim iRet    As Integer
    
    Select Case Index
        Case Is = 0     '** ermitteln *
        
            Tabcheck "BB"
            FormatGridOverTablay "BB"

            Dim j As Integer
            
            With MSHFLEX1
                .Redraw = False
                .Visible = False
                .Clear
                .Rows = 25
                .Cols = byAnzahlSpalten
                .FixedCols = 0
                .FixedRows = 1
                .Row = 0
                
                For j = 0 To byAnzahlSpalten - 1
                    .Col = j
                    .Text = sSpaltenname(j)
                    aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
                Next j
            End With
   
            Me.Refresh
            ermitteln
            
            If Option1(0).Value = True Then
                sOrder = " Order by ertrag desc"
            ElseIf Option1(1).Value = True Then
                sOrder = " Order by preis desc" 'Umsatz
            ElseIf Option1(2).Value = True Then
                sOrder = " Order by menge desc" 'Menge
            ElseIf Option1(3).Value = True Then
                sOrder = " Order by bednu asc" 'Bedienernummer
            ElseIf Option1(9).Value = True Then
                sOrder = " Order by bedname" 'Bedienername
            End If
            
            GridFuellen "Select * from BEDTOPI " & sOrder
            Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
            
        Case 1 'Drucken
        
            iRet = MsgBox("Als Artikelansicht ausdrucken?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                Drucke138_Artikel
            Else
                Drucke138
            End If
            
        
        Case 2
            Frame1.Visible = True
            MSHFLEX1.Visible = False
            
            Text1(9).Visible = False
            Text1(8).Visible = False
            Label1(11).Visible = False
            Command1(3).Visible = False
            Command1(5).Visible = False
            Command1(1).Visible = False
            Command1(2).Visible = False
            Command1(7).Visible = False
            
            Command1(6).Visible = True
        Case 3
            Text1(9).Text = ""
            Label1(13).Caption = "Umsatz"
            If Text1(8).Text <> "" Then
                If IsNumeric(Text1(8).Text) Then
                    sSQL = "Update BEDTOPI SET PROV = Preis * '" & CDbl(Text1(8).Text) & "' / 100 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from BEDTOPI " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        Case Is = 4
            gsZSpalte = "Artnr"
            gstab = "BB"
            frmWKL36.Show 1
            'fertig
        Case 5
            Text1(9).Text = ""
            Label1(13).Caption = "Rohertrag"
        
            If Text1(8).Text <> "" Then
                If IsNumeric(Text1(8).Text) Then
                    sSQL = "Update BEDTOPI SET PROV = Ertrag * '" & CDbl(Text1(8).Text) & "' / 100 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from BEDTOPI " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        Case 6
            List1.Clear
            List2.Clear
            List3.Clear
            List4.Clear
            List5.Clear
            List1.Visible = False
            List2.Visible = False
            List3.Visible = False
            List4.Visible = False
            List5.Visible = False
            
            For i = 0 To 8
                Text1(i).Text = ""
            Next i
        Case 7
            Text1(8).Text = ""
            Label1(13).Caption = "Euro pro Stück"
            If Text1(9).Text <> "" Then
                If IsNumeric(Text1(9).Text) Then
                    sSQL = "Update BEDTOPI SET PROV = Menge * '" & CDbl(Text1(9).Text) & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from BEDTOPI " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        Case 8
            Frame5.Visible = True
        Case 9
            Frame5.Visible = False
        Case 10
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedKuCut txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 11 'KUCUT Bediener Entwicklung
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedKuCutDEVELo txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 12
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedKuCut1 txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 13
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedKuCut2 txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 14
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedProvision txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 15
            schreibeProtokollProgrammablauf " löst Liste aus    " & Label2(Index).Caption
            BestBedProvisionRab txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & Label2(Index).Caption
        Case Is = 16
            If Option2(5).Value = True Then
                If Option2(4).Value = True Then
                    iMonat = CByte(Mid$(Combo3.Text, 1, InStr(1, Combo3.Text, "/") - 1))
                    iJahr = CInt(Right(Combo3.Text, 4))
                
                    BedienernachLINR iMonat, iJahr
                ElseIf Option2(3).Value = True Then
                    iMonat = CByte(Mid$(Combo3.Text, 1, InStr(1, Combo3.Text, "/") - 1))
                    iJahr = CInt(Right(Combo3.Text, 4))
                            
                    BedienernachAGN iMonat, iJahr
                End If
            ElseIf Option2(6).Value = True Then 'neu mit Zeitraum
                If Option2(4).Value = True Then
                    iMonat = CByte(Mid$(Combo3.Text, 1, InStr(1, Combo3.Text, "/") - 1))
                    iJahr = CInt(Right(Combo3.Text, 4))
                
                    BedienernachLINR iMonat, iJahr
                ElseIf Option2(3).Value = True Then
                    cVon = Text2(2).Value
                    cBis = Text2(3).Value
                            
                    BedienernachAGN_ZR cVon, cBis
                End If
            End If
    End Select
   
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Drucke138()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim cProv       As String
    Dim cproz       As String
    Dim sOrder      As String
    Dim sSorti      As String
    Dim dGumsatz    As Double
    Dim i           As Integer
    
    Dim sEAN        As String
    Dim slibesnr    As String
    Dim sMARKE      As String
    Dim sLiefBez    As String
    Dim lLinr       As Long
    Dim sAGN        As String
    Dim sPGN        As String
    Dim sLPZ        As String
    Dim sFilauswahl As String
    Dim sBedbez     As String
    Dim lbednu      As Long
    
    
    Screen.MousePointer = 11
    
    cDatVon = Text2(0).Value
    cDatBis = Text2(1).Value
    dGumsatz = CDbl(Label1(14).Caption)
    
    If Text1(8).Text <> "" Then
        cProv = Label1(13).Caption
        cproz = Text1(8).Text & " % vom"
    End If
    
    If Text1(9).Text <> "" Then
'        cProv = Label1(13).Caption
        cproz = Text1(9).Text & " € pro Stück"
    End If
    
    sMARKE = Text1(7).Text
    sEAN = Text1(6).Text
    slibesnr = Text1(3).Text
'    sFilauswahl = cboFil.Text
    
    
    
'    If Text1(2).Text <> "" Then
'        lLinr = Val(Text1(2).Text)
'    Else
'        lLinr = 0
'    End If
    
'    sLiefBez = ermLiefBez(lLinr)
    
    
    'Lieferant
    If List5.ListCount > 0 Then
        sLiefBez = List5.list(0) & vbCrLf
        For i = 1 To List5.ListCount - 1
            sLiefBez = sLiefBez & List5.list(i) & vbCrLf
        Next i
    Else
        'Lieferant
        If Text1(2).Text <> "" Then
            lLinr = Val(Text1(2).Text)
        Else
            lLinr = 0
        End If
        
        sLiefBez = ermLiefBez(lLinr)
    
    End If
    
    
    
    
    'Linie
    If List3.ListCount > 0 Then
        sLPZ = Left(List3.list(0), 20) & vbCrLf
        For i = 1 To List3.ListCount - 1
            sLPZ = sLPZ & Left(List3.list(i), 20) & vbCrLf
        Next i
    Else
        'Linie
        sLPZ = Trim$(Text1(5).Text)
    End If
    
    'AGN
    If List1.ListCount > 0 Then
        sAGN = List1.list(0) & vbCrLf
        For i = 1 To List1.ListCount - 1
            sAGN = sAGN & List1.list(i) & vbCrLf
        Next i
    Else
        'agn
        sAGN = Trim$(Text1(4).Text)
    End If
    
    'PGN
    If List2.ListCount > 0 Then
        sPGN = List2.list(0) & vbCrLf
        For i = 1 To List2.ListCount - 1
            sPGN = sPGN & List2.list(i) & vbCrLf
        Next i
    Else
        'sPGN
        sPGN = Trim$(Text1(0).Text)
    End If

    'Bediener
    If List4.ListCount > 0 Then
        sBedbez = List4.list(0) & vbCrLf
        For i = 1 To List4.ListCount - 1
            sBedbez = sBedbez & List4.list(i) & vbCrLf
        Next i
    Else
        'Bediener
        If Text1(1).Text <> "" Then
            lbednu = Val(Text1(1).Text)
        Else
            lbednu = 0
        End If
        sBedbez = ermBEDbez(lbednu)
    End If
    

    Screen.MousePointer = 11
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
    
    If Option1(0).Value = True Then
        sOrder = " Order by ertrag desc"
        sSorti = "sortiert nach Rohertrag"
    ElseIf Option1(1).Value = True Then
        sOrder = " Order by preis desc" 'Umsatz
        sSorti = "sortiert nach Umsatz"
    ElseIf Option1(2).Value = True Then
        sOrder = " Order by menge desc" 'Menge
        sSorti = "sortiert nach Stückzahlen"
    ElseIf Option1(3).Value = True Then
        sOrder = " Order by bednu asc" 'Bediener
        sSorti = "sortiert nach Bedienernummer"
    ElseIf Option1(9).Value = True Then
        sOrder = " Order by bedname " 'Bedienername
        sSorti = "sortiert nach Bedienername"
    End If
            
    loeschNEW "a138tmp", gdBase
    
    cSQL = "Select * into a138tmp from BEDTOPI "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "BEDTOPI", gdBase
    
    cSQL = "Select * into BEDTOPI from a138tmp " & sOrder
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "a138tmp", gdBase
    
    loeschNEW "Kopf138", gdBase
    CreateTableT2 "KOPF138", gdBase
    
    cSQL = "Insert into KOPF138 (DATVON,DATBIS,Prov,Proz,Sortierung,Gumsatz"
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MARKE "
    cSQL = cSQL & ", Liefbez "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", AGN "
    cSQL = cSQL & ", PGN "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", FILAUSWAHL "
    cSQL = cSQL & ", BEDBEZ "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & " ) values ("
    cSQL = cSQL & " '" & cDatVon & "'  "
    cSQL = cSQL & ", '" & cDatBis & "'  "
    cSQL = cSQL & ", '" & cProv & "'  "
    cSQL = cSQL & ", '" & cproz & "'  "
    cSQL = cSQL & ", '" & sSorti & "'  "
    cSQL = cSQL & ", '" & dGumsatz & "'  "
    
    cSQL = cSQL & ", '" & sEAN & "'  "
    cSQL = cSQL & ", '" & slibesnr & "'  "
    cSQL = cSQL & ", '" & sMARKE & "'  "
    cSQL = cSQL & ", '" & sLiefBez & "'  "
    cSQL = cSQL & ", " & lLinr & "  "
    cSQL = cSQL & ", '" & sAGN & "'  "
    cSQL = cSQL & ", '" & sPGN & "'  "
    cSQL = cSQL & ", '" & sLPZ & "'  "
    cSQL = cSQL & ", '" & sFilauswahl & "'  "
    cSQL = cSQL & ", '" & sBedbez & "'  "
    cSQL = cSQL & ", " & lbednu & "  "
    
    
    cSQL = cSQL & "  ) "
    gdBase.Execute cSQL, dbFailOnError
    
    reportbildschirm "", "aZEN138a"
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke138"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Drucke138_Artikel()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim cProv       As String
    Dim cproz       As String
    Dim sOrder      As String
    Dim sSorti      As String
    Dim dGumsatz    As Double
    Dim i           As Integer
    
    Dim sEAN        As String
    Dim slibesnr    As String
    Dim sMARKE      As String
    Dim sLiefBez    As String
    Dim lLinr       As Long
    Dim sAGN        As String
    Dim sPGN        As String
    Dim sLPZ        As String
    Dim sFilauswahl As String
    Dim sBedbez     As String
    Dim lbednu      As Long
    
    Screen.MousePointer = 11
    
    cDatVon = Text2(0).Value
    cDatBis = Text2(1).Value
    dGumsatz = CDbl(Label1(14).Caption)
    
    If Text1(8).Text <> "" Then
        cProv = Label1(13).Caption
        cproz = Text1(8).Text & " % vom"
    End If
    
    If Text1(9).Text <> "" Then

        cproz = Text1(9).Text & " € pro Stück"
    End If
    
    sMARKE = Text1(7).Text
    sEAN = Text1(6).Text
    slibesnr = Text1(3).Text

    'Lieferant
    If List5.ListCount > 0 Then
        sLiefBez = List5.list(0) & vbCrLf
        For i = 1 To List5.ListCount - 1
            sLiefBez = sLiefBez & List5.list(i) & vbCrLf
        Next i
    Else
        'Lieferant
        If Text1(2).Text <> "" Then
            lLinr = Val(Text1(2).Text)
        Else
            lLinr = 0
        End If
        
        sLiefBez = ermLiefBez(lLinr)
    
    End If
    
    'Linie
    If List3.ListCount > 0 Then
        sLPZ = Left(List3.list(0), 20) & vbCrLf
        For i = 1 To List3.ListCount - 1
            sLPZ = sLPZ & Left(List3.list(i), 20) & vbCrLf
        Next i
    Else
        'Linie
        sLPZ = Trim$(Text1(5).Text)
    End If
    
    'AGN
    If List1.ListCount > 0 Then
        sAGN = List1.list(0) & vbCrLf
        For i = 1 To List1.ListCount - 1
            sAGN = sAGN & List1.list(i) & vbCrLf
        Next i
    Else
        'agn
        sAGN = Trim$(Text1(4).Text)
    End If
    
    'PGN
    If List2.ListCount > 0 Then
        sPGN = List2.list(0) & vbCrLf
        For i = 1 To List2.ListCount - 1
            sPGN = sPGN & List2.list(i) & vbCrLf
        Next i
    Else
        'sPGN
        sPGN = Trim$(Text1(0).Text)
    End If

    'Bediener
    If List4.ListCount > 0 Then
        sBedbez = List4.list(0) & vbCrLf
        For i = 1 To List4.ListCount - 1
            sBedbez = sBedbez & List4.list(i) & vbCrLf
        Next i
    Else
        'Bediener
        If Text1(1).Text <> "" Then
            lbednu = Val(Text1(1).Text)
        Else
            lbednu = 0
        End If
        sBedbez = ermBEDbez(lbednu)
    End If
    

    Screen.MousePointer = 11
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
    
    If Option1(0).Value = True Then
        sOrder = " Order by ertrag desc"
        sSorti = "sortiert nach Rohertrag"
    ElseIf Option1(1).Value = True Then
        sOrder = " Order by preis desc" 'Umsatz
        sSorti = "sortiert nach Umsatz"
    ElseIf Option1(2).Value = True Then
        sOrder = " Order by menge desc" 'Menge
        sSorti = "sortiert nach Stückzahlen"
    ElseIf Option1(3).Value = True Then
        sOrder = " Order by bednu asc" 'Bediener
        sSorti = "sortiert nach Bedienernummer"
    ElseIf Option1(9).Value = True Then
        sOrder = " Order by bedname " 'Bedienername
        sSorti = "sortiert nach Bedienername"
    End If
            

    
    
    
    loeschNEW "Kopf138", gdBase
    CreateTableT2 "KOPF138", gdBase
    
    cSQL = "Insert into KOPF138 (DATVON,DATBIS,Prov,Proz,Sortierung,Gumsatz"
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MARKE "
    cSQL = cSQL & ", Liefbez "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", AGN "
    cSQL = cSQL & ", PGN "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", FILAUSWAHL "
    cSQL = cSQL & ", BEDBEZ "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & " ) values ("
    cSQL = cSQL & " '" & cDatVon & "'  "
    cSQL = cSQL & ", '" & cDatBis & "'  "
    cSQL = cSQL & ", '" & cProv & "'  "
    cSQL = cSQL & ", '" & cproz & "'  "
    cSQL = cSQL & ", '" & sSorti & "'  "
    cSQL = cSQL & ", '" & dGumsatz & "'  "
    
    cSQL = cSQL & ", '" & sEAN & "'  "
    cSQL = cSQL & ", '" & slibesnr & "'  "
    cSQL = cSQL & ", '" & sMARKE & "'  "
    cSQL = cSQL & ", '" & sLiefBez & "'  "
    cSQL = cSQL & ", " & lLinr & "  "
    cSQL = cSQL & ", '" & sAGN & "'  "
    cSQL = cSQL & ", '" & sPGN & "'  "
    cSQL = cSQL & ", '" & sLPZ & "'  "
    cSQL = cSQL & ", '" & sFilauswahl & "'  "
    cSQL = cSQL & ", '" & sBedbez & "'  "
    cSQL = cSQL & ", " & lbednu & "  "
    
    
    cSQL = cSQL & "  ) "
    gdBase.Execute cSQL, dbFailOnError
    
'    reportbildschirm "", "aZEN138a"
    reportbildschirm "", "aZEN138b"
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke138_Artikel"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
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
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSHFLEX1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
    
        anzeige "normal", "Es werden " & lMax & " Bediener angezeigt...", lblanzeige
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
                        Case Is = "Umsatz", "NS", "Rohertrag", "K. Schnitt", "Provision"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                        Case Is = "Umsatz EK", "Ums pro Stück", "Ums pro Kunde"
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
        
        Frame1.Visible = False
        
        Text1(9).Visible = True
        Text1(8).Visible = True
        Label1(11).Visible = True
        Command1(3).Visible = True
        Command1(5).Visible = True
        Command1(1).Visible = True
        Command1(2).Visible = True
        Command1(7).Visible = True
        
        Command1(6).Visible = False
        .Visible = True
        anzeige "normal", "Es wurden " & lMax & " Bediener ermittelt.", lblanzeige
    Else
        Frame1.Visible = True
        anzeige "rot", "Es wurden keine Daten ermittelt.", lblanzeige
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
'    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermitteln()
On Error GoTo LOKAL_ERROR

    Dim cVon            As String
    Dim cBis            As String
    Dim lVon            As Long
    Dim lBis            As Long
    Dim corder          As String
    Dim i               As Integer
    Dim bAnd            As Boolean
    Dim ctmp            As String
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsb             As Recordset
    Dim iAnzahlKunden   As Long

    'vorbereitung
    If Text2(0).Value <> "" Then
        cVon = Text2(0).Value
    Else
        cVon = DateValue(Now) - 30
        Text2(0).Value = DateValue(Now) - 30
    End If
    
    If Text2(1).Value <> "" Then
        cBis = Text2(1).Value
    Else
        cBis = DateValue(Now)
        Text2(1).Value = DateValue(Now)
    End If
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))

    'Vorbereitung ende
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", lblanzeige
    
    sSQL = "Select "
    sSQL = sSQL & " Sum(preis) as Maxi "
    sSQL = sSQL & " from Kassjour A "
    sSQL = sSQL & " where A.adate between  " & cVon & " And " & cBis
'    sSQL = sSQL & " and UMS_OK = 'J' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!maxi) Then
            Label1(14).Caption = Format(rsrs!maxi, "########0.00")
        End If
    End If
    rsrs.Close
        
    loeschNEW "BEDB", gdBase
    CreateTableT2 "BEDB", gdBase
    
    sSQL = "Insert into BEDB Select "
    sSQL = sSQL & "  A.preis "
    sSQL = sSQL & " ,A.menge "
    sSQL = sSQL & " , A.artnr "
    sSQL = sSQL & " , A.ekpr "
    sSQL = sSQL & " , A.mwst "
    sSQL = sSQL & " , A.BEDIENER as bednu "
    sSQL = sSQL & " , A.BELEGNR  "
    sSQL = sSQL & " , A.adate  "
    sSQL = sSQL & " from Kassjour A "
    
    'LiefBestNr
    ctmp = Trim$(Text1(3).Text)
    If ctmp <> "" Then
        sSQL = sSQL & " inner join Artlief B on A.Artnr = B.Artnr "
    End If
    
    sSQL = sSQL & " where A.adate between  " & cVon & " And " & cBis
    
    bAnd = True
     
    If Check2.Value = vbChecked Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & "  A.UMS_OK = 'J' "
    End If
    
    
    
    'Linr
    If List5.Visible = True And List5.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( a.linr=" & Trim$(Left$(List5.list(0), InStr(1, List5.list(0), " ")))
        For i = 1 To List5.ListCount - 1
            sSQL = sSQL & " or a.linr=" & Trim$(Left$(List5.list(i), InStr(1, List5.list(i), " ")))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'linr
        ctmp = Trim$(Text1(2).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.linr = " & ctmp & " "
            bAnd = True
        End If
    End If
    
    
    
    'ArtNr oder EAN
    ctmp = Trim$(Text1(6).Text)
    If ctmp <> "" Then
    
        If bAnd Then
            sSQL = sSQL & " and "
        End If
       
        If Len(ctmp) <= 6 Then
            'KISS-ArtNr
            sSQL = sSQL & " A.ARTNR = " & ctmp & " "
            bAnd = True
            
        ElseIf Len(ctmp) = 8 Then
            'KISS-ArtNr als Barcode oder echter EAN-8
            If Left$(ctmp, 1) = "2" Or Left$(ctmp, 1) = "0" Then
                ctmp = Mid$(ctmp, 2, 6)
                sSQL = sSQL & " A.ARTNR = " & ctmp & " "
                bAnd = True
            Else
                sSQL = sSQL & " A.EAN = '" & ctmp & "' "
                bAnd = True
            End If
        Else
            'Irgendwas anderes für die EAN-Felder
            sSQL = sSQL & " A.EAN = '" & ctmp & "' "
            bAnd = True
        End If
    End If
    
    

    'Linie
    If List3.Visible = True And List3.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( lpz=" & Mid$(List3.list(0), 1, InStr(1, List3.list(0), " "))
        For i = 1 To List3.ListCount - 1
            sSQL = sSQL & " or lpz=" & Mid$(List3.list(i), 1, InStr(1, List3.list(i), " "))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'Linie
        ctmp = Trim$(Text1(5).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.LPZ = " & ctmp & " "
            bAnd = True
        End If
    End If

    'Marke
    ctmp = Trim$(Text1(7).Text)
    If ctmp <> "" Then
        If LoeseMarkenInArtnr(ctmp) Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & " A.artnr in(Select artnr from MY" & srechnertab & ")"
            bAnd = True
        End If
    End If

    'LiefBestNr
    ctmp = Trim$(Text1(3).Text)
    If ctmp <> "" Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & " B.LIBESNR like '" & ctmp & "*' "
        bAnd = True
    End If

    'AGN
    If List1.Visible = True And List1.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( agn=" & Trim$(Left$(List1.list(0), InStr(1, List1.list(0), " ")))
        For i = 1 To List1.ListCount - 1
            sSQL = sSQL & " or agn=" & Trim$(Left$(List1.list(i), InStr(1, List1.list(i), " ")))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'agn
        ctmp = Trim$(Text1(4).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.AGN = " & ctmp & " "
            bAnd = True
        End If
    End If
    
    'Bediener
    If List4.Visible = True And List4.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( Bediener=" & Trim$(Left$(List4.list(0), InStr(1, List4.list(0), " ")))
        For i = 1 To List4.ListCount - 1
            sSQL = sSQL & " or Bediener=" & Trim$(Left$(List4.list(i), InStr(1, List4.list(i), " ")))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'Bediener
        ctmp = Trim$(Text1(1).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.Bediener = " & ctmp & " "
            bAnd = True
        End If
    End If
'    MsgBox sSQL
    gdBase.Execute sSQL, dbFailOnError
    
    'Farben
    ctmp = Trim$(Label4(32).Tag)
    If ctmp <> "" Then
        sSQL = "delete BEDB.* from BEDB inner join ARTIKEL on ARTIKEL.artnr = BEDB.artnr "
        sSQL = sSQL & " where artikel.awm <> '" & ctmp & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    'PGN
    If List2.Visible = True And List2.ListCount > 0 Then
        For i = 1 To List2.ListCount - 1
            sSQL = "delete BEDB.* from BEDB inner join ARTIKEL on ARTIKEL.artnr = BEDB.artnr "
            sSQL = sSQL & " where artikel.pgn <> " & Trim$(Left$(List2.list(i), 2))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    Else
        'Pgn
        ctmp = Trim$(Text1(0).Text)
        If ctmp <> "" Then
            sSQL = "delete BEDB.* from BEDB inner join ARTIKEL on ARTIKEL.artnr = BEDB.artnr "
            sSQL = sSQL & " where artikel.pgn <> " & ctmp
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    'Warengruppen
    If Option2(0).Value = True Then
        'einschließen
    
        
    ElseIf Option2(1).Value = True Then
        'ausschließen
        
        If checkwarengru Then
            sSQL = "Delete from BEDB where artnr in(select artnr from warengru) "
            gdBase.Execute sSQL, dbFailOnError
        End If
    ElseIf Option2(2).Value = True Then
        'nur Warengruppen
        
        If checkwarengru Then
            sSQL = "Delete from BEDB where artnr not in(select artnr from warengru) "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    'End Warengruppen
    
    
    'Gutscheine
    If Check1.Value = vbChecked Then
        sSQL = "delete from BEDB where artnr = 666666 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    'End Gutscheine
    sSQL = " Create index  MWST on BEDB(MWST) "
    gdBase.Execute sSQL, dbFailOnError

    
    sSQL = "Update BEDB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStV & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStV & "))* 100)"
    
    sSQL = sSQL & " where MWST = 'V' "
    sSQL = sSQL & " and PREIS <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStE & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStE & "))* 100)"
    
    sSQL = sSQL & " where MWST = 'E' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStO & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStO & "))* 100)"
    
    sSQL = sSQL & " where MWST = 'O' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update BEDB set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDB set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDB set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "BEDUMSATZ", gdBase
    
    sSQL = " Create index  bednu on BEDB(bednu) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select bednu, sum(rertrag) as mrertrag,sum(preis) as mpreis,sum(menge) as mmenge ,sum(menge* ekpr) as mekpr,avg(ens1) as ens into BEDUMSATZ "
    sSQL = sSQL & " from BEDB group by bednu "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    'Vorbereitung Artikelansicht
    loeschNEW "BEDB_ARTIKEL", gdBase
    
    sSQL = "Select * into BEDB_ARTIKEL from BEDB "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = " Alter table BEDB_ARTIKEL add BEZEICH Text(35)  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Alter table BEDB_ARTIKEL add BEDNAME Text(35)  "
    gdBase.Execute sSQL, dbFailOnError

    
    sSQL = "Update BEDB_ARTIKEL inner join Artikel on BEDB_ARTIKEL.artnr = Artikel.artnr "
    sSQL = sSQL & " set BEDB_ARTIKEL.BEZEICH = Artikel.BEZEICH"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDB_ARTIKEL inner join Bedname on BEDB_ARTIKEL.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET BEDB_ARTIKEL.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'Vorbereitung Artikelansicht Ende
    
    
    
    
    
    

    loeschNEW "BEDTOPI", gdBase
    CreateTableT2 "BEDTOPI", gdBase
    
    sSQL = "Insert into BEDTOPI SELECT  bednu, mrertrag as ertrag, mmenge as menge , mekpr as umsek"
    sSQL = sSQL & " , mpreis as preis , ens  "
    sSQL = sSQL & " from BEDUMSATZ "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Create index  BELEGNR on BEDB(BELEGNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Create index  adate on BEDB(adate) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Select bednu from BEDUMSATZ "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN "
                sSQL = sSQL & " from BEDB "
                sSQL = sSQL & " where bednu = " & rsrs!BEDNU & " "
            
                Set rsb = gdBase.OpenRecordset(sSQL)
            
                If Not rsb.EOF Then
                    iAnzahlKunden = rsb.RecordCount
                Else
                    iAnzahlKunden = 0
                End If
                rsb.Close
                
                sSQL = "Update BEDTOPI set anzkunden = " & iAnzahlKunden
                sSQL = sSQL & " where BEDTOPI.BEDNU = " & rsrs!BEDNU
                gdBase.Execute sSQL, dbFailOnError
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "Update BEDTOPI inner join Bedname on BEDTOPI.bednu = Bedname.BEDNU "
    sSQL = sSQL & " SET BEDTOPI.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDTOPI SET KUSCHNI = Menge/anzkunden "
    sSQL = sSQL & " where BEDTOPI.anzkunden <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDTOPI SET UMSST = Preis/Menge "
    sSQL = sSQL & " where BEDTOPI.Menge <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDTOPI SET UMSKU = PREIS/anzkunden "
    sSQL = sSQL & " where BEDTOPI.anzkunden <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDTOPI SET PROV = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "BEDB", gdBase

    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0
   
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermitteln"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
    
    Case 11
        gsHelpstring = "Bedienerbeteiligung"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
       
    WKLatPositionieren
    Skalieren Me, True, True: Schrift Me
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Screen.MousePointer = 11
    Option3_Click 4
    If NewTableSuchenDBKombi("C138E", gdApp) Then
        If SpalteInTabellegefundenNEW("C138E", "iOpt1", gdApp) Then
            voreinstellungladen

        End If
    End If
    
    If Month(DateValue(Now)) = 1 Then
        Text2(2).Value = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
        Text2(3).Value = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
    Else
        Text2(2).Value = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
        Select Case Month(DateValue(Now)) - 1
            Case 1, 3, 5, 7, 8, 10, 12
                Text2(3).Value = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
            Case 2
                Text2(3).Value = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
            Case Else
                Text2(3).Value = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
        End Select
    End If
    
    fuellecombo
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub fuellecombo()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim iMonat As Integer
    Dim iJahr As Integer
    
    iMonat = Month(Now)
    iJahr = Year(Now)
    
    With Combo3
        .Clear
        For i = 1 To 12
        
            If iMonat = 1 Then
                iMonat = 12
                iJahr = iJahr - 1
            Else
                iMonat = iMonat - 1
                iJahr = iJahr
            End If
            
            .AddItem iMonat & "/" & iJahr
            If .Text = "" Then
                .Text = iMonat & "/" & iJahr
            End If
            
        Next i
        
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BedienernachLINR(imon As Integer, iJahr As Integer)
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    Dim sSQL As String

    loeschNEW "BEDZLINR", gdBase
    CreateTableT2 "BEDZLINR", gdBase

    sSQL = "Insert into BEDZLINR Select KASSJOUR.Bediener, BEDNAME.BEDNAME, LISRT.LIEFBEZ  ,KASSJOUR.LINR "
    sSQL = sSQL & " , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis "
    sSQL = sSQL & " , 0 as NUMSATZ "
    sSQL = sSQL & " from Kassjour, BEDNAME, LISRT"
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and LISRT.LINR = kassjour.LINR "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME ,LISRT.LIEFBEZ  ,KASSJOUR.LINR "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "T", gdBase
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStV & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.LINR "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, LISRT "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'V' "
    sSQL = sSQL & " and LISRT.LINR = kassjour.LINR "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.LINR  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR inner join t on BEDZLINR.BEDIENER = t.Bediener and BEDZLINR.LINR = T.LINR "
    sSQL = sSQL & " set BEDZLINR.NUMSATZ = t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "T", gdBase
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStE & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.LINR "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, LISRT "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'E' "
    sSQL = sSQL & " and LISRT.LINR = kassjour.LINR "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.LINR  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR inner join t on BEDZLINR.BEDIENER = t.Bediener and BEDZLINR.LINR = T.LINR "
    sSQL = sSQL & " set BEDZLINR.NUMSATZ = BEDZLINR.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "T", gdBase
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStO & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.LINR "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, LISRT "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'O' "
    sSQL = sSQL & " and LISRT.LINR = kassjour.LINR "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.LINR  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR inner join t on BEDZLINR.BEDIENER = t.Bediener and BEDZLINR.LINR = T.LINR "
    sSQL = sSQL & " set BEDZLINR.NUMSATZ = BEDZLINR.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR SET ERTRAG = NUMSATZ - EINKPREIS "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR SET NSP = ERTRAG * 100 / NUMSATZ where NUMSATZ <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    'Vorjahr
    loeschNEW "T", gdBase
    
    sSQL = "Select Sum(Preis)as Umsatzvj,KASSJOUR.Bediener,KASSJOUR.LINR "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, LISRT "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr - 1
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and LISRT.LINR = kassjour.LINR "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.LINR  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BEDZLINR inner join t on BEDZLINR.BEDIENER = t.Bediener and BEDZLINR.LINR = T.LINR "
    sSQL = sSQL & " set BEDZLINR.UMSATZvj = t.UMSATZvj  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Kopfdaten
    loeschNEW "BEDZKOPF", gdBase
    CreateTableT2 "BEDZKOPF", gdBase
    
    Dim sdat As String
    
    sdat = MonthName(imon) & " " & iJahr

    sSQL = "Insert into BEDZKOPF (Auswertungsdat) values ('" & sdat & "')"
    gdBase.Execute sSQL, dbFailOnError

    reportbildschirm "", "aWKLate"

    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BedienernachLINR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BedienernachAGN(imon As Integer, iJahr As Integer)
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    Dim sSQL As String
    Dim rsrs As Recordset

    loeschNEW "BEDZ", gdBase
    CreateTableT2 "BEDZ", gdBase
    
    anzeige "normal", "Schritt 1 von 34", lblanze

    sSQL = "Insert into BEDZ Select KASSJOUR.Bediener, BEDNAME.BEDNAME, AGNDBF.AGTEXT  ,KASSJOUR.AGN "
    sSQL = sSQL & " , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis "
    sSQL = sSQL & " , 0 as NUMSATZ "
    sSQL = sSQL & " , 0 as Kundenzahl "
    sSQL = sSQL & " from Kassjour, BEDNAME, AGNDBF"
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME ,KASSJOUR.AGN , AGNDBF.AGTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KUANTE", gdBase
    
    anzeige "normal", "Schritt 2 von 34", lblanze
    
    sSQL = "select adate, BELEGNR,Bediener  "
    sSQL = sSQL & " into KUANTE from Kassjour "
    
    sSQL = sSQL & " Where month(ADATE) = " & imon
    sSQL = sSQL & " and year(ADATE) = " & iJahr
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 3 von 34", lblanze
    
    sSQL = " Create index  Bediener on KUANTE(Bediener) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 4 von 34", lblanze
    sSQL = " Create index  BELEGNR on KUANTE(BELEGNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    anzeige "normal", "Schritt 5 von 34", lblanze
    
    sSQL = "Select * from BEDZ order by Bediener"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDIENER) Then
                rsrs.Edit
                
                anzeige "normal", "Schritt 6 von 34" & rsrs!BEDIENER & " " & rsrs!bedname, lblanze
                anzeige "normal", rsrs!BEDIENER, lblanze
                rsrs!Kundenzahl = KundenZahl_Bediener_Mon(imon, iJahr, CLng(rsrs!BEDIENER))
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    loeschNEW "T", gdBase
    
    anzeige "normal", "Schritt 7 von 34", lblanze
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStV & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'V' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 8 von 34", lblanze
    
    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "T", gdBase
    
    anzeige "normal", "Schritt 9 von 34", lblanze
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStE & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'E' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 10 von 34", lblanze
    
    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = BEDZ.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "T", gdBase
    
    anzeige "normal", "Schritt 11 von 34", lblanze
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStO & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'O' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 12 von 34", lblanze
    
    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = BEDZ.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 13 von 34", lblanze
    
    sSQL = "Update BEDZ SET ERTRAG = NUMSATZ - EINKPREIS "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 14 von 34", lblanze
    
    sSQL = "Update BEDZ SET NSP = ERTRAG * 100 / NUMSATZ where NUMSATZ <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    'Vorjahr
    loeschNEW "T", gdBase
    
    anzeige "normal", "Schritt 15 von 34", lblanze
    
    sSQL = "Select Sum(Preis)as Umsatzvj,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where month(Kassjour.ADATE) = " & imon
    sSQL = sSQL & " and year(Kassjour.ADATE) = " & iJahr - 1
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Schritt 16 von 34", lblanze
    
    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.UMSATZvj = t.UMSATZvj  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Kopfdaten
    loeschNEW "BEDZKOPF", gdBase
    CreateTableT2 "BEDZKOPF", gdBase
    
    Dim sdat As String
    
    sdat = MonthName(imon) & " " & iJahr
    
    anzeige "normal", "Schritt 17 von 34", lblanze

    sSQL = "Insert into BEDZKOPF (Auswertungsdat) values ('" & sdat & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblanze

    reportbildschirm "", "aWKLatd"
    
    anzeige "normal", "", lblanze

    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BedienernachAGN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub BedienernachAGN_ZR(cVon As String, cBis As String)
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    Dim sSQL As String
    Dim lVon As Long
    Dim lBis As Long
    Dim cDbis As String
    Dim cDvon As String
    Dim rsrs As Recordset
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
'    MsgBox (KundenZahl_Bediener(lVon, lBis, 9))

    cDvon = Trim$(Str$(lVon))
    cDbis = Trim$(Str$(lBis))

    loeschNEW "BEDZ", gdBase
    CreateTableT2 "BEDZ", gdBase

    sSQL = "Insert into BEDZ Select KASSJOUR.Bediener, BEDNAME.BEDNAME, AGNDBF.AGTEXT  ,KASSJOUR.AGN "
    sSQL = sSQL & " , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis "
    sSQL = sSQL & " , 0 as NUMSATZ "
    sSQL = sSQL & " , 0 as Kundenzahl "
    sSQL = sSQL & " from Kassjour, BEDNAME, AGNDBF"
    sSQL = sSQL & " Where Kassjour.ADATE between " & cDvon & " "
    sSQL = sSQL & " and  " & cDbis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME ,KASSJOUR.AGN , AGNDBF.AGTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from BEDZ order by Bediener"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDIENER) Then
                rsrs.Edit
                rsrs!Kundenzahl = KundenZahl_Bediener(lVon, lBis, CLng(rsrs!BEDIENER))
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    loeschNEW "T", gdBase
    
    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStV & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where Kassjour.ADATE between " & cDvon & " "
    sSQL = sSQL & " and  " & cDbis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'V' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "T", gdBase

    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStE & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where Kassjour.ADATE between " & cDvon & " "
    sSQL = sSQL & " and  " & cDbis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'E' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = BEDZ.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "T", gdBase

    sSQL = "Select Sum(Preis* 100/(100 + " & gdMWStO & "))as NUmsatz,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where Kassjour.ADATE between " & cDvon & " "
    sSQL = sSQL & " and  " & cDbis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.MWST = 'O' "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.NUMSATZ = BEDZ.NUMSATZ + t.NUMSATZ  "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ SET ERTRAG = NUMSATZ - EINKPREIS "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ SET NSP = ERTRAG * 100 / NUMSATZ where NUMSATZ <> 0 "
    gdBase.Execute sSQL, dbFailOnError

    'Vorjahr
    
    Dim cVonvj As String
    Dim cBisvj As String
    Dim lVonVJ As Long
    Dim lBisVJ As Long
    Dim cDbisVJ As String
    Dim cDvonVJ As String
    
    cVonvj = Left(cVon, 6) & Year(cVon) - 1
    cVonvj = Format(cVonvj, "DD.MM.YY")
    
    cBisvj = Left(cBis, 6) & Year(cBis) - 1
    cBisvj = Format(cBisvj, "DD.MM.YY")
    
    lVonVJ = DateValue(cVonvj)
    lBisVJ = DateValue(cBisvj)

    cDvonVJ = Trim$(Str$(lVonVJ))
    cDbisVJ = Trim$(Str$(lBisVJ))
    loeschNEW "T", gdBase

    sSQL = "Select Sum(Preis)as Umsatzvj,KASSJOUR.Bediener,KASSJOUR.AGN "
    sSQL = sSQL & " into t from Kassjour, BEDNAME, AGNDBF "
    sSQL = sSQL & " Where Kassjour.ADATE between " & cDvonVJ & " "
    sSQL = sSQL & " and  " & cDbisVJ & " "

    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " group BY  KASSJOUR.Bediener,KASSJOUR.AGN  "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update BEDZ inner join t on BEDZ.BEDIENER = t.Bediener and BEDZ.AGN = T.AGN "
    sSQL = sSQL & " set BEDZ.UMSATZvj = t.UMSATZvj  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Kopfdaten
    loeschNEW "BEDZKOPF", gdBase
    CreateTableT2 "BEDZKOPF", gdBase
    
    Dim sdat As String
    
    sdat = cVon & " - " & cBis

    sSQL = "Insert into BEDZKOPF (Auswertungsdat) values ('" & sdat & "')"
    gdBase.Execute sSQL, dbFailOnError

    reportbildschirm "", "aWKLatd"

    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BedienernachAGN_ZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    
    Set rsrs = gdApp.OpenRecordset("C138E")
    
    
    If Not rsrs.EOF Then
        
        Text2(0).Value = rsrs!Von
        Text2(1).Value = rsrs!Bis
        Option1(rsrs!iOpt1).Value = True
        Option2(rsrs!iOpt2).Value = True
        
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    Dim lVon As Long
    Dim lBis As Long
    Dim iOpt1 As Integer
    Dim iOpt2 As Integer
    
    loeschNEW "C138E", gdApp
    CreateTableT2 "C138E", gdApp
    
    lVon = Text2(0).Value
    lBis = Text2(1).Value
    
    If Option1(0).Value Then
        iOpt1 = 0
    ElseIf Option1(1).Value Then
        iOpt1 = 1
    ElseIf Option1(2).Value Then
        iOpt1 = 2
    ElseIf Option1(3).Value Then
        iOpt1 = 3
    ElseIf Option1(9).Value Then
        iOpt1 = 9
    End If
    
    If Option2(0).Value Then
        iOpt2 = 0
    ElseIf Option2(1).Value Then
        iOpt2 = 1
    ElseIf Option2(2).Value Then
        iOpt2 = 2
    End If
    
    sSQL = "Insert into C138E (von,bis,iopt1,iopt2) "
    sSQL = sSQL & " values (" & lVon & " ," & lBis & "," & iOpt1 & " ," & iOpt2 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label4_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 32 Then
    Label4(Index).Caption = "alle Farben"
    Label4(Index).Tag = ""
    Label4(Index).BackColor = Label1(1).BackColor
    Label4(Index).ForeColor = Label1(1).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR
    
    sortierenHGrid MSHFLEX1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        
        Case Is = 5     ' monat aktiv
            Text2(2).Visible = False
            Text2(3).Visible = False
            Command0(9).Visible = False
            Command0(10).Visible = False
            
            Combo3.Visible = True
        
        Case Is = 6     'zeitraum aktiv
            Text2(2).Visible = True
            Text2(3).Visible = True
            Command0(9).Visible = True
            Command0(10).Visible = True
            
            Combo3.Visible = False
        
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Option3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 4     'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text2(0).Value = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text2(1).Value = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text2(0).Value = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text2(1).Value = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case 2
                        Text2(1).Value = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case Else
                        Text2(1).Value = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
        Case Is = 5     'ak monat
            Text2(0).Value = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM.YY")
        Case Is = 8     'vorjahrzr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM") & "." & Year(Now) - 1
        Case Is = 6     'vorjahr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Text2(1).Value = "31.12." & Year(Now) - 1
        Case Is = 7     'ak jahr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)), "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM.YY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option3_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

Dim sAuswahlfeld As String
Dim ctmp As String
Dim lcount As Long

If KeyCode = vbKeyF2 Then
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    Select Case Index
        Case Is = 0
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "PGN"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List2.Visible = False
                List2.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List2.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List2.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List2.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), 2)
                        End If
                        
                    End If
                Next lcount
            End If
        Case Is = 1
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "BED"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List4.Visible = False
                List4.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List4.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List4.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List4.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                        End If
                        
                    End If
                Next lcount
                
            End If
        Case Is = 2
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            
            List3.Visible = False 'die linien auf standard
            List3.Clear
            
            List5.Visible = False
            List5.Clear
            For lcount = 0 To 100
                If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                    List5.Visible = True
                    Text1(Index).Text = ""
                    
                    If gF2Prompt.cArray(lcount) <> "" Then
                        List5.AddItem gF2Prompt.cArray(lcount)
                    End If
                
                Else
                    If gF2Prompt.cArray(lcount) <> "" Then
                       
                        List5.AddItem gF2Prompt.cArray(lcount)
                        Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                    End If
                    
                End If
            Next lcount
            
            
            
            
            
            
            
            
            
            
'            If gF2Prompt.cWahl <> "" Then
'                Text1(Index).Text = gF2Prompt.cWahl
'            End If
        Case Is = 4
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "AGN"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                
                
                List1.Visible = False
                List1.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List1.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List1.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List1.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                        End If
                        
                    End If
                Next lcount

            End If

        Case 5
            ctmp = Text1(7).Text
            ctmp = Trim$(ctmp)
            If ctmp = "" Then
                ctmp = Text1(2).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    anzeige "Rot", "Bitte einen Lieferanten oder eine Marke angeben!", lblanzeige
                    Text1(7).SetFocus
                    Exit Sub
                Else
                    sAuswahlfeld = "LINR"
                End If
            Else
                sAuswahlfeld = "MARKE"
            End If
            
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "LPZ"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cEsFeld = sAuswahlfeld
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount) & Space(50) & Right(gF2Prompt.cArray(lcount), 6)
                        End If
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 3)
                        End If
                    End If
                Next lcount
            End If

        Case Is = 7
            gF2Prompt.cFeld = "MARKE"
            
            ctmp = Text1(2).Text 'Linr eventuell
            gF2Prompt.cEsFeld = ctmp
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
            End If

        End Select
        Text1(Index).SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
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
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

voreinstellungspeichern
LogtoEnd Me

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 6, 1, 2, 5, 4, 0 'ARTNR, EAN, LIEFNR, ARTGRU,linie
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 8  'Proz
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 3, 7       'BEZEICH, LIBESNR
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß#"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            'alle Zeichen erlaubt
        
            
    End Select
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Bedienerbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

