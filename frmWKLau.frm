VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKLau 
   BackColor       =   &H00C0C000&
   Caption         =   "Lieferanten Statistik"
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
   Icon            =   "frmWKLau.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "sortiert nach"
      Height          =   1215
      Left            =   7560
      TabIndex        =   47
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Libesnr"
         ForeColor       =   &H00404000&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   50
         ToolTipText     =   "Hitliste gemessen an den verkauften Artikeln"
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung"
         ForeColor       =   &H00404000&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   49
         ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Linie"
         ForeColor       =   &H00404000&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   48
         ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Druck sortiert nach:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   1095
      End
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   6
      Left            =   10920
      TabIndex        =   40
      Top             =   120
      Width           =   345
      _ExtentX        =   609
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
      ButtonStyle     =   2
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   9840
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "nur umsatzrelevante"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   720
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin sevCommand3.Command cmdGo 
      Height          =   360
      Left            =   9120
      TabIndex        =   38
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
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
      MaskColor       =   16777215
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
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   11
      Left            =   11400
      TabIndex        =   37
      Top             =   600
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   635
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
      Height          =   375
      Left            =   9120
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
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
      Caption         =   "Details"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdBack 
      Height          =   375
      Left            =   9120
      TabIndex        =   27
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
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
      Caption         =   "zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
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
      Left            =   4680
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optEw 
         BackColor       =   &H00C0C000&
         Caption         =   "Entwicklung"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optLW 
         BackColor       =   &H00C0C000&
         Caption         =   "Lagerwerte"
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
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
      Left            =   5160
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton opt5 
         BackColor       =   &H00C0C000&
         Caption         =   "Top 5"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton opt10 
         BackColor       =   &H00C0C000&
         Caption         =   "Top 10"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   4335
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         Caption         =   "sortiert nach"
         Height          =   855
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         Begin VB.OptionButton option5 
            BackColor       =   &H00C0C000&
            Caption         =   "Liefbezeichnung"
            ForeColor       =   &H00404000&
            Height          =   195
            Index           =   6
            Left            =   1320
            TabIndex        =   52
            ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton option5 
            BackColor       =   &H00C0C000&
            Caption         =   "Penneranteil"
            ForeColor       =   &H00404000&
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   36
            ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton option5 
            BackColor       =   &H00C0C000&
            Caption         =   "Nettospanne"
            ForeColor       =   &H00404000&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   34
            ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton option5 
            BackColor       =   &H00C0C000&
            Caption         =   "Umsatz"
            ForeColor       =   &H00404000&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   33
            ToolTipText     =   "Hitliste gemessen an den verkauften Artikeln"
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "sortiert nach"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.OptionButton optH 
         BackColor       =   &H00C0C000&
         Caption         =   "Hitliste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   3120
         TabIndex        =   19
         ToolTipText     =   "sortiert nach Lieferanten, verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optqp 
         BackColor       =   &H00C0C000&
         Caption         =   "Verkaufszahlen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag, Umsatz VK LJZR, Umsatz EK LJZR, Lager VK, Lager EK, Lager LEK"
         Top             =   120
         Width           =   2775
      End
      Begin VB.OptionButton optD 
         BackColor       =   &H00C0C000&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame5"
      Height          =   735
      Left            =   3360
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton optA 
         BackColor       =   &H00C0C000&
         Caption         =   "vk. Artikel"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         ToolTipText     =   "Hitliste gemessen an den verkauften Artikeln"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optU 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz (VK)"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "Hitliste gemessen am erzielten Umsatz (VK)"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H00C0C000&
         Caption         =   "Ertrag"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         ToolTipText     =   "Hitliste gemessen am erzielten Ertrag"
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "nach"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin sevCommand3.Command cmdPrint 
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   7440
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdEnd 
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   7920
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdDel 
      Height          =   360
      Left            =   11400
      TabIndex        =   2
      ToolTipText     =   "Löschen Ihrer Eingaben"
      Top             =   120
      Width           =   345
      _ExtentX        =   609
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
      ButtonStyle     =   2
      Caption         =   "L"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   5655
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Height          =   495
      Left            =   6240
      TabIndex        =   41
      Top             =   480
      Width           =   2775
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "akt Jahr"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   44
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Vormonat"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "aktueller Monat"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton Option1 
         Caption         =   "Schnitt - EK"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Listen - EK"
         Height          =   210
         Index           =   1
         Left            =   1440
         TabIndex        =   29
         Top             =   0
         Width           =   1215
      End
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   0
      Left            =   10920
      TabIndex        =   53
      Top             =   600
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   635
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
      Picture         =   "frmWKLau.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   20
      Left            =   3840
      TabIndex        =   54
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   21
      Left            =   5760
      TabIndex        =   55
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "kein Lieferant"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6360
      TabIndex        =   46
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "von:"
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
      Left            =   2400
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "bis:"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   720
      Width           =   375
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
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   7455
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferanten Statistik"
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
      Width           =   6135
   End
End
Attribute VB_Name = "frmWKLau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSortklick As String

Dim bLWD As Boolean
Dim bEW As Boolean
Dim bLW As Boolean
Dim bHit As Boolean
Dim bGoPlus As Boolean
Dim bGo As Boolean
Dim bDetail As Boolean
Dim bDetailPlus As Boolean
Private Sub cmdBack_Click()
    On Error GoTo LOKAL_ERROR

    cmdBack.Visible = False
    If bDetail Then
    
        
        cmdGo_Click
        
    ElseIf bDetailPlus Then
        
        Frame3.Visible = False
        optqp.Value = True
        cmdGo_Click
    
    ElseIf bLWD Then
        
        optLW.Value = True
        cmdGo_Click
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdBack_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub
Private Sub cmdDel_Click()
    On Error GoTo LOKAL_ERROR

    cmdBack.Visible = False

    bLWD = False
    bLW = False
    bEW = False
    bHit = False
    bGoPlus = False
    bGo = False
    bDetail = False
    bDetailPlus = False

    Frame5.Visible = False
    Frame6.Visible = True
    Frame7.Visible = False
    Frame8.Visible = False
    
    optqp.Value = True
    optE.Value = True
    lblanzeige.Caption = ""
    lblanzeige.Refresh
        
    Text1(2).Text = ""
    Text1(2).SetFocus
    
    
    MSHFLEX1.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdDel_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKLau
        
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleMSHFLEXGoPlus()
    On Error GoTo LOKAL_ERROR

    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 12
        .FixedRows = 1
        .FixedCols = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1000
        .Text = "Lieferant"
        
        .Col = 1
        .ColWidth(1) = 3400
        .Text = "Lieferantenname"
        
        .Col = 2
        .ColWidth(2) = 1200
        .Text = "verk.Artikel"
        
        .Col = 3
        .ColWidth(3) = 1200
        .Text = "Umsatz (VK)"
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Umsatz (EK)"
        
        .Col = 5
        .ColWidth(5) = 1000
        .Text = "Ertrag"
        
        .Col = 6
        .ColWidth(6) = 1700
        .Text = "Umsatz (VK) VJ ZR"
        
        .Col = 7
        .ColWidth(7) = 1800
        .Text = "Umsatz (EK) VJ ZR"
        
        .Col = 8
        .ColWidth(8) = 1400
        .Text = "Lager VK-Wert"
        
        .Col = 9
        .ColWidth(9) = 1400
        .Text = "Lager EK-Wert"
        
        .Col = 10
        .ColWidth(10) = 1600
        .Text = "Lager LEK-Wert"
        
        .Col = 11
        .ColWidth(11) = 1000
        .Text = "Bestand"
        
    End With
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXGoPlus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleMSHFLEXLW()
    On Error GoTo LOKAL_ERROR
    
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
        .Text = "Lieferant"
        
        .Col = 1
        .ColWidth(1) = 3200
        .Text = "Lieferantenname"
        
        .Col = 2
        .ColWidth(2) = 800
        .Text = "Bestand"
        
        .Col = 3
        .ColWidth(3) = 1200
        .Text = "Lager VK-Wert"
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Lager EK-Wert"
        
        
        
        
        
    End With
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXLW"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function ermitteln()
    On Error GoTo LOKAL_ERROR

    Dim sSQLLinie           As String
    Dim sSQLlief            As String
    Dim sSQLFil             As String
    Dim sSQL                As String
    Dim sLiefname           As String
    Dim sFILname            As String
    Dim sLiefnr             As String
    
    Dim lLinr               As Long
    Dim lVon                As Long
    Dim lBis                As Long
        
    Dim cVon                As String
    Dim cBis                As String
    
    Dim cVonvj              As String
    Dim cBisvj              As String
    Dim i                   As Integer
    Dim rsrs                As Recordset
    Dim dLagerwertzumSEK    As Double
    Dim lLagerST            As Long
    Dim dPennerwertzumSEK   As Double
    Dim lPennerST           As Long
    Dim lAnz                As Long
    Dim dPennerAnteilSEK    As Double
    Dim dPennerAnteilST     As Double
    Dim iFil                As Integer
    
    Dim bgef                As Boolean
    Dim cVorjahr            As String
    
    pbrZeit.Visible = True
    pbrZeit.Max = 1000
    pbrZeit.Value = 100
    
    lVon = DateValue(Trim$(Text1(0).Text))
    lBis = DateValue(Trim$(Text1(1).Text))
   
    anzeige "normal", "Daten für diesen Zeitraum werden ermittelt...", lblanzeige
    
    sSQLLinie = ""

    If Trim$(Text1(2).Text) = "" Then
        sSQLlief = ""
        sLiefname = "alle Lieferanten"
    Else
        If IsNumeric(Trim$(Text1(2).Text)) Then
        
            sSQLlief = " and KASSJOUR.artnr in (Select artnr from artlief where linr = " & Trim$(Text1(2).Text) & ")"
'            sSQLlief = " and KASSJOUR.LINR = " & Trim$(Text1(2).Text) & ""
            
            sLiefname = ermLiefBez(CLng(Trim$(Text1(2).Text)))
        Else
            sSQLlief = ""
            sLiefname = "alle Lieferanten"
        End If
    End If

    cVon = Format(Text1(0).Text, "DD.MM.YY")
    cBis = Format(Text1(1).Text, "DD.MM.YY")
    
    
    cVorjahr = Val(Right$(cVon, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    
    cVonvj = Left$(cVon, 6) & cVorjahr
    
    cVorjahr = Val(Right$(cBis, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    cBisvj = Left$(cBis, 6) & cVorjahr
    
    If cBisvj = "29.02.11" Or cBisvj = "29.02.2011" Then
    
        cBisvj = "28.02.11"
        
    ElseIf cBisvj = "29.02.15" Or cBisvj = "29.02.2015" Then
        cBisvj = "28.02.15"
    ElseIf cBisvj = "29.02.19" Or cBisvj = "29.02.2019" Then
        cBisvj = "28.02.19"
    ElseIf cBisvj = "29.02.23" Or cBisvj = "29.02.2023" Then
        cBisvj = "28.02.23"
    End If
    
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    lVon = DateValue(cVonvj)
    
    
    lBis = DateValue(cBisvj)
    
    cVonvj = Trim$(Str$(lVon))
    cBisvj = Trim$(Str$(lBis))
    pbrZeit.Value = 200

    loeschNEW "liefplus", gdBase
    CreateTable "LIEFPLUS", gdBase

    pbrZeit.Value = 150
    
    sSQL = "INSERT into LIEFPLUS Select LINR, LIEFBEZ  "
    sSQL = sSQL & " from LISRT "
    gdBase.Execute sSQL, dbFailOnError

    If Check1.Value = vbChecked Then
        bgef = True
    Else
        bgef = False
    End If
    
    ermittleLiefdetail "Sum(KASSJOUR.Menge)", "Anzahl", cVon, cBis, sSQLlief & sSQLFil & sSQLLinie, bgef, 200, "VK Mengen werden ermittelt...", lblanzeige
    ermittleLiefdetail "Sum(KASSJOUR.Preis)", "Umsatz", cVon, cBis, sSQLlief & sSQLFil & sSQLLinie, bgef, 250, "Umsätze werden ermittelt...", lblanzeige
    ermittleLiefdetail "Sum(KASSJOUR.EKPR*KASSJOUR.Menge)", "EINKPREIS", cVon, cBis, sSQLlief & sSQLFil & sSQLLinie, bgef, 300, "Umsätze zum EK werden ermittelt...", lblanzeige
    
    'umsatznetto V
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStV & ")as umsatznettoakt "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.MWST = 'V'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY LINR"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettoakt =  Te12.umsatznettoakt"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 350
    
    'umsatznetto E
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStE & ")as umsatznettoakt "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.MWST = 'E'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE Liefplus set umsatznettoakt = 0 where umsatznettoakt is null "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettoakt = Liefplus.umsatznettoakt + Te12.umsatznettoakt"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 400
    
    'umsatznetto O
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStO & ")as umsatznettoakt "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.MWST = 'O'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE Liefplus set umsatznettoakt = 0 where umsatznettoakt is null "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettoakt = Liefplus.umsatznettoakt + Te12.umsatznettoakt"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set Ertragakt = umsatznettoakt - einkPreis"
    gdBase.Execute sSQL, dbFailOnError

    
    pbrZeit.Value = 450
    sSQL = "Update Liefplus set mindat = '" & Text1(0).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set maxdat = '" & Text1(1).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set auswahl = '" & sLiefname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set Fauswahl = '" & sFILname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UmsatzVKVJ = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UmsatzEKVJ = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 500
    
    ermittleLiefdetail "Sum(KASSJOUR.Menge)", "Anzahlvj", cVonvj, cBisvj, sSQLlief & sSQLFil & sSQLLinie, bgef, 550, "VK Mengen vom Vorjahreszeitraum werden ermittelt...", lblanzeige
    ermittleLiefdetail "Sum(KASSJOUR.Preis)", "UmsatzVKVJ", cVonvj, cBisvj, sSQLlief & sSQLFil & sSQLLinie, bgef, 600, "Umsätze vom Vorjahreszeitraum werden ermittelt...", lblanzeige
    ermittleLiefdetail "Sum(EKPR*Menge)", "UmsatzEKVJ", cVonvj, cBisvj, sSQLlief & sSQLFil & sSQLLinie, bgef, 650, "Umsätze zum EK vom Vorjahreszeitraum werden ermittelt...", lblanzeige
    
    
    
    anzeige "normal", "Nettoumsätze vom Vorjahreszeitraum werden ermittelt...", lblanzeige
    'umsatznetto V
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStV & ")as umsatznettovj "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVonvj & " And " & cBisvj & " "
    sSQL = sSQL & " and Kassjour.MWST = 'V'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettovj =  Te12.umsatznettovj"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 700
    
    'umsatznetto E
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStE & ")as umsatznettovj "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVonvj & " And " & cBisvj & " "
    sSQL = sSQL & " and Kassjour.MWST = 'E'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettovj = Liefplus.umsatznettovj + Te12.umsatznettovj"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 750
    
    'umsatznettovj O
    loeschNEW "te12", gdBase
    
    If IsNumeric(Trim$(Text1(2).Text)) Then
        sSQL = "SELECT " & Trim$(Text1(2).Text) & " as Linr "
    Else
        sSQL = "SELECT KASSJOUR.Linr"
    End If
    sSQL = sSQL & ", (Sum(Preis)* 100)/(100 + " & gdMWStO & ")as umsatznettovj "
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVonvj & " And " & cBisvj & " "
    sSQL = sSQL & " and Kassjour.MWST = 'O'"
    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sSQLlief & sSQLFil & sSQLLinie
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus.umsatznettovj = Liefplus.umsatznettovj + Te12.umsatznettovj"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 800
    
    sSQL = "Update Liefplus set Ertragvj = umsatznettovj - UmsatzEKVJ"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set NSPvj = Ertragvj*100 / umsatznettovj"
    sSQL = sSQL & " where umsatznettovj <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set NSPakt = Ertragakt*100 / umsatznettoakt"
    sSQL = sSQL & " where umsatznettoakt <> 0 "
    gdBase.Execute sSQL, dbFailOnError
   
    sSQL = "Update Liefplus set DIFFUMSEUR = UMSATZ - UmsatzVKVJ"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSPROZ =DIFFUMSEUR * 100 / UmsatzVKVJ " 'UMSATZ "
    sSQL = sSQL & " where UmsatzVKVJ <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSEUREK = EINKPREIS - UmsatzEKVJ"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSPROZEK =DIFFUMSEUREK * 100 / UmsatzEKVJ " 'EINKPREIS "
    sSQL = sSQL & " where UmsatzEKVJ <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 900
    
    loeschNEW "te12", gdBase
    
    anzeige "normal", "Lagerwerte werden ermittelt...", lblanzeige
    
    sSQL = "Update Liefplus set UMSATZ = 0 where UMSATZ is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UmsatzVKVJ = 0 where UmsatzVKVJ is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Liefplus where UMSATZ = 0 and UmsatzVKVJ = 0  "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Delete from Liefplus where UMSATZ <= 0 and UmsatzVKVJ <= 0  "
'    gdBase.Execute sSQL, dbFailOnError
    

    LagerwerteschreibenLINRJetzt lblanzeige

    pbrZeit.Value = 950

    sSQL = "Select * from Liefplus"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr

                lAnz = lAnz - 1
                anzeige "normal", "Lieferant: " & rsrs!LIEFBEZ & " noch " & CStr(lAnz) & " Lieferanten ...", lblanzeige

                dLagerwertzumSEK = LAGEREKermittlungJetzt(lLinr)
                lLagerST = LAGERStückErmittlungJetzt(lLinr)

                If iFil <> 0 Then
                    dPennerwertzumSEK = 0
                    lPennerST = 0
                Else
                    dPennerwertzumSEK = PennerEKermittlungJetzt(lLinr)
                    lPennerST = PennerStückErmittlungJetzt(lLinr)
                End If
            Else
                dLagerwertzumSEK = 0
                lLagerST = 0

                dPennerwertzumSEK = 0
                lPennerST = 0
            End If

            rsrs.Edit
            rsrs!LAGERWSEK = dLagerwertzumSEK
            rsrs!LAGERST = lLagerST

            rsrs!PENNERWSEK = dPennerwertzumSEK
            rsrs!PENNERST = lPennerST

            dPennerAnteilSEK = 0
            If dLagerwertzumSEK <> 0 Then
                dPennerAnteilSEK = 100 * dPennerwertzumSEK / dLagerwertzumSEK
            End If

            dPennerAnteilST = 0
            If lLagerST <> 0 Then
                dPennerAnteilST = 100 * lPennerST / lLagerST
            End If

            rsrs!PENANTEILST = dPennerAnteilST
            rsrs!PENANTEILSEK = dPennerAnteilSEK

            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "Update Liefplus set PENNERWSEK = 0 where PENNERWSEK is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set PENNERST = 0 where PENNERST is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set PENANTEILST = 0 where PENANTEILST is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set PENANTEILSEK = 0 where PENANTEILSEK is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UMSATZNETTOakt = 0 where UMSATZNETTOakt is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UMSATZNETTOvj = 0 where UMSATZNETTOvj is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSEUREK = 0 where DIFFUMSEUREK is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSPROZEK = 0 where DIFFUMSPROZEK is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set ERTRAGakt = 0 where ERTRAGakt is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set ERTRAGvj = 0 where ERTRAGvj is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set NSPakt = 0 where NSPakt is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set NSPvj = 0 where NSPvj is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set LAGERWSEK = 0 where LAGERWSEK is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set LAGERST = 0 where LAGERST is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set ANZAHL = 0 where anzahl is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set ANZAHLvj = 0 where ANZAHLvj is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSEUR = 0 where DIFFUMSEUR is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set DIFFUMSPROZ = 0 where DIFFUMSPROZ is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set EINKPREIS = 0 where EINKPREIS is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Liefplus set UmsatzEKVJ = 0 where UmsatzEKVJ is null"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    

    pbrZeit.Value = 1000

    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermitteln"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Function
Private Sub ermittleLiefdetail(sSum As String, sAlswas As String, cVon As String, cBis As String, sEinschr As String, bgef As Boolean, iWert As Integer, sMsgText As String, lblx As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    anzeige "normal", sMsgText, lblx
    
    'Anzahl
    loeschNEW "te12", gdBase
    
    sSQL = "SELECT KASSJOUR.Linr, " & sSum & " as " & sAlswas
    sSQL = sSQL & " INTO Te12 "
    sSQL = sSQL & " From KASSJOUR "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    If bgef Then
        sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    End If
    sSQL = sSQL & sEinschr
    sSQL = sSQL & " GROUP BY KASSJOUR.LINR"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "UPDATE Liefplus INNER JOIN Te12 ON "
    sSQL = sSQL & " Liefplus.linr = Te12.linr "
    sSQL = sSQL & " set Liefplus." & sAlswas & " = Te12." & sAlswas
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = iWert

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleLiefdetail"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."

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
    
        anzeige "normal", "Es werden " & lMax & " Lieferanten angezeigt...", lblanzeige
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
                        Case Is = "Ums akt", "Ums akt EK", "Ertrag", "Ums VJ ZR", "Ums VJ ZR EK", "Ertrag akt", "Ertrag VJ"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                        Case Is = "NSP akt", "NSP VJ"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                        Case Is = "LAGER(SEK)", "Penner(SEK)"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "#######0.00")
                                
                        Case Is = "Panteil Stück in %", "Panteil SEK in %"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "#######0.00")
                            
                        Case Is = "Diff Ums Eur EK", "Diff Ums % EK", "Diff Ums Eur", "Diff Ums %"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            If CDbl(sWert) < 0 Then
                                .CellForeColor = vbRed
                            Else
                                .CellForeColor = vbBlack
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
                    
                    If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                        aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                    End If
                    
                End If
            Next i
                                
            rsrs.MoveNext
        Loop
        

        anzeige "normal", "Es wurden " & lMax & " Lieferanten ermittelt.", lblanzeige
    Else
        Frame1.Visible = True
        anzeige "rot", "Es wurden keine Lieferanten ermittelt.", lblanzeige
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
    
    If lMax > 1 Then
        .RowHeight(1) = 0
    End If
    
    lrow = lrow - 1
    .Redraw = True
    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
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
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdGo_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sOrder As String
    Dim sSQL As String
    
    cmdBack.Visible = False
    lblanzeige.Caption = ""
    lblanzeige.Refresh
    
    Screen.MousePointer = 11
    
    
    MSHFLEX1.Visible = False
   
    
    If optqp.Value = True Then
    
        Tabcheck "LIEFSTAT"
        FormatGridOverTablay "LIEFSTAT"

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
        
        If Option5(0).Value = True Then
            sOrder = " Order by Umsatz desc" 'Umsatz
        ElseIf Option5(1).Value = True Then
            sOrder = " Order by NSPAkt desc" 'NSPAkt
        ElseIf Option5(2).Value = True Then
            sOrder = " Order by PENANTEILST desc" 'Penner
        ElseIf Option5(6).Value = True Then
            sOrder = " Order by Liefbez " 'Liefbez
        End If
        
        
        
        'hier Hintergrundtabelle vorsortieren
        
        loeschNEW "LItot", gdBase
        sSQL = "select * into LItot from LIEFPLUS " & sOrder
        gdBase.Execute sSQL
        
        loeschNEW "LIEFPLUS", gdBase
    
        sSQL = "select * into LIEFPLUS from LItot "
        gdBase.Execute sSQL, dbFailOnError
        loeschNEW "LItot", gdBase
        'Ende hier Hintergrundtabelle vorsortieren
        
        GridFuellen "Select * from LIEFPLUS " & sOrder
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        
        Command1.Visible = True
        
        bGoPlus = True
        
        bLWD = False
        bLW = False
        bGo = False
        bDetail = False
        bDetailPlus = False
        bEW = False
        bHit = False
        
    ElseIf optH.Value = True Then
    
        ErstelleMSHFLEXGo
        liefstatErstellen
        hitlisteVKWERT
        
        bHit = True
        
        bLWD = False
        bGoPlus = False
        bGo = False
        bDetail = False
        bEW = False
        bLW = False
        bDetailPlus = False
       
    ElseIf optEw.Value = True Then
    
        liefstatEWErstellen

        bEW = True
    
        bLWD = False
        bGoPlus = False
        bGo = False
        bHit = False
        bDetail = False
        bLW = False
        bDetailPlus = False

    ElseIf optLW.Value = True Then
        
        ErstelleMSHFLEXLW
        liefstatLWErstellen
        LWAuswertung
    
        bLW = True
        
        bLWD = False
        bGoPlus = False
        bGo = False
        bHit = False
        bDetail = False
        bEW = False
        bDetailPlus = False
        
        
    Else
    
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub GoAuswertungD()
    On Error GoTo LOKAL_ERROR

    
    Dim sSQL As String
    Dim sSort As String
    Dim rs As Recordset
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim cFeld As String
    Dim dEkpr As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sArtnr As String
    Dim sLPZ As String
    
    sSort = "LPZ"
    
    loeschNEW "Lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from gode "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "gode", gdBase
    
    sSQL = "Select * into gode from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from gode order by " & sSort
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 1
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!artnr) Then
                sArtnr = rs!artnr
            Else
                sArtnr = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sArtnr
            
            If Not IsNull(rs!BEZEICH) Then
            cFeld = rs!BEZEICH
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!LPZ) Then
            cFeld = rs!LPZ
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            MSHFLEX1.Col = 4
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 5
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!ERTRAG) Then
                dErtrag = rs!ERTRAG
            Else
                dErtrag = "0"
            End If
            
            MSHFLEX1.Col = 6
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
             rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
            
    pbrZeit.Visible = False
    
    If lrow = 0 Then
        lblanzeige.Caption = "Keine Daten gefunden"
        lblanzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.RowHeight(1) = 0
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
  
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GoAuswertungD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub GoAuswertungDPlus()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim sSQL As String
    Dim sSort As String
    Dim cFeld As String
    Dim dEkpr As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sArtnr As String
    Dim lBestand As Long
    
    Dim dUVKVJ As Double
    Dim dUEKVJ As Double
    Dim dLVK As Double
    Dim dLEK As Double
    Dim dLLEK As Double
    
    sSort = "LPZ"
    
    loeschNEW "Lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from gode "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "gode", gdBase
    
    sSQL = "Select * into gode from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from gode order by " & sSort
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 1
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!artnr) Then
                sArtnr = rs!artnr
            Else
                sArtnr = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sArtnr
            
            If Not IsNull(rs!BEZEICH) Then
            cFeld = rs!BEZEICH
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!LIBESNR) Then
            cFeld = rs!LIBESNR
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!LPZ) Then
            cFeld = rs!LPZ
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 4
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            MSHFLEX1.Col = 5
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 6
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!ERTRAG) Then
                dErtrag = rs!ERTRAG
            Else
                dErtrag = "0"
            End If
            
            MSHFLEX1.Col = 7
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            
            
            
            If Not IsNull(rs!UmsatzVKVJ) Then
                dUVKVJ = rs!UmsatzVKVJ
            Else
                dUVKVJ = "0"
            End If
            
            MSHFLEX1.Col = 8
            cFeld = Format$(dUVKVJ, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!UmsatzEKVJ) Then
                dUEKVJ = rs!UmsatzEKVJ
            Else
                dUEKVJ = "0"
            End If
            
            MSHFLEX1.Col = 9
            cFeld = Format$(dUEKVJ, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!LagerVK) Then
                dLVK = rs!LagerVK
            Else
                dLVK = "0"
            End If
            
            MSHFLEX1.Col = 10
            cFeld = Format$(dLVK, "######0.00")
            MSHFLEX1.Text = cFeld
        
            If Not IsNull(rs!LagerEK) Then
                dLEK = rs!LagerEK
            Else
                dLEK = "0"
            End If
            
            MSHFLEX1.Col = 11
            cFeld = Format$(dLEK, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!LagerLEK) Then
                dLLEK = rs!LagerLEK
            Else
                dLLEK = "0"
            End If
            
            MSHFLEX1.Col = 12
            cFeld = Format$(dLLEK, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!BESTAND) Then
                lBestand = rs!BESTAND
            Else
                lBestand = "0"
            End If
            
            MSHFLEX1.Col = 13
            MSHFLEX1.Text = lBestand
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    pbrZeit.Visible = False
    
    
    If lrow = 0 Then
        lblanzeige.Caption = "Keine Daten gefunden"
        lblanzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.RowHeight(1) = 0
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
  
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GoAuswertungDPlus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GoAuswertungLWD()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim sSQL As String
    Dim sSort As String
    Dim cFeld As String
    Dim dEkpr As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sArtnr As String
    
    Dim dUVKVJ As Double
    Dim dUEKVJ As Double
    Dim dLVK As Double
    Dim dLEK As Double
    Dim dLLEK As Double
    
    sSort = "LPZ"
    
    loeschNEW "lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from gode "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "gode", gdBase
    
    sSQL = "Select * into GODE from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from gode order by " & sSort
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 1
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!artnr) Then
                sArtnr = rs!artnr
            Else
                sArtnr = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sArtnr
            
            If Not IsNull(rs!BEZEICH) Then
            cFeld = rs!BEZEICH
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!LPZ) Then
            cFeld = rs!LPZ
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!BESTAND) Then
                iAnz = rs!BESTAND
            Else
                iAnz = "0"
            End If

            MSHFLEX1.Col = 3
            MSHFLEX1.Text = iAnz
            
            If Not IsNull(rs!LagerVK) Then
                dLVK = rs!LagerVK
            Else
                dLVK = "0"
            End If
            
            MSHFLEX1.Col = 4
            cFeld = Format$(dLVK, "######0.00")
            MSHFLEX1.Text = cFeld
        
            If Not IsNull(rs!LagerEK) Then
                dLEK = rs!LagerEK
            Else
                dLEK = "0"
            End If
            
            MSHFLEX1.Col = 5
            cFeld = Format$(dLEK, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    pbrZeit.Visible = False
    
    
    If lrow = 0 Then
        lblanzeige.Caption = "Keine Daten gefunden"
        lblanzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.RowHeight(1) = 0
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
  
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GoAuswertungLWD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub
Private Sub LWAuswertung()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim lAnzahl
    Dim lrow As Long
    Dim counter As Integer
    Dim sSQL As String
    Dim sSort As String
    Dim cFeld As String
    Dim dEkpr As Single
    Dim lAnz As Long
    Dim sLinr As String
    
    Dim dLVK As Double
    Dim dLEK As Double
    Dim dLLEK As Double
    
    sSort = ""
    
'    If optUVK.Value Then
'        sSORT = "Bestand desc"
'    Else
        sSort = "Liefbez"
'    End If
    
    
    sSQL = "Select * from LiefLW order by " & sSort
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    
    counter = 0
    lrow = 1
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!linr) Then
                sLinr = rs!linr
            Else
                sLinr = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sLinr
            
            If Not IsNull(rs!LIEFBEZ) Then
            cFeld = rs!LIEFBEZ
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!BESTAND) Then
                lAnz = rs!BESTAND
            Else
                lAnz = "0"
            End If
    
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = lAnz
            
            
            If Not IsNull(rs!LagerVK) Then
                dLVK = rs!LagerVK
            Else
                dLVK = "0"
            End If
            
            MSHFLEX1.Col = 3
            cFeld = Format$(dLVK, "######0.00")
            MSHFLEX1.Text = cFeld
        
            If Not IsNull(rs!LagerEK) Then
                dLEK = rs!LagerEK
            Else
                dLEK = "0"
            End If
            
            MSHFLEX1.Col = 4
            cFeld = Format$(dLEK, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
            
    pbrZeit.Visible = False
    lblanzeige.Caption = "Lagerwerte von " & lrow & " Lieferanten erstellt, deren Artikel heute im Bestand sind."
    lblanzeige.Refresh
    
    If lrow = 0 Then
        lblanzeige.Caption = "Keine Daten gefunden"
        lblanzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.RowHeight(1) = 0
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        Command1.Visible = True
    End If
  
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LWAuswertung"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ErstelleMSHFLEXGo()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 8350
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 6
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
        .Text = "Lieferant"
        
        .Col = 1
        .ColWidth(1) = 3200
        .Text = "Lieferantenname"
        
        .Col = 2
        .ColWidth(2) = 1000
        .Text = "verk.Artikel"
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "Umsatz (VK)"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Umsatz (EK)"
        
        .Col = 5
        .ColWidth(5) = 1000
        .Text = "Ertrag"

    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXGo"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleMSHFLEXGoD()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 9350
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
        .Text = "Artnr"
        
        .Col = 1
        .ColWidth(1) = 3200
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .ColWidth(2) = 700
        .Text = "Linie"
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "verk.Artikel"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Umsatz (VK)"
        
        .Col = 5
        .ColWidth(5) = 1000
        .Text = "Umsatz (EK)"
        
        .Col = 6
        .ColWidth(6) = 1000
        .Text = "Ertrag"

    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXGoD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErstelleMSHFLEXGoDPlus()
    On Error GoTo LOKAL_ERROR
        
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 14
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
        .Text = "Artnr"
        
        .Col = 1
        .ColWidth(1) = 3200
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .ColWidth(2) = 1000
        .Text = "LiefBestNr"
        
        .Col = 3
        .ColWidth(3) = 700
        .Text = "Linie"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "verk.Artikel"
        
        .Col = 5
        .ColWidth(5) = 1000
        .Text = "Umsatz (VK)"
        
        .Col = 6
        .ColWidth(6) = 1000
        .Text = "Umsatz (EK)"
        
        .Col = 7
        .ColWidth(7) = 1000
        .Text = "Ertrag"

        .Col = 8
        .ColWidth(8) = 1500
        .Text = "Umsatz (VK) VJ ZR"
        
        .Col = 9
        .ColWidth(9) = 1500
        .Text = "Umsatz (EK) VJ ZR"
        
        .Col = 10
        .ColWidth(10) = 1200
        .Text = "Lager VK-Wert"
        
        .Col = 11
        .ColWidth(11) = 1200
        .Text = "Lager EK-Wert"
        
        .Col = 12
        .ColWidth(12) = 1300
        .Text = "Lager LEK-Wert"
        
        .Col = 13
        .ColWidth(13) = 1000
        .Text = "Bestand"
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXGoDPlus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErstelleMSHFLEXLWD()
    On Error GoTo LOKAL_ERROR
    

    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 6
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
        .Text = "Artnr"
        
        .Col = 1
        .ColWidth(1) = 3200
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .ColWidth(2) = 700
        .Text = "Linie"
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "Bestand"
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Lager VK-Wert"
        
        .Col = 5
        .ColWidth(5) = 1200
        .Text = "Lager EK-Wert"
        
        
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXLWD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub cmdHit_Click()
'    On Error GoTo LOKAL_ERROR
'
'    Screen.MousePointer = 11
'
'
'    MSHFLEX1.Visible = False
'
'    ErstelleMSHFLEXGo
'    liefstatErstellen
'    hitlisteVKWERT
'
'    bHit = True
'
'    bEW = False
'    bLW = False
'    bLWD = False
'    bGoPlus = False
'    bGo = False
'    bDetail = False
'    bDetailPlus = False
'
'    Screen.MousePointer = 0
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "cmdHit_Click"
'    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub liefstatErstellen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sLiefname As String
    Dim sLiefnr As String
    Dim sSQLlief As String
    Dim lLinr As Long
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    Dim rsLIEF As Recordset
    
    pbrZeit.Visible = True
    pbrZeit.Max = 400
    pbrZeit.Value = 100
    
    lblanzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblanzeige.Refresh
    
    sLiefnr = ""
    sLiefnr = Trim(Text1(2).Text)
   
    If sLiefnr <> "" Then
        If IsNumeric(sLiefnr) Then
            sSQLlief = " and Kassjour.LINR = " & sLiefnr & ""
        Else
            sSQLlief = ""
        End If
    Else
        sSQLlief = ""
    End If
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "liefstat", gdBase
    CreateTable "LIEFSTAT", gdBase
    
    pbrZeit.Value = 150
    sSQL = "INSERT into LIEFSTAT Select KASSJOUR.LINR, LISRT.LIEFBEZ , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis, Sum (Preis)- Sum(EKPR*Menge)as Ertrag "
    sSQL = sSQL & " from Kassjour, LISRT"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.LINR = LISRT.LINR "
    sSQL = sSQL & " and Kassjour.Filiale = " & gcFilNr
    sSQL = sSQL & " and Kassjour.UMS_OK = 'J' "
    sSQL = sSQL & sSQLlief
    
    sSQL = sSQL & " group BY  KASSJOUR.LINR, LISRT.LIEFBEZ "
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 200
    
    sSQL = "Update Liefstat set mindat = '" & Text1(0).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 250
    sSQL = "Update Liefstat set maxdat = '" & Text1(1).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 300
    
    sSQL = "Update Liefstat set auswahl = '" & sLiefname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    pbrZeit.Value = 400
    
    pbrZeit.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatErstellen"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub liefstatErstellenD(sLiefname As String, sLiefnr As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSQLlief As String
    Dim lLinr As Long
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    Dim rsLIEF As Recordset
    
    pbrZeit.Visible = True
    pbrZeit.Max = 400
    pbrZeit.Value = 100
    Screen.MousePointer = 11
    lblanzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblanzeige.Refresh
    

    
    If sLiefnr <> "" Then
        If sLiefnr = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            If IsNumeric(sLiefnr) Then
                sSQLlief = " and Kassjour.LINR = " & sLiefnr & ""
            Else
                sSQLlief = ""
            End If
        End If
    Else
    
        If sLiefname = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            sSQL = "Select Linr from lisrt"
            sSQL = sSQL & "  where liefbez = '" & sLiefname & "'"
            Set rsLIEF = gdBase.OpenRecordset(sSQL)
            
            If Not rsLIEF.EOF Then
            rsLIEF.MoveFirst
            
                If Not IsNull(rsLIEF!linr) Then
                    lLinr = rsLIEF!linr
                    sSQLlief = " and Kassjour.LINR = " & lLinr & ""
                Else
                    sSQLlief = ""
                End If
            End If
            rsLIEF.Close
        End If
    End If
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))

    loeschNEW "gode", gdBase
    
    pbrZeit.Value = 150
    sSQL = "Select kassjour.artnr "
    sSQL = sSQL & ", artikel.bezeich "
    sSQL = sSQL & ", artikel.LPZ "
    sSQL = sSQL & ", artikel.BESTAND "
    sSQL = sSQL & ", Sum(kassjour.Menge)as Anzahl "
    sSQL = sSQL & ", Sum (kassjour.Preis)as Umsatz "
    sSQL = sSQL & ", Sum(kassjour.EKPR*kassjour.Menge)as EinKPreis "
    sSQL = sSQL & ", Sum (Preis)- Sum(kassjour.EKPR*Menge)as Ertrag "
    sSQL = sSQL & ", min(adate) as mindat "
    sSQL = sSQL & ", max(adate) as maxdat "
    sSQL = sSQL & ", min(azeit) as auswahl "
    sSQL = sSQL & " into gode"
    sSQL = sSQL & " from Kassjour, artikel"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.artnr = artikel.artnr "
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & " group BY  KASSJOUR.artnr, artikel.bezeich, artikel.bestand, artikel.lpz "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    pbrZeit.Value = 200
    sSQL = "Update gode set mindat = " & cVon & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    pbrZeit.Value = 250
    sSQL = "Update gode set maxdat = " & cBis & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    pbrZeit.Value = 300
    sSQL = "Update gode set auswahl = '" & sLiefname & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    pbrZeit.Value = 400
    pbrZeit.Visible = False
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatErstellenD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub liefstatErstellenDPlus(sLiefname As String, sLiefnr As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSQLlief As String
    Dim lLinr As Long
    Dim cVon As String
    Dim cBis As String
    Dim cVonvj As String
    Dim cBisvj As String
    Dim lVon As Long
    Dim lBis As Long
    Dim rsLIEF As Recordset
    Dim cVorjahr            As String
    
    Dim iStufe As Integer
    
    iStufe = 0
    
    
    pbrZeit.Visible = True
    pbrZeit.Max = 1200
    pbrZeit.Value = 100
    Screen.MousePointer = 11
    lblanzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblanzeige.Refresh
    
    If sLiefnr <> "" Then
        If sLiefnr = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            If IsNumeric(sLiefnr) Then
                sSQLlief = " and Kassjour.LINR = " & sLiefnr & ""
            Else
                sSQLlief = ""
            End If
        End If
    Else
        If sLiefname = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            sSQL = "Select Linr from lisrt"
            sSQL = sSQL & "  where liefbez = '" & sLiefname & "'"
            Set rsLIEF = gdBase.OpenRecordset(sSQL)
            
            If Not rsLIEF.EOF Then
            rsLIEF.MoveFirst
            
                If Not IsNull(rsLIEF!linr) Then
                    lLinr = rsLIEF!linr
                    sSQLlief = " and Kassjour.LINR = " & lLinr & ""
                Else
                    sSQLlief = ""
                End If
            End If
            rsLIEF.Close
        End If
    End If
    pbrZeit.Value = 150
    
    iStufe = 1
    
    cVon = Format(Text1(0).Text, "DD.MM.YY")
    cBis = Format(Text1(1).Text, "DD.MM.YY")
    
    
    cVorjahr = Val(Right$(cVon, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    
    cVonvj = Left$(cVon, 6) & cVorjahr
    
    cVorjahr = Val(Right$(cBis, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    cBisvj = Left$(cBis, 6) & cVorjahr
    
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    lVon = DateValue(cVonvj)
    lBis = DateValue(cBisvj)
    
    cVonvj = Trim$(Str$(lVon))
    cBisvj = Trim$(Str$(lBis))
    
    iStufe = 2

    loeschNEW "tempo", gdBase
    
    loeschNEW "gode", gdBase
    
    iStufe = 3
    pbrZeit.Value = 200
    sSQL = "Select kassjour.artnr "
    sSQL = sSQL & ", artikel.bezeich "
    sSQL = sSQL & ", artikel.LPZ "
    sSQL = sSQL & ", artikel.BESTAND "
    sSQL = sSQL & ", Sum(kassjour.Menge)as Anzahl "
    sSQL = sSQL & ", Sum (kassjour.Preis)as Umsatz "
    sSQL = sSQL & ", Sum(kassjour.EKPR*kassjour.Menge)as EinKPreis "
    sSQL = sSQL & ", Sum (Preis)- Sum(kassjour.EKPR*Menge)as Ertrag "
    sSQL = sSQL & ", min(adate) as mindat "
    sSQL = sSQL & ", max(adate) as maxdat "
    sSQL = sSQL & ", min(azeit) as auswahl "
    
    sSQL = sSQL & ", Sum(Preis) as UmsatzVKVJ"
    sSQL = sSQL & ", Sum(Preis) as UmsatzEKVJ"
    
    sSQL = sSQL & ", Sum(Preis) as LagerVK"
    sSQL = sSQL & ", Sum(Preis) as LagerEK"
    sSQL = sSQL & ", Sum(Preis) as LagerLEK"
    
    sSQL = sSQL & ", artikel.ean as libesnr "
    
    sSQL = sSQL & " into gode"
    sSQL = sSQL & " from Kassjour, artikel"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.artnr = artikel.artnr "
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & " group BY  KASSJOUR.artnr, artikel.bezeich, artikel.bestand, artikel.lpz,artikel.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 4
    pbrZeit.Value = 250
    
    sSQL = "Update gode set mindat = " & cVon & ""
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 5
    pbrZeit.Value = 300
    
    sSQL = "Update gode set maxdat = " & cBis & ""
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 6
    pbrZeit.Value = 350
    
    sSQL = "Update gode set auswahl = '" & sLiefname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 7
    pbrZeit.Value = 400
    
    sSQL = "Update gode set UmsatzVKVJ = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 8
    pbrZeit.Value = 450
    
    sSQL = "Update gode set UmsatzEKVJ = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 9
    pbrZeit.Value = 500
    
    sSQL = "SELECT KASSJOUR.artnr, Sum(KASSJOUR.Preis) "
    sSQL = sSQL & " AS UmsatzVKVJ INTO TEMPO"
    sSQL = sSQL & " From KASSJOUR"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVonvj & " And " & cBisvj & " "
    sSQL = sSQL & " GROUP BY KASSJOUR.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 10
    pbrZeit.Value = 550
    
    sSQL = "UPDATE gode INNER JOIN TEMPO ON "
    sSQL = sSQL & " gode.artnr = TEMPO.artnr "
    sSQL = sSQL & " set gode.UmsatzVKVJ = TEMPO.UmsatzVKVJ"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 11
    pbrZeit.Value = 600
    
    sSQL = "DROP Table TEMPO"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 12
    pbrZeit.Value = 650
    
    sSQL = "SELECT KASSJOUR.artnr, Sum(EKPR*Menge) "
    sSQL = sSQL & " AS UmsatzEKVJ INTO TEMPO"
    sSQL = sSQL & " From KASSJOUR"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVonvj & " And " & cBisvj & " "
    sSQL = sSQL & " GROUP BY KASSJOUR.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 13
    pbrZeit.Value = 700
    
    sSQL = "UPDATE gode INNER JOIN TEMPO ON "
    sSQL = sSQL & " gode.artnr = TEMPO.artnr "
    sSQL = sSQL & " set gode.UmsatzEKVJ = TEMPO.UmsatzEKVJ"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 14
    pbrZeit.Value = 750
    
    sSQL = "DROP Table TEMPO"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 15
    pbrZeit.Value = 800
    
    sSQL = "SELECT Artikel.artnr, Sum(Artikel.KVKPR1*Artikel.Bestand) "
    sSQL = sSQL & " AS LagerVK INTO TEMPO"
    sSQL = sSQL & " From gode, artikel"
    sSQL = sSQL & " Where gode.artnr = artikel.artnr "
    sSQL = sSQL & " GROUP BY Artikel.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 16
    pbrZeit.Value = 850
    
    sSQL = "UPDATE gode INNER JOIN TEMPO ON "
    sSQL = sSQL & " gode.artnr = TEMPO.artnr "
    sSQL = sSQL & " set gode.LagerVK = TEMPO.LagerVK"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 17
    pbrZeit.Value = 900
    
    sSQL = "DROP Table TEMPO"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 18
    pbrZeit.Value = 950
    
    sSQL = "SELECT Artikel.artnr, Sum(Artikel.EKPR*Artikel.Bestand) "
    sSQL = sSQL & " AS LagerEK INTO TEMPO"
    sSQL = sSQL & " From gode, artikel"
    sSQL = sSQL & " Where gode.artnr = artikel.artnr "
    sSQL = sSQL & " GROUP BY Artikel.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 19
    pbrZeit.Value = 1000
    
    sSQL = "UPDATE gode INNER JOIN TEMPO ON "
    sSQL = sSQL & " gode.artnr = TEMPO.artnr "
    sSQL = sSQL & " set gode.LagerEK = TEMPO.LagerEK"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 20
    pbrZeit.Value = 1050
    
    sSQL = "DROP Table TEMPO"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 21
    pbrZeit.Value = 1100
    
    sSQL = "SELECT Artikel.artnr, Sum(Artikel.LEKPR*Artikel.Bestand) "
    sSQL = sSQL & " AS LagerLEK INTO TEMPO"
    sSQL = sSQL & " From gode, artikel"
    sSQL = sSQL & " Where gode.artnr = artikel.artnr "
    sSQL = sSQL & " GROUP BY Artikel.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 22
    pbrZeit.Value = 1150
    
    sSQL = "UPDATE gode INNER JOIN TEMPO ON "
    sSQL = sSQL & " gode.artnr = TEMPO.artnr "
    sSQL = sSQL & " set gode.LagerLEK = TEMPO.LagerLEK"
    gdBase.Execute sSQL, dbFailOnError
    
    iStufe = 23
    
    If Val(sLiefnr) > 0 Then
    
        sSQL = "UPDATE gode INNER JOIN ARTLIEF ON "
        sSQL = sSQL & " gode.artnr = ARTLIEF.artnr "
        sSQL = sSQL & " set gode.libesnr = ARTLIEF.libesnr"
        sSQL = sSQL & " where ARTLIEF.LINR = " & sLiefnr & ""
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    iStufe = 24
    
    pbrZeit.Value = 1200
    pbrZeit.Visible = False
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatErstellenDPlus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten. " & iStufe & " " & sLiefnr
    
    Fehlermeldung1
End Sub
Private Sub liefstatErstellenLWD(sLiefname As String, sLiefnr As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSQLlief As String
    Dim lLinr As Long
    Dim rsLIEF As Recordset
    
    pbrZeit.Visible = True
    pbrZeit.Max = 800
    pbrZeit.Value = 100
    Screen.MousePointer = 11
    lblanzeige.Caption = "Lagerwerte werden ermittelt..."
    lblanzeige.Refresh
    
    If sLiefnr <> "" Then
        If sLiefnr = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            sSQLlief = " and a.LINR = " & sLiefnr & ""
        End If
    Else
    
        If sLiefname = "alle Lieferanten" Then
            sSQLlief = ""
        Else
            sSQL = "Select Linr from lisrt"
            sSQL = sSQL & "  where liefbez = '" & sLiefname & "'"
            Set rsLIEF = gdBase.OpenRecordset(sSQL)
            
            If Not rsLIEF.EOF Then
            rsLIEF.MoveFirst
            
                If Not IsNull(rsLIEF!linr) Then
                    lLinr = rsLIEF!linr
                    sSQLlief = " and a.LINR = " & lLinr & ""
                Else
                    sSQLlief = ""
                End If
            End If
            rsLIEF.Close
        End If
    End If
    pbrZeit.Value = 150

    loeschNEW "gode", gdBase
    CreateTable "GODE", gdBase
    
    pbrZeit.Value = 200
    
    If Option1(0).Value = True Then 'SEK
        sSQL = "INSERT into GODE Select a.artnr, a.BEZEICH , a.EAN , a.LPZ, Sum(a.BESTAND) as BESTAND "
        sSQL = sSQL & ", Sum(a.KVKPR1*a.BESTAND) as LagerVK "
        sSQL = sSQL & ", Sum(a.EKPR*a.BESTAND) as LagerEK "
        sSQL = sSQL & ", a.LIBESNR "
        sSQL = sSQL & " from ARTIKEL a "
        sSQL = sSQL & " where a.Bestand > 0 "
        sSQL = sSQL & sSQLlief
        sSQL = sSQL & " group BY a.artnr, a.BEZEICH, a.LPZ, a.libesnr , a.EAN"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update gode set BGrund = 'Schnitteinkaufswert' "
        gdBase.Execute sSQL, dbFailOnError
        
    ElseIf Option1(1).Value = True Then 'lEK
    
        sSQL = "INSERT into GODE Select a.artnr, a.BEZEICH , a.EAN , a.LPZ, Sum(a.BESTAND) as BESTAND "
        sSQL = sSQL & ", Sum(a.KVKPR1 * a.BESTAND) as LagerVK "
        sSQL = sSQL & ", Sum(a.LEKPR * a.BESTAND) as LagerEK "
        sSQL = sSQL & ", A.LIBESNR "
        sSQL = sSQL & " from  ARTIKEL A inner join artlief B on B.artnr = A.artnr "
        sSQL = sSQL & " and a.linr = b.linr "
        sSQL = sSQL & " where a.Bestand > 0 "
        sSQL = sSQL & sSQLlief
        sSQL = sSQL & " group BY A.artnr, A.BEZEICH, A.LPZ, A.libesnr, A.EAN "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update gode set BGrund = 'Listeneinkaufswert' "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    
    sSQL = "Update gode set auswahl = '" & sLiefname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    pbrZeit.Value = 800
    pbrZeit.Visible = False
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatErstellenLWD"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub liefstatLWErstellen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sLiefname As String
    Dim sLiefnr As String
    Dim sSQLlief As String
    Dim lLinr As Long
    Dim rsLIEF As Recordset
    
    pbrZeit.Visible = True
    pbrZeit.Max = 800
    pbrZeit.Value = 100
    
    lblanzeige.Caption = "Lagerwerte werden ermittelt..."
    lblanzeige.Refresh
    
    sLiefnr = ""
    sLiefnr = Trim(Text1(2).Text)
    
    
    loeschNEW "Lieflw", gdBase
    CreateTable "LIEFLW", gdBase
    loeschNEW "Tempo", gdBase
    
    
    
    If Option1(0).Value = True Then 'Schnittek
    
        If sLiefnr <> "" Then
            sSQLlief = " and LINR = " & sLiefnr & ""
        Else
            sSQLlief = ""
        End If
    
        loeschNEW "ArtTemp", gdBase
    
        sSQL = "select * into arttemp from artikel "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Update arttemp inner join artikel on arttemp.artnr = artikel.artnr "
        sSQL = sSQL & " set arttemp.ekpr = artikel.lekpr where arttemp.ekpr = 0 "
        gdBase.Execute sSQL, dbFailOnError
    
    
        sSQL = "INSERT into LIEFLW Select LINR, Sum(arttemp.BESTAND) as BESTAND "
        sSQL = sSQL & ", Sum(KVKPR1* arttemp.BESTAND) as LagerVK"
        sSQL = sSQL & ", Sum(EKPR* arttemp.BESTAND) as LagerEK"
        sSQL = sSQL & " from arttemp "
        sSQL = sSQL & " Where arttemp.Bestand > 0  "
        sSQL = sSQL & sSQLlief
        sSQL = sSQL & " group BY arttemp.LINR "
        gdBase.Execute sSQL, dbFailOnError
    
        loeschNEW "ArtTemp", gdBase
        
        sSQL = "Update LIEFLW set BGrund = 'Schnitteinkaufswert' "
        gdBase.Execute sSQL, dbFailOnError
    
    
    ElseIf Option1(1).Value = True Then 'Listenek
    
        If sLiefnr <> "" Then
            sSQLlief = " and a.LINR = " & sLiefnr & ""
        Else
            sSQLlief = ""
        End If
    
        sSQL = "INSERT into LIEFLW Select a.LINR, Sum(a.BESTAND) as BESTAND "
        sSQL = sSQL & ", Sum(a.KVKPR1* a.BESTAND) as LagerVK"
        sSQL = sSQL & ", Sum(b.lEKPR* a.BESTAND) as LagerEK"
        sSQL = sSQL & " from  ARTIKEL A inner join artlief B on B.artnr = A.artnr "
        sSQL = sSQL & " and a.linr = b.linr "
        sSQL = sSQL & " Where a.Bestand > 0  "
        sSQL = sSQL & sSQLlief
        sSQL = sSQL & " group BY a.LINR "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update LIEFLW set BGrund = 'Listeneinkaufswert' "
        gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    sSQL = "Update LiefLW set auswahl = '" & sLiefname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LiefLW inner join lisrt on lieflw.linr = lisrt.linr set lieflw.LIEFBEZ = lisrt.liefbez"
    gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatLWErstellen"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub liefstatEWErstellen()
    On Error GoTo LOKAL_ERROR

    Dim cVon As String
    Dim cBis As String
    
    Dim lVon As Long
    Dim lBis As Long
    
    Dim cVJVon As String
    Dim cVJBis As String
    
    Dim lVJVon As Long
    Dim lVJBis As Long
    
    Dim sSQL As String
    Dim sLiefname As String
    Dim sLiefnr As String
    Dim sSQLlief As String
    Dim lLinr As Long
    Dim rsLIEF As Recordset
    Dim cVorjahr As String
    
'    cVon = Text1(0).Text
'    cBis = Text1(1).Text
    
    cVon = Format(Text1(0).Text, "DD.MM.YY")
    cBis = Format(Text1(1).Text, "DD.MM.YY")
    
    cVorjahr = Val(Right$(cVon, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    
    cVJVon = Left$(cVon, 6) & cVorjahr
    
    cVorjahr = Val(Right$(cBis, 2)) - 1
    If Len(cVorjahr) = 1 Then
        cVorjahr = "0" & cVorjahr
    End If
    cVJBis = Left$(cBis, 6) & cVorjahr
    
'    cVJVon = Left(Text1(0).Text, 7) & Right(Text1(0).Text, 1) - 1
'    cVJBis = Left(Text1(1).Text, 7) & Right(Text1(1).Text, 1) - 1
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    lVJVon = DateValue(cVJVon)
    lVJBis = DateValue(cVJBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    cVJVon = Trim$(Str$(lVJVon))
    cVJBis = Trim$(Str$(lVJBis))
    
    lblanzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblanzeige.Refresh
    
    sLiefnr = ""
    sLiefnr = Trim(Text1(2).Text)
    
    If sLiefnr <> "" Then
        If IsNumeric(sLiefnr) Then
            sSQLlief = " and Kassjour.LINR = " & sLiefnr & ""
        Else
            sSQLlief = ""
        End If
    Else
        sSQLlief = ""
    End If
    
    loeschNEW "LiefEW", gdBase
    CreateTable "LIEFEW", gdBase

    sSQL = "INSERT into LIEFEW Select Kassjour.LINR, LISRT.LIEFBEZ, Sum(Kassjour.Preis) as Umsatz "
    sSQL = sSQL & " , Kassjour.adate "
    sSQL = sSQL & " from Kassjour, LISRT "
    sSQL = sSQL & " Where Kassjour.LINR = LISRT.LINR "
    sSQL = sSQL & " and Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.Filiale = " & gcFilNr
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & " group BY Kassjour.LINR, LISRT.LIEFBEZ , Kassjour.adate "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LiefEW set von = " & cVon & ""
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update LiefEW set bis = " & cBis & ""
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LiefEW set vrvon = " & cVJVon & ""
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update LiefEW set vrbis = " & cVJBis & ""
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "liefewp", gdBase
    CreateTable "LIEFEWP", gdBase
    
    sSQL = "INSERT into LIEFEWP Select Kassjour.LINR, Sum(Kassjour.Preis) as Umsatz"
    sSQL = sSQL & " , Sum(Kassjour.menge) as menge,  Sum(Kassjour.vkpr*kassjour.menge) as gVJZRVK"
    sSQL = sSQL & " ,  Sum(Kassjour.EKPR*kassjour.Menge) as gVRZREK "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " Where "
    sSQL = sSQL & " Kassjour.ADATE Between " & cVJVon & " And " & cVJBis & " "
    sSQL = sSQL & " and Kassjour.Filiale = " & gcFilNr
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & " group BY Kassjour.LINR "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "LiefEWZ", gdBase
    CreateTable "LIEFEWZ", gdBase
    sSQL = "Insert into LIEFEWZ Select Kassjour.LINR "
    sSQL = sSQL & " , Sum(Kassjour.menge) as menge,  Sum(Kassjour.vkpr*kassjour.menge) as gZRVK"
    sSQL = sSQL & " ,  Sum(Kassjour.EKPR*kassjour.Menge) as gZREK "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " Where "
    sSQL = sSQL & " Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.Filiale = " & gcFilNr
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & " group BY Kassjour.LINR "
    gdBase.Execute sSQL, dbFailOnError
    
    If sSQLlief = "" Then
        reportbildschirm "ewal", "aWKLaua"
    Else
        reportbildschirm "ew", "aWKLaub"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatEWErstellen"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub hitlisteVKWERT()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim cFeld As String
    Dim dEkpr As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sLinr As String
    Dim iTop As Integer
    Dim sHit As String
    
    loeschNEW "Liefhit", gdBase
    
    If optE.Value Then
        sHit = " Liefstat.ERTRAG DESC"
    ElseIf optU.Value Then
        sHit = " Liefstat.Umsatz DESC"
    ElseIf optA.Value Then
        sHit = " Liefstat.Anzahl DESC"
    End If
    
    iTop = 10
    If opt5.Value Then
        iTop = 5
    Else
        iTop = 10
    End If
    
    sSQL = "SELECT TOP " & iTop & " Liefstat.LINR, Liefstat.LIEFBEZ, Liefstat.UMSATZ "
    sSQL = sSQL & " , Liefstat.Anzahl , Liefstat.EINKPREIS, Liefstat.Ertrag "
    sSQL = sSQL & " , Liefstat.mindat , Liefstat.maxdat "
    sSQL = sSQL & " into Liefhit from Liefstat order by " & sHit
    gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("Liefhit", dbOpenTable)
    
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 1
    If Not rs.EOF Then
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!linr) Then
                sLinr = rs!linr
            Else
                sLinr = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sLinr
            
            If Not IsNull(rs!LIEFBEZ) Then
            cFeld = rs!LIEFBEZ
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            MSHFLEX1.Col = 3
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 4
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            

            If Not IsNull(rs!ERTRAG) Then
                dErtrag = rs!ERTRAG
            Else
                dErtrag = "0"
            End If
            
            MSHFLEX1.Col = 5
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    pbrZeit.Visible = False
    lblanzeige.Caption = "Fertig"
    lblanzeige.Refresh
    
    If lrow = 0 Then
        lblanzeige.Caption = "Keine Daten gefunden"
        lblanzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.RowHeight(1) = 0
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
  
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "hitlisteVKWERT"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sOrder As String
    Dim sSQL As String
    
    If bHit Then
    
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lblanzeige
        
        If optE.Value Then
            reportbildschirm "hite", "aWKLauc"
        ElseIf optA.Value Then
            reportbildschirm "hita", "aWKLaud"
        ElseIf optU.Value Then
            reportbildschirm "hitu", "aWKLaue"
        End If
        
        anzeige "normal", "", lblanzeige
        
    ElseIf bGoPlus Then
    
        anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lblanzeige
    
'        loeschNEW "LItot", gdBase
'        sSQL = " Select * into Litot from LIEFPLUS"
'        gdBase.Execute sSQL, dbFailOnError
        
'        loeschNEW "LIEFPLUS", gdBase
'        CreateTable "LIEFPLUS", gdBase
        
'        If Option5(0).Value = True Then
'            sOrder = " Order by Umsatz desc" 'Umsatz
'        ElseIf Option5(1).Value = True Then
'            sOrder = " Order by NSPAkt desc" 'NSPAkt
'        ElseIf Option5(2).Value = True Then
'            sOrder = " Order by PENANTEILST desc" 'Penner
'        ElseIf Option5(6).Value = True Then
'            sOrder = " Order by Liefbez" 'Liefbez
'        End If
        
'        sSQL = "Insert into LIEFPLUS Select * from Litot " & sOrder
'        gdBase.Execute sSQL, dbFailOnError
'
'        loeschNEW "LItot", gdBase
        reportbildschirm "", "ZaGoPlus"
        anzeige "normal", "", lblanzeige
    
    ElseIf bLW Then
        reportbildschirm "lw", "aWKLauh"
    ElseIf bLWD Then
    
        If Modul6.FindFile(gcDBPfad, "aWKLaus.rpt") Then
            reportbildschirm "lwd", "aWKLaus"
        Else
            reportbildschirm "lwd", "aWKLaui"
        End If
        
    ElseIf bDetail Then
        reportbildschirm "liefd", "aWKLauj"
    ElseIf bDetailPlus Then
        If Option5(5).Value = True Then
            sOrder = " Order by Libesnr" 'Lieferantenbestellnummer
        ElseIf Option5(4).Value = True Then
            sOrder = " Order by Bezeich" 'Bezeichnung
        ElseIf Option5(3).Value = True Then
            sOrder = " Order by lpz " 'Linie
        End If
        
        loeschNEW "godetmp", gdBase
        
        sSQL = "Select * into godetmp from gode " & sOrder
        gdBase.Execute sSQL, dbFailOnError
        
        loeschNEW "gode", gdBase
        
        sSQL = "Select * into gode from godetmp "
        gdBase.Execute sSQL, dbFailOnError
        
        loeschNEW "godetmp", gdBase
        
        reportbildschirm "liefdp", "aWKLauk"
    Else
        Exit Sub
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPrint_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR

    If MSHFLEX1.Row > 1 Then
        Command1.Visible = False
    Else
        Exit Sub
    End If

    Dim lrow As Long
    Dim sLiefname As String
    Dim sLiefnr As String

    If bGo Then
        
        lrow = MSHFLEX1.Row

        sLiefnr = MSHFLEX1.TextMatrix(lrow, 0)
        sLiefname = MSHFLEX1.TextMatrix(lrow, 1)

        ErstelleMSHFLEXGoD
        liefstatErstellenD sLiefname, sLiefnr
        GoAuswertungD

        bDetail = True

        bLWD = False
        bGo = False
        bGoPlus = False
        bHit = False
        bEW = False
        bLW = False
        bDetailPlus = False


        If Not lblanzeige.Caption = "Keine Daten gefunden" Then
            lblanzeige.Caption = sLiefname & ", " & MSHFLEX1.Rows - 1 & " verschiedene Artikel wurden in diesem Zeitraum verkauft."
            lblanzeige.Refresh
            cmdBack.Visible = True
        End If
        
    ElseIf bGoPlus Then
    
        Frame3.Visible = True
        lrow = MSHFLEX1.Row

        sLiefnr = MSHFLEX1.TextMatrix(lrow, 0)
        sLiefname = MSHFLEX1.TextMatrix(lrow, 1)

        ErstelleMSHFLEXGoDPlus
        liefstatErstellenDPlus sLiefname, sLiefnr
        GoAuswertungDPlus

        bDetailPlus = True

        bLWD = False
        bDetail = False
        bGo = False
        bGoPlus = False
        bHit = False
        bEW = False
        bLW = False

        If Not lblanzeige.Caption = "Keine Daten gefunden" Then
            lblanzeige.Caption = sLiefname & ", " & MSHFLEX1.Rows - 1 & " verschiedene Artikel wurden in diesem Zeitraum verkauft."
            lblanzeige.Refresh
            cmdBack.Visible = True
        End If
    ElseIf bLW Then
        lrow = MSHFLEX1.Row

        sLiefnr = MSHFLEX1.TextMatrix(lrow, 0)
        sLiefname = MSHFLEX1.TextMatrix(lrow, 1)

        ErstelleMSHFLEXLWD
        liefstatErstellenLWD sLiefname, sLiefnr
        GoAuswertungLWD

        bLWD = True

        bDetailPlus = False
        bDetail = False
        bGo = False
        bGoPlus = False
        bHit = False
        bEW = False
        bLW = False

        If Not lblanzeige.Caption = "Keine Daten gefunden" Then
            lblanzeige.Caption = sLiefname & ", " & MSHFLEX1.Rows - 1 & " Artikel, die heute im Bestand sind."
            lblanzeige.Refresh
            cmdBack.Visible = True
        End If
    End If
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        Case 0
            gsZSpalte = "LINR"
            gstab = "LIEFSTAT"
            frmWKL36.Show 1
            'fertig
        Case 6
            Text1_KeyUp 6, vbKeyF2, 0
        Case 11
            gsHelpstring = "Lieferantenstatistik"
            frmWKL110.Show 1
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(1).SetFocus
            
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub positionierenwklau()
    On Error GoTo LOKAL_ERROR
    
    MSHFLEX1.Height = 5655
    MSHFLEX1.Left = 120
    MSHFLEX1.Top = 1200
    MSHFLEX1.Width = 11655

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwklau"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    positionierenwklau
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    bLW = False
    bEW = False
    bHit = False
    bGoPlus = False
    bGo = False
    bDetail = False
    bDetailPlus = False
    bLWD = False
    
    optqp.Value = True
    
    Option1_Click 2
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    loeschNEW "liefstat", gdBase
    loeschNEW "lieftemp", gdBase
    loeschNEW "gode", gdBase
    loeschNEW "liefplus", gdBase
    loeschNEW "te12", gdBase
    loeschNEW "tempo", gdBase
    loeschNEW "lieflw", gdBase
    loeschNEW "liefew", gdBase
    loeschNEW "liefewp", gdBase
    loeschNEW "liefewz", gdBase
    loeschNEW "liefhit", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LiefplusSort()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSort As String
    
    If sSortklick <> "" Then
        Select Case sSortklick
            Case Is = "Umsatz (VK)"
                sSort = "Umsatz desc"
            Case Is = "Umsatz (EK)"
                sSort = "Einkpreis desc"
            Case Is = "Ertrag"
                sSort = "Ertrag desc"
            Case Is = "verk.Artikel"
                sSort = "Anzahl desc"
            Case Is = "Umsatz (VK) VJ ZR"
                sSort = "UmsatzVKVJ desc"
            Case Is = "Umsatz (EK) VJ ZR"
                sSort = "UmsatzEKVJ desc"
            Case Is = "Lager VK-Wert"
                sSort = "Lagervk desc"
            Case Is = "Lager EK-Wert"
                sSort = "LagerEk desc"
            Case Is = "Lager LEK-Wert"
                sSort = "LagerLEk desc"
            Case Is = "Lieferantenname"
                sSort = "Liefbez asc"
            Case Else
                sSort = "Liefbez"
        End Select
    Else
        
        sSort = "Liefbez"
        
    End If
    
    loeschNEW "lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from Liefplus "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "liefplus", gdBase
    
    sSQL = "Select * into Liefplus from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LiefplusSort"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DetailSort()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSort As String
    
    If sSortklick <> "" Then
        Select Case sSortklick
            Case Is = "Umsatz (VK)"
                sSort = "Umsatz desc"
            Case Is = "Umsatz (EK)"
                sSort = "Einkpreis desc"
            Case Is = "Ertrag"
                sSort = "Ertrag desc"
            Case Is = "verk.Artikel"
                sSort = "Anzahl desc"
            Case Is = "Umsatz (VK) VJ ZR"
                sSort = "UmsatzVKVJ desc"
            Case Is = "Umsatz (EK) VJ ZR"
                sSort = "UmsatzEKVJ desc"
            Case Is = "Lager VK-Wert"
                sSort = "Lagervk desc"
            Case Is = "Lager EK-Wert"
                sSort = "LagerEk desc"
            Case Is = "Lager LEK-Wert"
                sSort = "LagerLEk desc"
            Case Is = "Bestand"
                sSort = "Bestand desc"
            Case Is = "Linie"
                sSort = "LPZ asc"
            Case Is = "Artikelbezeichnung"
                sSort = "bezeich asc"
            Case Else
                sSort = "bezeich asc"
        End Select
    Else
        
        sSort = "bezeich asc"
        
    End If
        
    loeschNEW "lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from gode "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "gode", gdBase
    
    sSQL = "Select * into gode from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DetailSort"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LiefLWSort()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSort As String
    
    If sSortklick <> "" Then
        Select Case sSortklick
            Case Is = "Umsatz (VK)"
                sSort = "Umsatz desc"
            Case Is = "Umsatz (EK)"
                sSort = "Einkpreis desc"
            Case Is = "Ertrag"
                sSort = "Ertrag desc"
            Case Is = "verk.Artikel"
                sSort = "Anzahl desc"
            Case Is = "Umsatz (VK) VJ ZR"
                sSort = "UmsatzVKVJ desc"
            Case Is = "Umsatz (EK) VJ ZR"
                sSort = "UmsatzEKVJ desc"
            Case Is = "Lager VK-Wert"
                sSort = "Lagervk desc"
            Case Is = "Lager EK-Wert"
                sSort = "LagerEk desc"
            Case Is = "Lager LEK-Wert"
                sSort = "LagerLEk desc"
            Case Is = "Lieferantenname"
                sSort = "Liefbez asc"
            Case Else
                sSort = "Liefbez"
        End Select
    Else
        sSort = "Liefbez"
    End If

    loeschNEW "lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from Lieflw "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Lieflw", gdBase
    
    sSQL = "Select * into Lieflw from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LiefLWSort"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LiefstatSort()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sSort As String
    
    If sSortklick <> "" Then
        Select Case sSortklick
            Case Is = "Umsatz (VK)"
                sSort = "Umsatz desc"
            Case Is = "Umsatz (EK)"
                sSort = "Einkpreis desc"
            Case Is = "Ertrag"
                sSort = "Ertrag desc"
            Case Is = "verk.Artikel"
                sSort = "Anzahl desc"
            Case Is = "Lieferantenname"
                sSort = "Liefbez asc"
            
            Case Else
                sSort = "Liefbez"
        End Select
    Else
        
        sSort = "Liefbez"
        
    End If
    
    loeschNEW "lieftemp", gdBase
    
    sSQL = "Select * into Lieftemp from Liefstat "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Liefstat", gdBase
    
    sSQL = "Select * into Liefstat from Lieftemp order by " & sSort
    gdBase.Execute sSQL, dbFailOnError
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LiefstatSort"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MSHFLEX1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim lcol As Long
    Dim sSortKrit As String
    Dim sSpalteUeberschrift As String
    Dim lrow As Long
    Dim sSQL As String
    
    If MSHFLEX1.Row > 1 Then
        
    Else
            
        lcol = MSHFLEX1.Col
        lrow = 0
        sSpalteUeberschrift = MSHFLEX1.TextMatrix(lrow, lcol)
        Select Case sSpalteUeberschrift
            Case Is = "LiNr"
                sSortKrit = " order by  LINR"
            Case Is = "Lieferant"
                sSortKrit = " order by  LINR"
            Case Is = "Menge"
                sSortKrit = " order by  ANZAHL"
            Case Is = "Menge VJ"
                sSortKrit = " order by  ANZAHLVJ"
            Case Is = "Ums akt"
                sSortKrit = " order by  UMSATZ"
            Case Is = "Ums VJ ZR"
                sSortKrit = " order by  UmsatzVKVJ"
            Case Is = "Diff Ums Eur"
                sSortKrit = " order by  DIFFUMSEUR"
            Case Is = "Diff Ums %"
                sSortKrit = " order by  DIFFUMSPROZ"
            Case Is = "Ums akt EK"
                sSortKrit = " order by  EINKPREIS"
            Case Is = "Ums VJ ZR EK"
                sSortKrit = " order by  UmsatzEKVJ"
            Case Is = "Diff Ums Eur EK"
                sSortKrit = " order by  DIFFUMSEUREK"
            Case Is = "Diff Ums % EK"
                sSortKrit = " order by  DIFFUMSPROZEK"
            Case Is = "Ertrag akt"
                sSortKrit = " order by  ERTRAGakt"
            Case Is = "Ertrag VJ"
                sSortKrit = " order by  ERTRAGvj"
            Case Is = "NSP akt"
                sSortKrit = " order by  NSPakt"
            Case Is = "NSP VJ"
                sSortKrit = " order by  NSPvj"
            Case Is = "LAGER(Stück)"
                sSortKrit = " order by  LAGERST"
            Case Is = "Penner(Stück)"
                sSortKrit = " order by  PENNERST"
            Case Is = "Panteil Stück in %"
                sSortKrit = " order by  PENANTEILST"
            Case Is = "LAGER(SEK)"
                sSortKrit = " order by  LAGERWSEK"
            Case Is = "Penner(SEK)"
                sSortKrit = " order by  PENNERWSEK"
            Case Is = "Panteil SEK in %"
                sSortKrit = " order by  PENANTEILSEK"
        
        End Select
        
        
        loeschNEW "LItot", gdBase
        sSQL = "select * into LItot from LIEFPLUS " & sSortKrit
        
        If byteSortReihen = 1 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " asc"
            End If
            byteSortReihen = 2
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 1
        ElseIf byteSortReihen = 2 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " desc"
            End If
            byteSortReihen = 1
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 2
        End If
        
        gdBase.Execute sSQL
        
        loeschNEW "LIEFPLUS", gdBase
    
        sSQL = "select * into LIEFPLUS from LItot "
        gdBase.Execute sSQL, dbFailOnError
        loeschNEW "LItot", gdBase
    

    
    
    
''        sortierenHGrid MSHFLEX1
        
        
        
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub optA_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame7.Visible = True
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optA_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub optD_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1(2).Text = ""
    optLW.Value = False
    Frame8.Visible = True
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame1.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optD_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optE_Click()
    On Error GoTo LOKAL_ERROR

    Frame7.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optE_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optEw_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)) - 1
    Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
    
    Frame1.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optEw_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optH_Click()
    On Error GoTo LOKAL_ERROR

    Text1(2).Text = ""
    Frame5.Visible = True
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
    Frame1.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optH_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 2    'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(1).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        Else
                            Text1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        End If
                    
                    Case Else
                        Text1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
                
        Case Is = 5     'ak monat
            Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
            
        Case Is = 4     'vorjahr
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Text1(1).Text = Format("31.12." & Year(Now) - 1, "DD.MM.YY")
        Case Is = 3     'ak jahr
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
   
End Sub

Private Sub optLW_Click()
On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optLW_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optqp_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1(2).Text = ""
    Frame5.Visible = False
    Frame6.Visible = True
    Frame7.Visible = False
    Frame8.Visible = False
    Frame1.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optqp_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optU_Click()
    On Error GoTo LOKAL_ERROR

    Frame7.Visible = True
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optU_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 2 Then
        LiefKuerzelAufloesung Label1(10), Text1(2)
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
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
        
        
        gF2Prompt.cFeld = "LINR"
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            Text1(2).Text = gF2Prompt.cWahl
            Label1(10).Caption = gF2Prompt.cWert
            
        End If
        Text1(2).SetFocus
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optU_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Len(Text1(0).Text) = 0 Then
        lblanzeige.Caption = "Geben Sie ein Anfangsdatum ein!."
        lblanzeige.Refresh
        Text1(0).SetFocus
    End If
    
    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    MSHFLEX1.Visible = False
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


