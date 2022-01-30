VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK25a 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Umsatzstatistik"
   ClientHeight    =   8625
   ClientLeft      =   1815
   ClientTop       =   2130
   ClientWidth     =   11910
   Icon            =   "frmWK25a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Caption         =   "Umsatzinformationen"
      Height          =   1335
      Left            =   720
      TabIndex        =   60
      Top             =   3000
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   61
         Top             =   360
         Width           =   6615
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Left            =   7080
         TabIndex        =   62
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
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
            Size            =   14.25
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Umsatzinformationen,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   64
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   63
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1200
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   1920
   End
   Begin sevCommand3.Command cmdPlus 
      Height          =   310
      Left            =   10800
      TabIndex        =   29
      ToolTipText     =   "Starten Sie hier die Anzeige"
      Top             =   6720
      Width           =   450
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   ">"
      PictureAlign    =   2
      UseMaskColor    =   -1  'True
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdMinus 
      Height          =   310
      Left            =   480
      TabIndex        =   28
      ToolTipText     =   "Starten Sie hier die Anzeige"
      Top             =   6720
      Width           =   450
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "<"
      PictureAlign    =   2
      UseMaskColor    =   -1  'True
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdQuick 
      Height          =   420
      Left            =   10920
      TabIndex        =   8
      ToolTipText     =   "Starten Sie hier die Anzeige"
      Top             =   240
      Width           =   480
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Go"
      PictureAlign    =   2
      UseMaskColor    =   -1  'True
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Index           =   1
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Left            =   9480
      TabIndex        =   5
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
         Size            =   14.25
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   9375
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "eing. Zeitraum mit Monatsvergleich"
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
         Left            =   2400
         TabIndex        =   67
         ToolTipText     =   "Hier sehen Sie die Umsätze des aktuellen Monats tagesgenau."
         Top             =   480
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "bis Heute"
         Height          =   195
         Left            =   7320
         TabIndex        =   66
         Top             =   720
         Value           =   1  'Aktiviert
         Width           =   1215
      End
      Begin VB.OptionButton optWeek 
         BackColor       =   &H00C0C000&
         Caption         =   "Wochenansicht"
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
         Left            =   120
         TabIndex        =   65
         Top             =   480
         Width           =   2175
      End
      Begin sevCommand3.Command Command2 
         Height          =   310
         Index           =   1
         Left            =   7320
         TabIndex        =   56
         ToolTipText     =   "Starten Sie hier die Anzeige"
         Top             =   360
         Width           =   810
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "<"
         PictureAlign    =   2
         UseMaskColor    =   -1  'True
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   310
         Index           =   0
         Left            =   6480
         TabIndex        =   55
         ToolTipText     =   "Starten Sie hier die Anzeige"
         Top             =   360
         Width           =   810
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "<"
         PictureAlign    =   2
         UseMaskColor    =   -1  'True
         Version3        =   -1  'True
      End
      Begin VB.OptionButton optZRVJ 
         BackColor       =   &H00C0C000&
         Caption         =   "eing. Zeitraum mit Jahresvergleich"
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
         Left            =   2400
         TabIndex        =   54
         ToolTipText     =   "Hier sehen Sie die Umsätze des aktuellen Monats tagesgenau."
         Top             =   0
         Width           =   3975
      End
      Begin VB.OptionButton optZR 
         BackColor       =   &H00C0C000&
         Caption         =   "eingegebener Zeitraum"
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
         Left            =   2400
         TabIndex        =   13
         ToolTipText     =   "Hier sehen Sie die Umsätze des aktuellen Monats tagesgenau."
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optAkt 
         BackColor       =   &H00C0C000&
         Caption         =   "aktueller Monat"
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
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "Hier sehen Sie die Umsätze des aktuellen Monats tagesgenau."
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.OptionButton optqp 
         BackColor       =   &H00C0C000&
         Caption         =   "Monatsansicht"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optq 
         BackColor       =   &H00C0C000&
         Caption         =   "Jahresansicht"
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
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   2055
      End
      Begin VB.OptionButton optD 
         BackColor       =   &H00C0C000&
         Caption         =   "Tagesansicht"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Jahresvergleich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   57
         Top             =   30
         Width           =   1935
      End
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   20
      Left            =   8160
      TabIndex        =   58
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   21
      Left            =   10320
      TabIndex        =   59
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz in Euro: Klicken Sie auf die Umsatzzahlen, um Infomationen zu hinterlegen"
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
      Index           =   9
      Left            =   480
      TabIndex        =   53
      Top             =   960
      Width           =   11295
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Undurchsichtig
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   375
      Index           =   8
      Left            =   1920
      Top             =   6720
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Undurchsichtig
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   375
      Index           =   7
      Left            =   6600
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz in Euro / aktuelles Jahr"
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
      Index           =   8
      Left            =   2520
      TabIndex        =   52
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Umsatz in Euro / Vorjahr"
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
      Index           =   7
      Left            =   7080
      TabIndex        =   51
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   10560
      TabIndex        =   50
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   9000
      TabIndex        =   49
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   48
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   47
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   46
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   45
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   44
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   10575
      MouseIcon       =   "frmWK25a.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   43
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   9015
      MouseIcon       =   "frmWK25a.frx":074C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   42
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7455
      MouseIcon       =   "frmWK25a.frx":0A56
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   41
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5895
      MouseIcon       =   "frmWK25a.frx":0D60
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   40
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4335
      MouseIcon       =   "frmWK25a.frx":106A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   39
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2775
      MouseIcon       =   "frmWK25a.frx":1374
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   38
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1215
      MouseIcon       =   "frmWK25a.frx":167E
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   37
      Top             =   5280
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   6
      Left            =   10575
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   5
      Left            =   9015
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   4
      Left            =   7455
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   3
      Left            =   5895
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   2
      Left            =   4335
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   1
      Left            =   2775
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   0
      Left            =   1215
      Top             =   5760
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   9840
      TabIndex        =   36
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8280
      TabIndex        =   35
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   34
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   33
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   32
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   31
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   30
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   9840
      MouseIcon       =   "frmWK25a.frx":1988
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   27
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   8280
      MouseIcon       =   "frmWK25a.frx":1C92
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   26
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6720
      MouseIcon       =   "frmWK25a.frx":1F9C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   25
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5160
      MouseIcon       =   "frmWK25a.frx":22A6
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   24
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3600
      MouseIcon       =   "frmWK25a.frx":25B0
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   23
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      MouseIcon       =   "frmWK25a.frx":28BA
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   22
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      MouseIcon       =   "frmWK25a.frx":2BC4
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   21
      Top             =   5280
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   6
      Left            =   9840
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   5
      Left            =   8280
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   4
      Left            =   6720
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   3
      Left            =   5160
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   2
      Left            =   3600
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   1
      Left            =   2040
      Top             =   5760
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   15
      Index           =   0
      Left            =   480
      Top             =   5760
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   9840
      TabIndex        =   20
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   8280
      TabIndex        =   19
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Left            =   6720
      TabIndex        =   18
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   5160
      TabIndex        =   17
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   16
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   2040
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Label2"
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
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
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
      Left            =   240
      TabIndex        =   11
      Top             =   7200
      Width           =   10815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   11400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "von:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   10
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "bis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   9
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Umsatzstatistik"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmWK25a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPrueD  As Integer
Dim gitop           As Integer



Private Sub cmdMinus_Click()
    On Error GoTo LOKAL_ERROR

    Dim lDatum As Long
    lDatum = DateValue(Label2(6).Caption)
    ZeigeUmsatzgrafik lDatum

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMinus_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub cmdMinus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    Timer1.Enabled = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMinus_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub cmdMinus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    Timer1.Enabled = False
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMinus_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub cmdPlus_Click()
    On Error GoTo LOKAL_ERROR


    Dim lDatum As Long

    lDatum = DateValue(Label2(6).Caption)
    lDatum = lDatum + 2

    ZeigeUmsatzgrafik lDatum

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPlus_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub cmdPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR

    Timer2.Enabled = True
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPlus_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub cmdPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR

    Timer2.Enabled = False

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPlus_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub cmdQuick_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim rs As Recordset
    
    Dim Datum As Date
    Datum = Date
    
    lblanzeige.ForeColor = vbBlue
    lblanzeige.Caption = "Daten werden ermittelt, bitte warten..."
    lblanzeige.Refresh
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Screen.MousePointer = 11
    
    If optq.Value = True Then 'Jahresansicht
    
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        Text1(0).Text = Left(Text1(0).Text, 6) & Year(DateValue(Now))
        Text1(1).Text = Left(Text1(1).Text, 6) & Year(DateValue(Now))
        
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        
        umtmefuellen
        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat1", "aWKL25aa"
        End If
        rs.Close: Set rs = Nothing
    ElseIf optqp.Value = True Then 'Monatsansicht
    
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        Text1(0).Text = Left(Text1(0).Text, 6) & Year(DateValue(Now))
        Text1(1).Text = Left(Text1(1).Text, 6) & Year(DateValue(Now))
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        umtmefuellen
        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat2", "aWKL25ab"
        End If
        rs.Close: Set rs = Nothing
    ElseIf optWeek.Value = True Then 'Wochenansicht
    
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        Text1(0).Text = Left(Text1(0).Text, 6) & Year(DateValue(Now))
        Text1(1).Text = Left(Text1(1).Text, 6) & Year(DateValue(Now))
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        umtmefuellen
        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat2", "aWKL25ae"

        End If
        rs.Close: Set rs = Nothing
    ElseIf optD.Value = True Then
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        Text1(0).Text = Left(Text1(0).Text, 6) & Year(DateValue(Now))
        Text1(1).Text = Left(Text1(1).Text, 6) & Year(DateValue(Now))
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        umtmefuellen
        
        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat3", "aWKL25ac"
        End If
        rs.Close: Set rs = Nothing

        
    ElseIf optAkt.Value = True Then 'aktueller Monat tagesgenau
        
        Text1(0).Text = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = DateValue(Now)
        
        umtmefuellen
        
        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat6", "aWKL25af"
        End If
        rs.Close: Set rs = Nothing
    ElseIf optZR.Value = True Then

        umtmefuellen

        Set rs = gdBase.OpenRecordset("umtme")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            If Modul6.FindFile(gcDBPfad, "aWKL25s.rpt") Then
                reportbildschirm "spez5", "aWKL25s"
            Else
                reportbildschirm "Umstat4", "aWKL25ad"
            End If
        End If
        rs.Close: Set rs = Nothing
        
    ElseIf optZRVJ.Value = True Then

        umtZRVJfuellen

        Set rs = gdBase.OpenRecordset("umZRVJ")
        If rs.RecordCount = 0 Then
            lblanzeige.ForeColor = vbRed
            lblanzeige.Caption = "Es wurde keine Umsatzdaten im angegebenen Zeitraum ermittelt."
            lblanzeige.Refresh
        Else
            reportbildschirm "Umstat4", "aWKL25ag"
        End If
        rs.Close: Set rs = Nothing
    ElseIf Option1.Value = True Then 'eingegebner Zeitraum mit Monatsansicht
    
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        Text1(0).Text = Left(Text1(0).Text, 6) & Year(DateValue(Now))
        Text1(1).Text = Left(Text1(1).Text, 6) & Year(DateValue(Now))
        
        Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YY")
        Text1(1).Text = Format$(Text1(1).Text, "DD.MM.YY")
        
        
        umvergfuellenGemischt_ZR
        
        anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
        reportbildschirm "", "aWKL25ai"
        
        

    End If
    
    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdQuick_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub umvergfuellen(sMWS As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    Dim cVJVon      As String
    Dim cVJBis      As String
    Dim lVJVon      As Long
    Dim lVJBis      As Long
    Dim cJahr       As String
    Dim i           As Integer
    
    cJahr = Mid(Text1(1).Text, 7, 2)
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "Umverg", gdBase
    
    sSQL = "Create Table Umverg "
    sSQL = sSQL & "( Monat Text(2)"
    sSQL = sSQL & ", Jahr Integer"
    sSQL = sSQL & ", UMSGAJ Double"
    sSQL = sSQL & ", UMSGVJ Double"
    sSQL = sSQL & ", KZAJ Double"
    sSQL = sSQL & ", KZVJ Double"
    sSQL = sSQL & ", UMSVAJ Double"
    sSQL = sSQL & ", UMSEAJ Double"
    sSQL = sSQL & ", UMSOAJ Double"
    sSQL = sSQL & ", UMSVVJ Double"
    sSQL = sSQL & ", UMSEVJ Double"
    sSQL = sSQL & ", UMSOVJ Double"
    sSQL = sSQL & ", EKPAJ Double"
    sSQL = sSQL & ", EKPVJ Double"
    sSQL = sSQL & " )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    If cJahr < 50 Then
        cJahr = "20" & cJahr
    Else
        cJahr = "19" & cJahr
    End If
    For i = 1 To 12
        sSQL = "Insert into Umverg (Monat,Jahr) Values (" & i & ", " & CInt(cJahr) & " ) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    t
    For i = 1 To 12 'Bruttoumsatz aktuelles Jahr
        If sMWS = "BRUTTO" Then
            sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        Else
            sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / (100 + " & gdMWStV & ")) + ((sum(UMSe1)*100) / (100 + " & gdMWStE & ")) + sum(umso1))as t,  " & i
        End If
'        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.umsgaj = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
    t
    For i = 1 To 12 'Bruttoumsatz Vorjahr
        If sMWS = "BRUTTO" Then
            sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        Else
            Dim dMWStV As Double
            Dim dMWStE As Double
            
            If cJahr - 1 < 2007 Then
                dMWStV = 116
                dMWStE = 107
            Else
            
                dMWStV = 119
                dMWStE = 107
            End If
    
            sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / " & dMWStV & ") + ((sum(UMSe1)*100) / " & dMWStE & ") + sum(umso1)) as t,  " & i
        
        End If
'        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.umsgvj = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
    t
    For i = 1 To 12 'Kundenzahl aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(KUNZ1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.KZAJ = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
    t
    For i = 1 To 12 'Kundenzahl Vorjahr
        sSQL = "Insert into Umtemp Select sum(KUNZ1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.KZVJ = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    t
    For i = 1 To 12 'EK aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(EKPR1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.EKPAJ = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    t
    For i = 1 To 12 'EK vorjahr
        sSQL = "Insert into Umtemp Select sum(EKPR1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from UMSATZ where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        sSQL = sSQL & " group by month(DATUM)"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Next i
    sSQL = "update Umverg INNER JOIN Umtemp ON "
    sSQL = sSQL & " (umverg.Monat = umtemp.monat)"
    sSQL = sSQL & " set umverg.EKPVJ = umtemp.t "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "umvergfuellen"
        Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub t()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    loeschNEW "UMTEMP", gdBase
    
    sSQL = "Create Table Umtemp "
    sSQL = sSQL & "( Monat Text(2)"
    sSQL = sSQL & ", T Double"
    sSQL = sSQL & " )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "t"
        Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
    
End Sub
Private Sub umtmefuellen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    Dim cVJVon      As String
    Dim cVJBis      As String
    Dim lVJVon      As Long
    Dim lVJBis      As Long
    Dim cJahr       As String
    Dim i           As Integer
    
    cJahr = Year(Now)
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "UMTME", gdBase
    
    sSQL = "Create Table Umtme"
    sSQL = sSQL & "( DATUM DateTime"
    sSQL = sSQL & ", UMSG1 Double"
    sSQL = sSQL & ", UMSV1 Double"
    sSQL = sSQL & ", UMSE1 Double"
    sSQL = sSQL & ", UMSO1 Double"
    sSQL = sSQL & ", KUNZ1 Double"
    sSQL = sSQL & ", EKPR1 Double"
    sSQL = sSQL & ", KRED1 Double"
    sSQL = sSQL & ", von Text(5)"
    sSQL = sSQL & ", bis Text(5)"
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Umtme Select DATUM, UMSG1, UMSV1, UMSE1, KRED1,"
    sSQL = sSQL & " KUNZ1, EKPR1,UMSO1 "
    sSQL = sSQL & " from UMSATZ where "
    sSQL = sSQL & " DATUM >= " & cVon & " and DATUM <= " & cBis & " "
    sSQL = sSQL & " order by DATUM"
    gdBase.Execute sSQL, dbFailOnError
    
    If optAkt.Value = True Or optZR.Value = True Then
    
    Else
        For i = 1 To 7 '4
        
            If IsDate(Left$(Text1(0).Text, 6) & cJahr - 1) Then
                cVJVon = Left$(Text1(0).Text, 6) & cJahr - i
            Else
                If Left$(Text1(0).Text, 5) = "29.02" Then
                    cVJVon = "28.02." & cJahr - i
                End If
            End If
            If IsDate(Left$(Text1(1).Text, 6) & cJahr - i) Then
                cVJBis = Left$(Text1(1).Text, 6) & cJahr - i
            Else
                If Left$(Text1(1).Text, 5) = "29.02" Then
                    cVJBis = "28.02." & cJahr - i
                End If
            End If
        
            lVJVon = DateValue(cVJVon)
            lVJBis = DateValue(cVJBis)
            cVJVon = Trim$(Str$(lVJVon))
            cVJBis = Trim$(Str$(lVJBis))
            
            sSQL = "Insert into Umtme Select DATUM, UMSG1, UMSV1, UMSE1, KRED1,"
            sSQL = sSQL & " KUNZ1, EKPR1,UMSO1 "
            sSQL = sSQL & " from UMSATZ where "
            sSQL = sSQL & " DATUM >= " & cVJVon & " and DATUM <= " & cVJBis & " "
            sSQL = sSQL & " order by DATUM"
            gdBase.Execute sSQL, dbFailOnError
        
        Next i
    End If
    
    sSQL = "Delete from umtme where umsg1 = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from umtme where umsg1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update umtme set von = '" & Left(Text1(0).Text, 5) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update umtme set bis = '" & Left(Text1(1).Text, 5) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "temp", gdBase
    
    sSQL = "select * into temp from umtme"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "umtme", gdBase
    
    sSQL = "select * into umtme from temp"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "temp", gdBase
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "umtmefuellen"
        Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub umtZRVJfuellen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    Dim cVJVon      As String
    Dim cVJBis      As String
    Dim lVJVon      As Long
    Dim lVJBis      As Long
    Dim cJahr       As String
    Dim i           As Integer
    Dim lDat        As Long
    Dim rsUms As Recordset
    Dim rsUmsVJ As Recordset
    Dim sDatum As String
    
    Dim iJahr       As Integer
    Dim lHeuteVJ    As Long
    Dim lHeuteVJPlus    As Long
    Dim cheuteVJ    As String
    Dim siUGVJ     As Single
    Dim siUVVJ     As Single
    Dim siUEVJ     As Single
    Dim siUOVJ     As Single
    Dim siKRVJ     As Single
    Dim siEKVJ     As Single
    Dim siKUVJ     As Single
    
    
    
    
    cJahr = Year(Now)
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "umZRVJ", gdBase
    
    sSQL = "Create Table umZRVJ"
    sSQL = sSQL & "( DATUM DateTime"
    sSQL = sSQL & ", UMSG1 Double"
    sSQL = sSQL & ", UMSV1 Double"
    sSQL = sSQL & ", UMSE1 Double"
    sSQL = sSQL & ", UMSO1 Double"
    sSQL = sSQL & ", KUNZ1 Double"
    sSQL = sSQL & ", EKPR1 Double"
    sSQL = sSQL & ", KRED1 Double"
    
    sSQL = sSQL & ", UMSGVJ Double"
    sSQL = sSQL & ", UMSVVJ Double"
    sSQL = sSQL & ", UMSEVJ Double"
    sSQL = sSQL & ", UMSOVJ Double"
    sSQL = sSQL & ", KUNZVJ Double"
    sSQL = sSQL & ", EKPRVJ Double"
    sSQL = sSQL & ", KREDVJ Double"
    
    sSQL = sSQL & ", von Text(5)"
    sSQL = sSQL & ", bis Text(5)"
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    For lDat = lVon To lBis
        sSQL = "Insert into umZRVJ (datum) values (" & Trim$(Str$(lDat)) & ")"
        gdBase.Execute sSQL, dbFailOnError
    Next lDat
    
    sSQL = "Update umZRVJ inner Join umsatz on"
    sSQL = sSQL & " umZRVJ.datum = umsatz.datum "
    sSQL = sSQL & " set umZRVJ.UMSG1 = umsatz.umsg1"
    sSQL = sSQL & " , umZRVJ.UMSV1 = umsatz.UMSV1"
    sSQL = sSQL & " , umZRVJ.UMSE1 = umsatz.UMSE1"
    sSQL = sSQL & " , umZRVJ.UMSO1 = umsatz.UMSO1"
    sSQL = sSQL & " , umZRVJ.KUNZ1 = umsatz.KUNZ1"
    sSQL = sSQL & " , umZRVJ.EKPR1 = umsatz.EKPR1"
    sSQL = sSQL & " , umZRVJ.KRED1 = umsatz.KRED1"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsUms = gdBase.OpenRecordset("umZRVJ", dbOpenTable)
    If Not rsUms.RecordCount = 0 Then
        rsUms.MoveFirst
        Do While Not rsUms.EOF
        
            If Not IsNull(rsUms!Datum) Then
                sDatum = Format$(CLng(rsUms!Datum), "DD.MM.YYYY")
            Else
                sDatum = ""
            End If
            
            If sDatum = "29.02.2012" Then
            
            ElseIf sDatum = "29.02.2016" Then
            
            ElseIf sDatum = "29.02.2020" Then
            
            ElseIf sDatum = "29.02.2024" Then
   
            Else
                iJahr = CInt(Right(sDatum, 4))
                iJahr = iJahr - 1
    
                cheuteVJ = Left(sDatum, 6) & CStr(iJahr)
                lHeuteVJ = datumwandlung(cheuteVJ)
                
                If sDatum = "28.02.2013" Then
                    lHeuteVJPlus = lHeuteVJ + 1
                Else
                    lHeuteVJPlus = lHeuteVJ
                End If
                
                sSQL = "Select sum(umsg1) as umsg,sum(umsv1) as umsv,sum(umse1) as umse,sum(umso1) as umso,sum(kunz1) as kunz,sum(ekpr1) as ekpr,sum(kred1) as kred from umsatz"
                sSQL = sSQL & " where datum between " & lHeuteVJ & " and " & lHeuteVJPlus & " "
                
'                sSQL = "Select * from umsatz"
'                sSQL = sSQL & " where datum = " & lHeuteVJ
                Set rsUmsVJ = gdBase.OpenRecordset(sSQL)
                If Not rsUmsVJ.RecordCount = 0 Then
                    rsUmsVJ.MoveFirst
                    
                    If Not rsUmsVJ.EOF Then
                        If Not IsNull(rsUmsVJ!UMSG) Then
                            siUGVJ = rsUmsVJ!UMSG
                        Else
                            siUGVJ = 0
                        End If
                    
                        If Not IsNull(rsUmsVJ!UMSV) Then
                            siUVVJ = rsUmsVJ!UMSV
                        Else
                            siUVVJ = 0
                        End If
                   
                        If Not IsNull(rsUmsVJ!UMSe) Then
                            siUEVJ = rsUmsVJ!UMSe
                        Else
                            siUEVJ = 0
                        End If
        
                        If Not IsNull(rsUmsVJ!UMSo) Then
                            siUOVJ = rsUmsVJ!UMSo
                        Else
                            siUOVJ = 0
                        End If
                    
                        If Not IsNull(rsUmsVJ!KUNZ) Then
                            siKUVJ = rsUmsVJ!KUNZ
                        Else
                            siKUVJ = 0
                        End If
                   
                        If Not IsNull(rsUmsVJ!ekpr) Then
                            siEKVJ = rsUmsVJ!ekpr
                        Else
                            siEKVJ = 0
                        End If
                    
                        If Not IsNull(rsUmsVJ!KRED) Then
                            siKRVJ = rsUmsVJ!KRED
                        Else
                            siKRVJ = 0
                        End If
                        
                    End If
                    
                Else
                    siUGVJ = 0
                    siUVVJ = 0
                    siUEVJ = 0
                    siUOVJ = 0
                    siKRVJ = 0
                    siEKVJ = 0
                    siKUVJ = 0
                End If
                rsUmsVJ.Close
                
                rsUms.Edit
                rsUms!UMSGVJ = siUGVJ
                rsUms!UMSVVJ = siUVVJ
                rsUms!UMSEVJ = siUEVJ
                rsUms!UMSOVJ = siUOVJ
                rsUms!KREDVJ = siKRVJ
                rsUms!EKPRVJ = siEKVJ
                rsUms!KUNZVJ = siKUVJ
                
                rsUms.Update
            End If
            rsUms.MoveNext
        Loop
    End If
    rsUms.Close

    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "umtZRVJfuellen"
        Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        
        
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(1).SetFocus
            
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            
        End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
  
    Unload frmWK25a

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR
    
    speicher_Umsatzinfo Text2.Text, Right(Label1(1).Caption, 8)
    
    Frame1.Visible = False
    
    cmdPlus_Click
    cmdMinus_Click
    
Exit Sub
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher_Umsatzinfo(sInfo As String, sDate As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lDate As Long
    lDate = DateValue(sDate)
    
    sSQL = "Delete from Umsatzinfo where datum = " & lDate & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If sInfo <> "" Then
        sSQL = "Insert into Umsatzinfo (Datum,INFO) values (" & lDate & ", '" & sInfo & "')"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_Umsatzinfo"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function zeige_Umsatzinfo(sDate As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim lDate As Long
    
    If sDate = "Label6" Then
        Exit Function
    End If
    
    lDate = DateValue(sDate)
    
    zeige_Umsatzinfo = ""
    
    sSQL = "Select * from Umsatzinfo where datum = " & lDate & ""
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Info) Then
            zeige_Umsatzinfo = rsrs!Info
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Umsatzinfo"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lHeute      As Long
    lHeute = Fix(Now)
   
    Screen.MousePointer = 11
    PositionierenWK25a
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    gitop = Shape1(0).Top
    ZeigeUmsatzgrafik lHeute
    
    optZR.Value = True
    
    Text1(0).Text = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
    Text1(0).Text = Format$(Text1(0).Text, "DD.MM.YYYY")
    Text1(1).Text = DateValue(Now)
    
    Command2(0).Caption = Right(Year(Now) - 2, 2) & " / " & Right(Year(Now) - 1, 2)
    Command2(1).Caption = Right(Year(Now) - 1, 2) & " / " & Right(Year(Now), 2)
    
    Dim sSQL As String
    
    sSQL = "Delete * from  Umsatz where datum is null "
    gdBase.Execute sSQL, dbFailOnError
        
    Screen.MousePointer = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
Exit Sub
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ZeigeUmsatzgrafik(lDat As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsUms       As Recordset
    Dim rsUmsVJ     As Recordset
    Dim siUms(6)    As Single
    Dim sdat(6)     As String
    Dim siUmsVJ(6)  As Single
    Dim sDatVJ(6)   As String
    Dim i           As Integer
    Dim dbuffer     As Double
    Dim dMax        As Double
    Dim sDatum      As String
    Dim iJahr       As Integer
    Dim lHeuteVJ    As Long
    Dim cheuteVJ    As String
    Dim siWert      As Single
    Dim Wert1       As Integer
    Dim sTabname    As String
    
    Randomize
        
    Wert1 = Int((99 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99
    sTabname = "USS" & CStr(Wert1)
    
    For i = 0 To 6
        Shape1(i).Visible = False
        Shape2(i).Visible = False
    Next i
    
    loeschNEW sTabname, gdBase
    
    sSQL = "Create table " & sTabname
    sSQL = sSQL & " ( datum  Datetime "
    sSQL = sSQL & " , umsakt single "
    sSQL = sSQL & " , datumvj  Datetime "
    sSQL = sSQL & " , umsVj single "
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    For i = 0 To 10
        lDat = lDat - 1
        sSQL = "Insert into " & sTabname & " (datum) values (" & Trim$(Str$(lDat)) & ")"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    sSQL = "Update " & sTabname & " inner Join umsatz on "
    sSQL = sSQL & " " & sTabname & ".datum = umsatz.datum "
    sSQL = sSQL & " set " & sTabname & ".umsakt = umsatz.umsg1"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    Set rsUms = gdBase.OpenRecordset(sTabname, dbOpenTable)
    If Not rsUms.RecordCount = 0 Then
        rsUms.MoveFirst
        Do While Not rsUms.EOF
        
            If Not IsNull(rsUms!Datum) Then
                sDatum = Format$(CLng(rsUms!Datum), "DD.MM.YYYY")
            Else
                sDatum = ""
            End If
            
            If sDatum = "29.02.2012" Then
            
            ElseIf sDatum = "29.02.2016" Then
            
            ElseIf sDatum = "29.02.2020" Then
            
            ElseIf sDatum = "29.02.2024" Then
   
            Else
   
                iJahr = CInt(Right(sDatum, 4))
                iJahr = iJahr - 1
    
                cheuteVJ = Left(sDatum, 6) & CStr(iJahr)
    
                lHeuteVJ = DateValue(cheuteVJ)
    
                
                sSQL = "Select umsg1 from umsatz"
                sSQL = sSQL & " where datum = " & lHeuteVJ
                Set rsUmsVJ = gdBase.OpenRecordset(sSQL)
                If Not rsUmsVJ.RecordCount = 0 Then
                    rsUmsVJ.MoveFirst
                    If Not rsUmsVJ.EOF Then
                        If Not IsNull(rsUmsVJ!UMSG1) Then
                            siWert = rsUmsVJ!UMSG1
                        Else
                            siWert = 0
                        End If
                    End If
                Else
                    siWert = 0
                End If
                rsUmsVJ.Close
                
                rsUms.Edit
                rsUms!UMSVJ = siWert
                rsUms!datumVJ = lHeuteVJ
                rsUms.Update
            End If
            rsUms.MoveNext
        Loop
    End If
    rsUms.Close
                           
    Set rsUms = gdBase.OpenRecordset(sTabname, dbOpenTable)
    If Not rsUms.RecordCount = 0 Then
    i = 6
    rsUms.MoveFirst
    
        Do While Not rsUms.EOF
        If i < 0 Then Exit Do
        
            If Not IsNull(rsUms!UMSakt) Then
                siUms(i) = rsUms!UMSakt
            Else
                siUms(i) = 0
            End If
            
            If Not IsNull(rsUms!Datum) Then
                sdat(i) = Format$(CLng(rsUms!Datum), "DD.MM.YY")
            Else
                sdat(i) = "0"
            End If
            
            If Not IsNull(rsUms!UMSVJ) Then
                siUmsVJ(i) = rsUms!UMSVJ
            Else
                siUmsVJ(i) = 0
            End If
            
            If Not IsNull(rsUms!datumVJ) Then
                sDatVJ(i) = Format$(CLng(rsUms!datumVJ), "DD.MM.YY")
            Else
                sDatVJ(i) = ""
            End If
            
            
            i = i - 1
        rsUms.MoveNext
        Loop
    End If
    rsUms.Close
    
    For i = 0 To 6
        Label4(i).Caption = WeekdayName(Weekday(DateValue(sdat(i)) - 1)) & vbCrLf & Format$(sdat(i), "DD.MM.YY")
        Label4(i).Refresh

        Label6(i).Caption = WeekdayName(Weekday(DateValue(sDatVJ(i)) - 1)) & vbCrLf & Format$(sDatVJ(i), "DD.MM.YY")
        Label6(i).Refresh
    
        Label2(i).Caption = Format$(sdat(i), "DD.MM")
        Label2(i).Refresh
    Next i
    
    dbuffer = 0
    dMax = 0
    
    For i = 0 To 6
        dbuffer = siUms(i)
        If dbuffer > dMax Then
            dMax = dbuffer
        End If
    Next i
    
    For i = 0 To 6
        dbuffer = siUmsVJ(i)
        If dbuffer > dMax Then
            dMax = dbuffer
        End If
    Next i
    dMax = IIf(dMax = 0, 1, dMax)
    
    For i = 0 To 6
        If siUms(i) > 0 Then
            Shape1(i).Height = (4000 / dMax) * IIf(siUms(i) < 0, 0, siUms(i))
            Shape1(i).Top = gitop - ((4000 / dMax) * siUms(i))
            Label3(i).Top = Shape1(i).Top - 300
            Label3(i).Caption = Format$(siUms(i), "###,##0")
            Label3(i).Refresh
        Else
            Shape1(i).Height = 15
            Shape1(i).Top = gitop
            Label3(i).Top = gitop - 300
            Label3(i).Caption = siUms(i)
            Label3(i).Refresh
        End If
        
        If zeige_Umsatzinfo(Right(Label4(i).Caption, 8)) <> "" Then
            Label3(i).ForeColor = vbRed
        Else
            Label3(i).ForeColor = glS1
        End If
    Next i
    
    For i = 0 To 6
        If siUmsVJ(i) > 0 Then
            Shape2(i).Height = (4000 / dMax) * IIf(siUmsVJ(i) < 0, 0, siUmsVJ(i))
            Shape2(i).Top = gitop - ((4000 / dMax) * siUmsVJ(i))
            Label5(i).Top = Shape2(i).Top - 300
            Label5(i).Caption = Format$(siUmsVJ(i), "###,##0")
            Label5(i).Refresh
        Else
            Shape2(i).Height = 15
            Shape2(i).Top = gitop
            Label5(i).Top = gitop - 300
            Label5(i).Caption = siUmsVJ(i)
            Label5(i).Refresh
        End If
        
        If zeige_Umsatzinfo(Right(Label6(i).Caption, 8)) <> "" Then
            Label5(i).ForeColor = vbRed
        Else
            Label5(i).ForeColor = glS1
        End If
    Next i

    For i = 0 To 6
        Shape1(i).Visible = True
        Shape2(i).Visible = True
    Next i

    Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZeigeUmsatzgrafik"
        Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Private Sub PositionierenWK25a()
    On Error GoTo LOKAL_ERROR

    Frame1.Top = 7200
    Frame1.Height = 1335
    Frame1.Width = 9255
    Frame1.Left = 120
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK25a"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i       As Integer
    
    For i = 0 To 99
        loeschNEW "USS" & i, gdBase
    Next i
    
    LogtoEnd Me
    
    loeschNEW "UMVERG", gdBase
    loeschNEW "UMTEMP", gdBase
    loeschNEW "UMTME", gdBase

    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt, bitte warten...", lblanzeige
    Select Case Index
    
        Case 0
            
            Text1(1).Text = "31.12." & Year(Now) - 1
            Text1(0).Text = "01.01." & Year(Now) - 2
            umvergfuellenGemischt False
            
            
        Case 1
        
            Text1(1).Text = "31.12." & Year(Now)
            Text1(0).Text = "01.01." & Year(Now) - 1
            
            If Check1.Value = vbChecked Then
                
                umvergfuellenGemischt True
            Else
                umvergfuellenGemischt False
            End If
            
            
    End Select
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
    reportbildschirm "", "aWKL25ai"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub umvergfuellenGemischt(bBisHeute As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim dbrec       As Recordset
    Dim dbrec1      As Recordset
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    Dim cVJVon      As String
    Dim cVJBis      As String
    Dim lVJVon      As Long
    Dim lVJBis      As Long
    Dim cJahr       As String
    Dim i           As Integer
    Dim j           As Integer
    
    cJahr = Mid$(Text1(1).Text, 7, 4)
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "UMverg", gdBase

    sSQL = "Create Table Umverg "
    sSQL = sSQL & "( Monat Text(2)"
    sSQL = sSQL & ", Jahr Integer"
    sSQL = sSQL & ", UMSGAJ Double"
    sSQL = sSQL & ", UMSGVJ Double"
    sSQL = sSQL & ", KZAJ Double"
    sSQL = sSQL & ", KZVJ Double"
    sSQL = sSQL & ", UMSVAJ Double"
    sSQL = sSQL & ", UMSEAJ Double"
    sSQL = sSQL & ", UMSOAJ Double"
    sSQL = sSQL & ", UMSVVJ Double"
    sSQL = sSQL & ", UMSEVJ Double"
    sSQL = sSQL & ", UMSOVJ Double"
    sSQL = sSQL & ", EKPAJ Double"
    sSQL = sSQL & ", EKPVJ Double"
    
    sSQL = sSQL & ", UMSGAJb Double"
    sSQL = sSQL & ", UMSGVJb Double"
    sSQL = sSQL & ", UMSVAJb Double"
    sSQL = sSQL & ", UMSEAJb Double"
    sSQL = sSQL & ", UMSOAJb Double"
    sSQL = sSQL & ", UMSVVJb Double"
    sSQL = sSQL & ", UMSEVJb Double"
    sSQL = sSQL & ", UMSOVJb Double"
    
    sSQL = sSQL & ", Filiale Text (50)"
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    If CInt(cJahr) < 50 Then
        cJahr = "20" & cJahr
    ElseIf Len(cJahr) < 4 Then
        cJahr = "19" & cJahr
    End If
    For i = 1 To 12
        sSQL = "Insert into Umverg (Monat,Jahr) Values (" & i & ", " & CInt(cJahr) & " ) "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    t
    For i = 1 To 12 'Bruttoumsatz aktuelles Jahr
        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgajb from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    
    t
    For i = 1 To 12 'Nettoumsatz aktuelles Jahr
    
        sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / (100 + " & gdMWStV & ")) + ((sum(UMSe1)*100) / (100 + " & gdMWStE & ")) + sum(umso1))as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgaj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i

    Dim cDatum As String
    cDatum = Format(DateValue(Now), "DD.MM.") & cJahr - 1

    t
    For i = 1 To 12 'Bruttoumsatz Vorjahr
        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
        If bBisHeute Then
            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        
        gdBase.Execute sSQL, dbFailOnError
    Next i

    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgvjb from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    t
    For i = 1 To 12 'Nettoumsatz Vorjahr
    
        Dim dMWStV As Double
        Dim dMWStE As Double
        
        If cJahr - 1 < 2007 Then
            dMWStV = 116
            dMWStE = 107
        Else
        
            dMWStV = 119
            dMWStE = 107
        End If

        sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / " & dMWStV & ") + ((sum(UMSe1)*100) / " & dMWStE & ") + sum(umso1)) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
        If bBisHeute Then
            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i

    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    sSQL = "Update umverg Set umsgvj = null where umsgvj = 0"
    gdBase.Execute sSQL, dbFailOnError
  
    t
    For i = 1 To 12 'Kundenzahl aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(kunz1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select KZAJ from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i

    t
    For i = 1 To 12 'Kundenzahl Vorjahr
        sSQL = "Insert into Umtemp Select sum(kunz1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
        If bBisHeute Then
            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
        End If
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select KZvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    t
    For i = 1 To 12 'EK aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(ekpr1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select ekpaj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i

    t
    For i = 1 To 12 'EK vorjahr
        sSQL = "Insert into Umtemp Select sum(ekpr1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
        If bBisHeute Then
            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = 1 To 12
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select ekpvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i
   
    sSQL = "update Umverg set Filiale = '1 Geschäft'"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "umvergfuellenGemischt"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub umvergfuellenGemischt_ZR()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim dbrec       As Recordset
    Dim dbrec1      As Recordset
    
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cVon        As String
    Dim cBis        As String
    Dim cVJVon      As String
    Dim cVJBis      As String
    Dim lVJVon      As Long
    Dim lVJBis      As Long
    Dim cJahr       As String
    Dim i           As Integer
    Dim j           As Integer
    
    Dim iVonMonat   As Integer
    Dim iBisMonat   As Integer
    
    iVonMonat = 2
    iBisMonat = 4
    
    cJahr = Mid$(Text1(1).Text, 7, 4)
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    iVonMonat = Month(lVon)
    iBisMonat = Month(lBis)
    
    If CInt(cJahr) < 50 Then
        cJahr = "20" & cJahr
    ElseIf Len(cJahr) < 4 Then
        cJahr = "19" & cJahr
    End If
    
    loeschNEW "UMverg", gdBase

    sSQL = "Create Table Umverg "
    sSQL = sSQL & "( Monat Text(2)"
    sSQL = sSQL & ", Jahr Integer"
    sSQL = sSQL & ", UMSGAJ Double"
    sSQL = sSQL & ", UMSGVJ Double"
    sSQL = sSQL & ", KZAJ Double"
    sSQL = sSQL & ", KZVJ Double"
    sSQL = sSQL & ", UMSVAJ Double"
    sSQL = sSQL & ", UMSEAJ Double"
    sSQL = sSQL & ", UMSOAJ Double"
    sSQL = sSQL & ", UMSVVJ Double"
    sSQL = sSQL & ", UMSEVJ Double"
    sSQL = sSQL & ", UMSOVJ Double"
    sSQL = sSQL & ", EKPAJ Double"
    sSQL = sSQL & ", EKPVJ Double"
    
    sSQL = sSQL & ", UMSGAJb Double"
    sSQL = sSQL & ", UMSGVJb Double"
    sSQL = sSQL & ", UMSVAJb Double"
    sSQL = sSQL & ", UMSEAJb Double"
    sSQL = sSQL & ", UMSOAJb Double"
    sSQL = sSQL & ", UMSVVJb Double"
    sSQL = sSQL & ", UMSEVJb Double"
    sSQL = sSQL & ", UMSOVJb Double"
    
    sSQL = sSQL & ", Filiale Text (50)"
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    For i = iVonMonat To iBisMonat
        sSQL = "Insert into Umverg (Monat,Jahr) Values (" & i & ", " & CInt(cJahr) & " ) "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    t
    For i = iVonMonat To iBisMonat 'Bruttoumsatz aktuelles Jahr
        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgajb from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    
    t
    For i = iVonMonat To iBisMonat 'Nettoumsatz aktuelles Jahr
    
        sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / (100 + " & gdMWStV & ")) + ((sum(UMSe1)*100) / (100 + " & gdMWStE & ")) + sum(umso1))as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgaj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i

'    Dim cDatum As String
'    cDatum = Format(DateValue(Now), "DD.MM.") & cJahr - 1

    t
    For i = iVonMonat To iBisMonat 'Bruttoumsatz Vorjahr
        sSQL = "Insert into Umtemp Select sum(UMSG1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
'        If bBisHeute Then
'            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
'        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        
        gdBase.Execute sSQL, dbFailOnError
    Next i

    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgvjb from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    t
    For i = iVonMonat To iBisMonat 'Nettoumsatz Vorjahr
    
        Dim dMWStV As Double
        Dim dMWStE As Double
        
        If cJahr - 1 < 2007 Then
            dMWStV = 116
            dMWStE = 107
        Else
        
            dMWStV = 119
            dMWStE = 107
        End If

        sSQL = "Insert into Umtemp Select (((sum(UMSV1)*100) / " & dMWStV & ") + ((sum(UMSe1)*100) / " & dMWStE & ") + sum(umso1)) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
'        If bBisHeute Then
'            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
'        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i

    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select umsgvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    sSQL = "Update umverg Set umsgvj = null where umsgvj = 0"
    gdBase.Execute sSQL, dbFailOnError
  
    t
    For i = iVonMonat To iBisMonat 'Kundenzahl aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(kunz1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select KZAJ from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i

    t
    For i = iVonMonat To iBisMonat 'Kundenzahl Vorjahr
        sSQL = "Insert into Umtemp Select sum(kunz1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
'        If bBisHeute Then
'            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
'        End If
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select KZvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
    Next i
    
    t
    For i = iVonMonat To iBisMonat 'EK aktuelles jahr
        sSQL = "Insert into Umtemp Select sum(ekpr1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select ekpaj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i

    t
    For i = iVonMonat To iBisMonat 'EK vorjahr
        sSQL = "Insert into Umtemp Select sum(ekpr1) as t,  " & i
        sSQL = sSQL & " as Monat"
        sSQL = sSQL & " from Umsatz where "
        sSQL = sSQL & " month(Datum) = " & i
        sSQL = sSQL & " and year(Datum) = " & cJahr - 1
        
'        If bBisHeute Then
'            sSQL = sSQL & " and Datum <= " & CLng(DateValue(cDatum))
'        End If
        
        sSQL = sSQL & " group by month(DATUM)"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    For i = iVonMonat To iBisMonat
        Set dbrec = gdBase.OpenRecordset("select sum(t) from umtemp where monat = '" & CStr(i) & "'")
        If Not dbrec.EOF Then
            Set dbrec1 = gdBase.OpenRecordset("select ekpvj from umverg where monat = '" & CStr(i) & "'")
            dbrec1.Edit
            dbrec1.Fields(0) = dbrec.Fields(0)
            dbrec1.Update
            dbrec1.Close
        End If
        dbrec.Close
        
    Next i
   
    sSQL = "update Umverg set Filiale = '1 Geschäft'"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "umvergfuellenGemischt_ZR"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."

    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Label3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Label1(1).Caption = SwapStr(Label4(Index).Caption, vbCrLf, " ")
    Text2.Text = zeige_Umsatzinfo(Right(Label4(Index).Caption, 8))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label3(Index).ToolTipText = zeige_Umsatzinfo(Right(Label4(Index).Caption, 8))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label5(Index).ToolTipText = zeige_Umsatzinfo(Right(Label6(Index).Caption, 8))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Label1(1).Caption = SwapStr(Label6(Index).Caption, vbCrLf, " ")
    Text2.Text = zeige_Umsatzinfo(Right(Label6(Index).Caption, 8))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label5_Click"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR

    Text2.BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus()
    On Error GoTo LOKAL_ERROR

    Text2.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    lDatum = DateValue(Label2(6).Caption)
    ZeigeUmsatzgrafik lDatum
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer2_Timer()
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    lDatum = DateValue(Label2(6).Caption)
    lDatum = lDatum + 2
    ZeigeUmsatzgrafik lDatum
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Umsatzstatistik ist ein Fehler aufgetreten. " & lDatum
    
    Fehlermeldung1
End Sub
