VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL23 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Wareneingang aus Umverteilung (Filiale)"
   ClientHeight    =   8625
   ClientLeft      =   2145
   ClientTop       =   2655
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
   Icon            =   "frmWKL23.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame9 
      BackColor       =   &H008080FF&
      Caption         =   "Frame9"
      Height          =   375
      Left            =   8520
      TabIndex        =   130
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      Begin VB.FileListBox File2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   9120
         Pattern         =   "WV*.dbf"
         TabIndex        =   138
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox List13 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   120
         TabIndex        =   137
         Top             =   1920
         Width           =   6495
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   20
         Left            =   9360
         TabIndex        =   136
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   19
         Left            =   9360
         TabIndex        =   135
         Top             =   5640
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   18
         Left            =   9360
         TabIndex        =   134
         Top             =   5040
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   17
         Left            =   9720
         TabIndex        =   133
         Top             =   6960
         Width           =   1935
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
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   16
         Left            =   9360
         TabIndex        =   132
         Top             =   4440
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
         Caption         =   "Holen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   15
         Left            =   9360
         TabIndex        =   131
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
         Caption         =   "Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List15 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   148
         Top             =   1560
         Width           =   6495
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   22
         Left            =   9360
         TabIndex        =   149
         Top             =   3840
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
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Expressdateien annehmen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   139
         Top             =   240
         Width           =   6135
      End
   End
   Begin sevCommand3.Command Command11 
      Height          =   525
      Index           =   2
      Left            =   7200
      TabIndex        =   128
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
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
      Caption         =   "Kundenbestellungen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808000&
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   107
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   6000
         TabIndex        =   110
         Top             =   1080
         Width           =   5535
      End
      Begin VB.ListBox List11 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   112
         Top             =   1080
         Width           =   5535
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00C0C000&
         Caption         =   "mit Etikettenerstellung"
         Height          =   210
         Left            =   9360
         TabIndex        =   116
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   12
         Left            =   9360
         TabIndex        =   115
         Top             =   5520
         Visible         =   0   'False
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   11
         Left            =   9360
         TabIndex        =   114
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   113
         Top             =   840
         Width           =   5535
      End
      Begin VB.ListBox List4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   111
         Top             =   840
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   8
         Left            =   10440
         TabIndex        =   109
         Top             =   4560
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   108
         Top             =   6360
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel aus dem MDE - Gerät"
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
         Index           =   7
         Left            =   120
         TabIndex        =   121
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel, die nicht zugeordnet werden können"
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
         Index           =   6
         Left            =   6000
         TabIndex        =   120
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   119
         Top             =   4560
         Width           =   5535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   4
         Left            =   6000
         TabIndex        =   118
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Daten aus dem MDE Gerät "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   117
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Frame5"
      Height          =   375
      Left            =   10560
      TabIndex        =   102
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
      Begin VB.ComboBox cbofil3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4320
         Style           =   2  'Dropdown-Liste
         TabIndex        =   144
         Top             =   5760
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   125
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   8835
         TabIndex        =   124
         Top             =   6360
         Width           =   8895
      End
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   10
         Left            =   9360
         TabIndex        =   105
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   9
         Left            =   9360
         TabIndex        =   104
         Top             =   5520
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   6
         Left            =   9720
         TabIndex        =   103
         Top             =   6960
         Width           =   1935
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
         BackColor       =   &H00C0C000&
         Caption         =   "Bitte nicht  die Waren mehrerer  Filialen im MDE - Gerät aufnehmen!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   2040
         TabIndex        =   147
         Top             =   4560
         Visible         =   0   'False
         Width           =   7935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000001&
         Caption         =   "Filialauswahl:"
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
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   146
         Top             =   5880
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Aus welcher Filiale kommt die Ware?"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   145
         Top             =   5520
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWKL23.frx":0442
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   5
         Left            =   2040
         TabIndex        =   122
         Top             =   2160
         Width           =   7935
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10800
         MouseIcon       =   "frmWKL23.frx":0549
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL23.frx":0853
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Daten aus dem MDE Gerät "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   6135
      End
   End
   Begin sevCommand3.Command Command11 
      Height          =   525
      Index           =   0
      Left            =   9480
      TabIndex        =   69
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
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   375
      Left            =   240
      TabIndex        =   68
      Top             =   960
      Width           =   855
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8160
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Expresssendungen annehmen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   840
         TabIndex        =   129
         Top             =   3120
         Width           =   6615
      End
      Begin sevCommand3.Command Command11 
         Height          =   525
         Index           =   1
         Left            =   9360
         TabIndex        =   1
         Top             =   6240
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
         Caption         =   "weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "mit dem MDE - Gerät"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   840
         TabIndex        =   73
         Top             =   2520
         Width           =   6615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "mittels Datei (altes Verfahren)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   72
         Top             =   1920
         Width           =   6615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "manuell mit Scanner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   71
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label Label6 
         Caption         =   "Wie möchten Sie Ihren Wareneingang bearbeiten?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   70
         Top             =   600
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   495
      Left            =   12360
      TabIndex        =   23
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6060
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   9135
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9135
      End
      Begin sevCommand3.Command Command4 
         Height          =   525
         Index           =   0
         Left            =   9360
         TabIndex        =   11
         Top             =   5520
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   525
         Index           =   1
         Left            =   9360
         TabIndex        =   12
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame6"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   21
         Left            =   8160
         TabIndex        =   142
         Top             =   3360
         Width           =   1095
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
         Caption         =   "Protokoll löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   14
         Left            =   9360
         TabIndex        =   126
         Top             =   3360
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
         Caption         =   "Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   13
         Left            =   9360
         TabIndex        =   123
         Top             =   4080
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
         Caption         =   "Holen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   7
         Left            =   9720
         TabIndex        =   61
         Top             =   6960
         Width           =   1935
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
      Begin sevCommand3.Command cmdUpdate 
         Height          =   495
         Left            =   6720
         TabIndex        =   59
         Top             =   720
         Width           =   1575
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdStandardUp 
         Height          =   495
         Left            =   8400
         TabIndex        =   58
         Top             =   720
         Width           =   1575
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtZinPfad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   9855
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   5
         Left            =   9360
         TabIndex        =   56
         Top             =   4800
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   4
         Left            =   9360
         TabIndex        =   55
         Top             =   5520
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   3
         Left            =   9360
         TabIndex        =   54
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   5400
         Pattern         =   "WV*.dbf"
         TabIndex        =   52
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Dateien aus der Zentrale bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Pfad zu den Wareneingängen"
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
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   840
         Width           =   6135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C000&
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   150
         Top             =   5400
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Etiketten erzeugen"
         Height          =   255
         Left            =   9360
         TabIndex        =   141
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Druckdaten löschen"
         Height          =   255
         Left            =   9360
         TabIndex        =   140
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Height          =   3015
         Left            =   120
         TabIndex        =   79
         Top             =   3960
         Visible         =   0   'False
         Width           =   720
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   16
            Left            =   4200
            TabIndex        =   99
            Top             =   1935
            Width           =   2160
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "<<<"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "1"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   1
            Left            =   960
            TabIndex        =   97
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "2"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   2
            Left            =   1800
            TabIndex        =   96
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "3"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   3
            Left            =   2640
            TabIndex        =   95
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "4"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   4
            Left            =   3480
            TabIndex        =   94
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "5"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   5
            Left            =   4320
            TabIndex        =   93
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "6"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   6
            Left            =   5160
            TabIndex        =   92
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "7"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   7
            Left            =   6000
            TabIndex        =   91
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "8"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   8
            Left            =   6840
            TabIndex        =   90
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "9"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   840
            Index           =   9
            Left            =   7680
            TabIndex        =   89
            Top             =   240
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "0"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   17
            Left            =   6360
            TabIndex        =   88
            Top             =   1935
            Width           =   2160
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   ">>>"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   11
            Left            =   120
            TabIndex        =   87
            Top             =   1080
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "+"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   12
            Left            =   960
            TabIndex        =   86
            Top             =   1080
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "-"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   13
            Left            =   1800
            TabIndex        =   85
            Top             =   1080
            Width           =   2520
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "Leeren"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   14
            Left            =   4320
            TabIndex        =   84
            Top             =   1080
            Width           =   2520
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "Rückgängig"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   18
            Left            =   6840
            TabIndex        =   83
            Top             =   1080
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   ","
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   19
            Left            =   7680
            TabIndex        =   82
            Top             =   1080
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   20
            Left            =   3360
            TabIndex        =   81
            Top             =   1935
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "F4"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   855
            Index           =   10
            Left            =   2520
            TabIndex        =   80
            Top             =   1935
            Width           =   840
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
            Caption         =   "00"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Label3"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   21
         Left            =   8640
         TabIndex        =   78
         Top             =   120
         Width           =   615
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   15
         Left            =   9360
         TabIndex        =   75
         Top             =   4350
         Width           =   2160
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
         Caption         =   "Speichern"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Left            =   9360
         TabIndex        =   74
         Top             =   6500
         Width           =   2160
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
      Begin sevCommand3.Command Command10 
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   840
         Width           =   480
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List10 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "frmWKL23.frx":0E36
         Left            =   240
         List            =   "frmWKL23.frx":0E38
         TabIndex        =   64
         Top             =   4440
         Width           =   7335
      End
      Begin sevCommand3.Command Command9 
         Height          =   375
         Left            =   9360
         TabIndex        =   63
         Top             =   6100
         Width           =   2160
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
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   375
         Left            =   5040
         TabIndex        =   62
         Top             =   1650
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "in Filialen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   9720
         MaxLength       =   9
         TabIndex        =   34
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   9720
         MaxLength       =   9
         TabIndex        =   33
         Top             =   3360
         Width           =   1815
      End
      Begin sevCommand3.Command Command7 
         Height          =   375
         Left            =   9360
         TabIndex        =   8
         Top             =   3960
         Width           =   2160
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
         Caption         =   "Daten ändern"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2280
         MaxLength       =   13
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefNr leeren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   6480
         TabIndex        =   27
         Top             =   3600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefNr halten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   6480
         TabIndex        =   26
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2280
         MaxLength       =   13
         TabIndex        =   2
         Top             =   120
         Width           =   4935
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Left            =   9360
         TabIndex        =   3
         Top             =   120
         Width           =   2160
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
      Begin VB.ListBox List14 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         ItemData        =   "frmWKL23.frx":0E3A
         Left            =   240
         List            =   "frmWKL23.frx":0E3C
         TabIndex        =   143
         Top             =   4200
         Width           =   7335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferschein:"
         Height          =   255
         Index           =   6
         Left            =   9360
         TabIndex        =   151
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "insgsamt:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   67
         Top             =   6600
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "zuletzt gebuchte Wareneingänge"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   65
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   11520
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "neuer Kassen-VK:"
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
         Index           =   12
         Left            =   6960
         TabIndex        =   36
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EK-Preis:"
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
         Index           =   11
         Left            =   8160
         TabIndex        =   35
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestNr:"
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
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Min.Bestand:"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   9720
         TabIndex        =   29
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "alter Kassen-Vk:"
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
         Index           =   9
         Left            =   7080
         TabIndex        =   28
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "unbekannt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   25
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant:"
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
         Index           =   8
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Zugang:"
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
         Left            =   -240
         TabIndex        =   22
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Listen-Vk:"
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
         Index           =   5
         Left            =   7680
         TabIndex        =   21
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   9720
         TabIndex        =   20
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand:"
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
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "unbekannt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   18
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel:"
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
         Index           =   4
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EAN / ArtNr:"
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
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame4"
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
      Left            =   240
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
      Begin MSComctlLib.ProgressBar pbr 
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   6360
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   2
         Left            =   10440
         TabIndex        =   49
         Top             =   4560
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List7 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   6000
         TabIndex        =   44
         Top             =   1080
         Width           =   5535
      End
      Begin VB.ListBox List8 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   45
         Top             =   840
         Width           =   5535
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   5535
      End
      Begin VB.ListBox List6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   0
         Left            =   9360
         TabIndex        =   38
         Top             =   6240
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   1
         Left            =   9360
         TabIndex        =   39
         Top             =   5520
         Visible         =   0   'False
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00C0C000&
         Caption         =   "mit Etikettenerstellung"
         Height          =   210
         Left            =   9360
         TabIndex        =   40
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   23
         Left            =   4560
         TabIndex        =   152
         Top             =   4560
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   153
         Top             =   4920
         Width           =   3735
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Dateiname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6600
         TabIndex        =   127
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Dateien aus der Zentrale bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   3
         Left            =   6000
         TabIndex        =   48
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   4560
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel, die nicht zugeordnet werden können"
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
         Index           =   0
         Left            =   6000
         TabIndex        =   46
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel der Warenverteilung"
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
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   76
      Top             =   7800
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Wareneingang aus Umverteilung"
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
      Left            =   240
      TabIndex        =   32
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gbDrueck5 As Boolean
Dim dbwv As Database
Dim gcdatei As String

Dim bfoundauto As Boolean
Dim fromMde As Boolean
Dim bscanner As Boolean
Dim iPos    As Integer
Dim iSum    As Integer
Dim iNeg    As Integer
Private Sub PositionierenWKL23()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 3960
    Frame1.Left = 120
    Frame1.Height = 3015
    Frame1.Width = 8800
    Frame1.BorderStyle = 0
    
    
    Frame2.Top = 840
    Frame2.Left = 120
    Frame2.Height = 6855
    Frame2.Width = 11655
    Frame2.BorderStyle = 0
    
    Frame3.Top = 840
    Frame3.Left = 120
    Frame3.Height = 6855
    Frame3.Width = 11655
    Frame3.BorderStyle = 0
    
    Frame4.Top = 840
    Frame4.Left = 120
    Frame4.Height = 6855
    Frame4.Width = 11655
    Frame4.BorderStyle = 0
    
    Frame5.Top = 840
    Frame5.Left = 120
    Frame5.Height = 6855
    Frame5.Width = 11655
    Frame5.BorderStyle = 0
    
    Frame6.Top = 840
    Frame6.Left = 120
    Frame6.Height = 6855
    Frame6.Width = 11655
    Frame6.BorderStyle = 0
    
    Frame7.Top = 840
    Frame7.Left = 120
    Frame7.Height = 6855
    Frame7.Width = 11655
    Frame7.BorderStyle = 0
    
    Frame8.Top = 840
    Frame8.Left = 120
    Frame8.Height = 6855
    Frame8.Width = 11655
    Frame8.BorderStyle = 0
    
    Frame9.Top = 840
    Frame9.Left = 120
    Frame9.Height = 6855
    Frame9.Width = 11655
    Frame9.BorderStyle = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositioniereWKL23"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL15()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = gsWeEinzMe
'    Text1(1).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    If Option1(1).value Then
        Text1(4).Text = ""
        Label2(4).Caption = ""
    End If
    Label2(0).Caption = ""
    Label2(1).Caption = "0"
    Label2(2).Caption = "0"
    Label2(3).Caption = "0,00 " & gcWaehrung
    Label2(5).Caption = "0,00 " & gcWaehrung
    
    Label3.Caption = "0"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseLieferantenPreisWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLinr As String
    Dim cArtNr As String
    Dim dLEKPR As Double
    
    cArtNr = Label2(2).Caption
    cLinr = Trim$(Str$(Val(Text1(4).Text)))
    
    cSQL = "Select * from ARTLIEF where LINR = " & cLinr & " and ARTNR = " & cArtNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!lekpr) Then
            dLEKPR = rsrs!lekpr
        Else
            dLEKPR = 0
        End If
    Else
        dLEKPR = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If dLEKPR <> 0 Then
        Text1(3).Text = Format$(dLEKPR, "#####0.00")
    Else
        Text1(3).Text = ""
    End If

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseLieferantenPreisWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SchreibeDatenWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim lMengeSchwietz As Long
    Dim lMinBest       As Long
    Dim lLpz           As Long
    Dim lHeute         As Long
    Dim ctmp           As String
    Dim cSQL           As String
    Dim cMeld          As String
    Dim cArtNr         As String
    Dim cBezeich       As String
    Dim cEtiMerk       As String
    Dim cLiBesNr       As String
    Dim cLiBesNrDialog As String
    Dim cEkPr          As String
    Dim cLinr          As String
    Dim cLiNrDialog    As String
    Dim cJetzt         As String
    Dim cEAN           As String
    Dim cArtNrSchwietz As String
    Dim cPreis         As String
    
    Dim dAnzahl        As Double
    Dim dVkPr          As Double
    Dim dBestand       As Double
    Dim dPreis         As Double
    Dim dEkpr          As Double
    Dim dEkPrAlt       As Double
    Dim dEkPrSchnitt   As Double
    Dim dWertAlt       As Double
    Dim dWertNeu       As Double
    Dim dWert          As Double
    Dim dBWert         As Double
    Dim dAlt           As Double
    Dim iArtAnzahl     As Integer
    Dim iFehlerstufe   As Integer
    Dim iRet           As Integer
    Dim bNeu           As Boolean
    Dim bTrans         As Boolean
    Dim rsA            As Recordset
    Dim rsZ            As Recordset
    Dim rsrs           As Recordset
    Dim rsHis          As Recordset
    Dim rsArtlief      As Recordset
    Dim rsZutemp       As Recordset
    Dim i              As Integer
    Dim iZBestand      As Integer
    Dim cKVKVergleichsPreis As String
   
    cArtNr = Label2(2).Caption
    
    ctmp = Trim$(Text1(1).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dBWert = Val(ctmp)
    ctmp = Trim$(Text1(2).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    iFehlerstufe = 0
    bTrans = False
    
    cLiNrDialog = Text1(4).Text
    cLiBesNrDialog = Text1(6).Text
    cKVKVergleichsPreis = Trim$(Text1(2).Text)
    cPreis = Text1(2).Text
    cPreis = Trim$(cPreis)
    If cPreis <> "" Then
        If InStr(cPreis, ",") = 0 Then
            cPreis = Format$((Val(cPreis) / 100), "#####0.00")
        End If
        cPreis = fnMoveComma2Point$(cPreis)
        dPreis = Val(cPreis)
        If dPreis > 100000 Then
            MsgBox "Der eingegebene Preis ist zu hoch!", vbCritical, "STOP!"
            Text1(2).SetFocus
            Exit Sub
        End If
    End If
    cEkPr = Text1(3).Text
    cEkPr = Trim$(cEkPr)
    If cEkPr <> "" Then
        If InStr(cEkPr, ",") = 0 Then
            cEkPr = Format$((Val(cEkPr) / 100), "#####0.00")
        End If
        cEkPr = fnMoveComma2Point$(cEkPr)
        dEkpr = Val(cEkPr)
        If dEkpr > 100000 Then
            MsgBox "Der eingegebene Preis ist zu hoch!", vbCritical, "STOP!"
            Text1(3).SetFocus
            Exit Sub
        End If
    Else
        dEkpr = 0
    End If
    lMinBest = Val(Text1(5).Text)
    ctmp = Trim$(Text1(1).Text)
    
    If ctmp = "" Then
        ctmp = "0"
    End If
    dAnzahl = Val(ctmp)
    
    cArtNr = Label2(2).Caption
    cArtNr = Trim$(cArtNr)
    If cArtNr = "" Then
        MsgBox "Artikel-Nr fehlt! Daten speichern nicht möglich!", vbCritical, "FEHLER2"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    iFehlerstufe = 1
    
    cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLiNrDialog & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        iRet = MsgBox("Verbindung von Artikel und Lieferant ist neu!" & vbCrLf & vbCrLf & "Wollen Sie diese neue Verbindung speichern?", vbQuestion + vbYesNo, "NEUE VERBINDUNG")
        If iRet = vbNo Then
            rsrs.Close: Set rsrs = Nothing
            Exit Sub
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    iFehlerstufe = 2
    cSQL = "Select * from ARTIKEL where ARTNR = " & cArtNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount <> 1 Then
            MsgBox "Mehr als 1 Artikeleintrag gefunden! Daten speichern nicht möglich!", vbCritical, "FEHLER2"
            Text1(0).SetFocus
            rsrs.Close: Set rsrs = Nothing
            Exit Sub
        End If
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!BEZEICH) Then
            cBezeich = rsrs!BEZEICH
        Else
            cBezeich = ""
        End If
        If Not IsNull(rsrs!vkpr) Then
            dVkPr = rsrs!vkpr
        Else
            dVkPr = 0
        End If
        If Not IsNull(rsrs!BESTAND) Then
            dAlt = rsrs!BESTAND
        Else
            dAlt = 0
        End If
        
        If Not IsNull(rsrs!ETIMERK) Then
            cEtiMerk = rsrs!ETIMERK
        Else
            cEtiMerk = ""
        End If
        If Not IsNull(rsrs!LIBESNR) Then
            cLiBesNr = rsrs!LIBESNR
        Else
            cLiBesNr = ""
        End If
        If Not IsNull(rsrs!EAN) Then
            cEAN = rsrs!EAN
        Else
            cEAN = ""
        End If
        
        If Not IsNull(rsrs!linr) Then
            cLinr = rsrs!linr
        Else
            cLinr = ""
        End If
        
        If Not IsNull(rsrs!LPZ) Then
            lLpz = rsrs!LPZ
            If bNeu = True Then  '** fort. ETIDRU **
                Set rsA = gdBase.OpenRecordset("ETIDRU", dbOpenTable)
                rsA.index = "ARTNR"
                rsA.Seek "=", cArtNr
                If Not rsA.NoMatch Then
                    rsA.Edit
                    rsA!LPZ = lLpz
                    rsA.Update
                End If
                rsA.Close
            End If
        Else
            lLpz = 0
        End If
        
        If Not IsNull(rsrs!ekpr) Then
            dEkPrAlt = rsrs!ekpr
        Else
            dEkPrAlt = 0
        End If
        

        iFehlerstufe = 3
        cSQL = "Select * from umlager where ARTNR = -1"
        Set rsHis = gdBase.OpenRecordset(cSQL)
        
        cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLiNrDialog & " "
        Set rsArtlief = gdBase.OpenRecordset(cSQL)
        
        lHeute = Fix(Now)
        cJetzt = Format$(Now, "HH:MM")
        
        If Not tableSuchenDBKombi("ZuAusUV", 2) Then
            cSQL = "Create Table ZuAusUV "
            cSQL = cSQL & "( "
            cSQL = cSQL & "ARTNR LONG"
            cSQL = cSQL & ", BEZEICH Text (35) "
            cSQL = cSQL & ", EAN TEXT (13) "
            cSQL = cSQL & ", LINR LONG "
            cSQL = cSQL & ", ADATE DATETIME "
            cSQL = cSQL & ", UHRZEIT TEXT (5) "
            cSQL = cSQL & ", BEDNU long "
            cSQL = cSQL & ", BEDNAME TEXT (32) "
            cSQL = cSQL & ", FILIALNR BYTE "
            cSQL = cSQL & ", BESTANDALT INTEGER "
            cSQL = cSQL & ", BEWEGUNG INTEGER "
            cSQL = cSQL & ", BESTANDNEU INTEGER "
            cSQL = cSQL & ", EKPR SINGLE "
            cSQL = cSQL & ", KVKPR1 SINGLE "
            cSQL = cSQL & ", LIBESNR TEXT (13) "
            cSQL = cSQL & ", Lfnr autoincrement "
            cSQL = cSQL & ") "
            gdApp.Execute cSQL, dbFailOnError
        Else
            If Not SpalteInTabellegefundenNEW("ZuAusUV", "KVKPR1", gdApp) Then
                SpalteAnfuegenNEW "ZuAusUV", "KVKPR1", "single", gdApp
            End If
        End If
               
        iFehlerstufe = 31
               
        Set rsZutemp = gdApp.OpenRecordset("ZuAusUV", dbOpenTable)
        
        Bestandsveraenderung cArtNr, CLng(dAlt + dAnzahl), "WE aus Umverteilung"
         
        rsrs.Edit
        rsrs!SYNStatus = "E"

        Dim bKVKPR1     As Boolean
        bKVKPR1 = False

        If cPreis <> "" Then
        
            'Hat sich der KVKPR1 geändert
    
            If Not IsNull(rsrs!KVKPR1) Then
                If Trim(CStr(rsrs!KVKPR1)) <> Trim$(cKVKVergleichsPreis) Then
                    rsrs!KVKPR1 = CDbl(cKVKVergleichsPreis)
                    Artikelveraenderung cArtNr, Trim$(cKVKVergleichsPreis), "WE aus Umverteilung", "KVKPR1"
                    
                End If
            Else
                If IsNull(rsrs!KVKPR1) Then
                    bKVKPR1 = True
                Else
                    bKVKPR1 = False

                    rsrs!KVKPR1 = CDbl(cKVKVergleichsPreis)
                End If
            End If
        End If
        
        rsrs!MINBEST = lMinBest
        rsrs!GEFUEHRT = "J"
        
        If cLinr = cLiNrDialog Then
            rsrs!LIBESNR = cLiBesNrDialog
        End If
        
        iFehlerstufe = 4
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs.Update
        
        If bKVKPR1 Then
            Artikelveraenderung cArtNr, Trim$(cKVKVergleichsPreis), "WE aus Umverteilung", "KVKPR1"
        End If
                
        iFehlerstufe = 5
        rsHis.AddNew
        rsHis!artnr = Val(cArtNr)
        iFehlerstufe = 501
        rsHis!BEZEICH = cBezeich
        iFehlerstufe = 502
        rsHis!linr = Val(cLiNrDialog)
        iFehlerstufe = 503
        rsHis!EAN = cEAN
        iFehlerstufe = 504
        rsHis!ADATE = lHeute
        iFehlerstufe = 505
        rsHis!Uhrzeit = cJetzt
        iFehlerstufe = 506
        rsHis!BEDNU = Val(gcBedienerNr)
        iFehlerstufe = 507
        rsHis!bedname = gcUserName
        iFehlerstufe = 508
        rsHis!FILIALNR = 1
        iFehlerstufe = 509
        rsHis!bestandalt = dAlt
        iFehlerstufe = 510
        rsHis!BEWEGUNG = dAnzahl
        iFehlerstufe = 511
        rsHis!BESTANDneu = dAlt + dAnzahl
        iFehlerstufe = 512
        rsHis!ekpr = Val(cEkPr)
        iFehlerstufe = 513
        rsHis.Update
        rsHis.Close: Set rsHis = Nothing

        If KundenbestBestätigung(cArtNr, dAnzahl) = True Then
        
            Command11(2).Visible = True
            anzeige "ERFOLG", "", Label5
            
        End If
        
        rsZutemp.AddNew
        rsZutemp!artnr = Val(cArtNr)
        rsZutemp!BEZEICH = cBezeich
        rsZutemp!linr = Val(cLiNrDialog)
        rsZutemp!EAN = cEAN
        rsZutemp!ADATE = lHeute
        rsZutemp!Uhrzeit = cJetzt
        rsZutemp!BEDNU = Val(gcBedienerNr)
        rsZutemp!bedname = gcUserName
        rsZutemp!FILIALNR = 1
        rsZutemp!bestandalt = dAlt
        rsZutemp!BEWEGUNG = dAnzahl
        rsZutemp!BESTANDneu = dAlt + dAnzahl
        rsZutemp!ekpr = Val(cEkPr)
        rsZutemp!KVKPR1 = rsrs!KVKPR1
        rsZutemp!LIBESNR = cLiBesNrDialog
        rsZutemp.Update
        rsZutemp.Close: Set rsZutemp = Nothing
        
        iFehlerstufe = 6
        If rsArtlief.EOF Then
            rsArtlief.AddNew
        Else
            rsArtlief.Edit
        End If
        
        rsArtlief!artnr = Val(cArtNr)
        rsArtlief!linr = Val(cLiNrDialog)
        rsArtlief!lekpr = Val(cEkPr)
        rsArtlief!LIBESNR = cLiBesNrDialog
        rsArtlief.Update
        
'''        CommitTrans
'''        bTrans = False
'''        '***** Ende TRANSAKTIONS-Klammerung *****
        
        rsArtlief.Close: Set rsArtlief = Nothing
        
        iFehlerstufe = 7
        
        If Option2(3).value = True Then
        
        Else
        
            anzeige "LASER", "", Label5
            
        End If
        
        
    Else
        MsgBox "Keine Artikeldaten gefunden! Daten speichern nicht möglich!", vbCritical, "FEHLER2"
        Text1(0).SetFocus
        rsrs.Close: Set rsrs = Nothing
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Trim$(Text1(2).Text) <> Trim$(Left(Label2(5).Caption, Len(Label2(5).Caption) - 3)) Then
        If (Trim$(Label2(1).Caption) <> Trim$(Text1(1).Text)) And (Text1(1).Text <> 0) Then
            '** es gibt neue KVKPR1 UND neuer BESTAND **
            If gcFilNr = "1" Then
                dBestand = dAlt + dAnzahl '**  ARTIKEL **
                For i = 1 To giAnzFil
                    If i > 1 Then
                        iZBestand = ermBestandfromZbestand(cArtNr, i)
                    End If

                        cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
                        cSQL = cSQL & " and FILNR = " & i
                        Set rsrs = gdBase.OpenRecordset(cSQL)
                        If Not rsrs.EOF Then
                           rsrs.Edit
                        Else
                            rsrs.AddNew
                        End If
                        
                        rsrs!artnr = cArtNr
                        rsrs!BEZEICH = cBezeich
                        If cPreis <> "" Then
                            rsrs!vkpr = Val(cPreis)
                        Else
                            rsrs!vkpr = dVkPr
                        End If
                        If Not IsNull(rsrs!BESTAND) Then
                            If i = 1 Then
                                rsrs!BESTAND = rsrs!BESTAND + dBestand
                            Else
                                rsrs!BESTAND = rsrs!BESTAND + iZBestand
                            End If
                        Else
                            If i = 1 Then
                                rsrs!BESTAND = dBestand
                            Else
                                rsrs!BESTAND = iZBestand
                            End If
                        End If
                        If Not IsNull(rsrs!ANZAHL) Then
                            If i = 1 Then
                                rsrs!ANZAHL = rsrs!ANZAHL + dBestand
                            Else
                                rsrs!ANZAHL = rsrs!ANZAHL + iZBestand
                            End If
                        Else
                            If i = 1 Then
                                rsrs!ANZAHL = dBestand
                            Else
                                rsrs!ANZAHL = iZBestand
                            End If
                        End If
                        rsrs!LIBESNR = cLiBesNr
                        rsrs!EAN = cEAN
                        rsrs!linr = Val(cLinr)
                        rsrs!LPZ = lLpz
                        rsrs!filnr = i
                        rsrs!Pcname = srechnertab
                        rsrs.Update
                        rsrs.Close: Set rsrs = Nothing
weiter:
                Next i
            Else
                cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
                cSQL = cSQL & " and FILNR = " & gcFilNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                   rsrs.Edit
                Else
                    rsrs.AddNew
                End If
                
                rsrs!artnr = cArtNr
                rsrs!BEZEICH = cBezeich
                If cPreis <> "" Then
                    rsrs!vkpr = Val(cPreis)
                Else
                    rsrs!vkpr = dVkPr
                End If
                If Not IsNull(rsrs!BESTAND) Then
                    rsrs!BESTAND = rsrs!BESTAND + dAnzahl
                Else
                    rsrs!BESTAND = dAnzahl
                End If
                If Not IsNull(rsrs!ANZAHL) Then
                    rsrs!ANZAHL = rsrs!ANZAHL + dAnzahl
                Else
                    rsrs!ANZAHL = dAnzahl
                End If
                rsrs!LIBESNR = cLiBesNr
                rsrs!EAN = cEAN
                rsrs!linr = Val(cLinr)
                rsrs!LPZ = lLpz
                rsrs!filnr = gcFilNr
                rsrs!Pcname = srechnertab
                rsrs.Update
                rsrs.Close: Set rsrs = Nothing
            End If
        Else '** nur Preisveränderung,kein Bestandveränderung **
            If gcFilNr = "1" Then
                dBestand = dAlt + dAnzahl '**  ARTIKEL **
                For i = 1 To giAnzFil
                
                    If i > 1 Then
                        iZBestand = ermBestandfromZbestand(cArtNr, i)
                    End If
                    
                    cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
                    cSQL = cSQL & " and FILNR = " & i
                    Set rsrs = gdBase.OpenRecordset(cSQL)
                    If Not rsrs.EOF Then
                       rsrs.Edit
                    Else
                        rsrs.AddNew
                    End If
                    
                    rsrs!artnr = cArtNr
                    rsrs!BEZEICH = cBezeich
                    If cPreis <> "" Then
                        rsrs!vkpr = Val(cPreis)
                    Else
                        rsrs!vkpr = dVkPr
                    End If
                    If Not IsNull(rsrs!BESTAND) Then
                        If i = 1 Then
                            rsrs!BESTAND = rsrs!BESTAND + dBestand
                        Else
                            rsrs!BESTAND = rsrs!BESTAND + iZBestand
                        End If
                    Else
                        If i = 1 Then
                            rsrs!BESTAND = dBestand
                        Else
                            rsrs!BESTAND = iZBestand
                        End If
                    End If
                    If Not IsNull(rsrs!ANZAHL) Then
                        If i = 1 Then
                            rsrs!ANZAHL = rsrs!ANZAHL + dBestand
                        Else
                            rsrs!BESTAND = rsrs!BESTAND + iZBestand
                        End If
                    Else
                        If i = 1 Then
                            rsrs!ANZAHL = dBestand
                        Else
                            rsrs!ANZAHL = iZBestand
                        End If
                    End If
                    rsrs!LIBESNR = cLiBesNr
                    rsrs!EAN = cEAN
                    rsrs!linr = Val(cLinr)
                    rsrs!LPZ = lLpz
                    rsrs!filnr = i
                    rsrs!Pcname = srechnertab
                    rsrs.Update
                    rsrs.Close: Set rsrs = Nothing
weiter1:
                Next i
            Else
                cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
                cSQL = cSQL & " and FILNR = " & gcFilNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                   rsrs.Edit
                Else
                    rsrs.AddNew
                End If
                
                rsrs!artnr = cArtNr
                rsrs!BEZEICH = cBezeich
                If cPreis <> "" Then
                    rsrs!vkpr = Val(cPreis)
                Else
                    rsrs!vkpr = dVkPr
                End If
                If Not IsNull(rsrs!BESTAND) Then
                    rsrs!BESTAND = rsrs!BESTAND + dAnzahl
                Else
                    rsrs!BESTAND = dAnzahl
                End If
                If Not IsNull(rsrs!ANZAHL) Then
                    rsrs!ANZAHL = rsrs!ANZAHL + dAnzahl
                Else
                    rsrs!ANZAHL = dAnzahl
                End If
                rsrs!LIBESNR = cLiBesNr
                rsrs!EAN = cEAN
                rsrs!linr = Val(cLinr)
                rsrs!LPZ = lLpz
                rsrs!filnr = gcFilNr
                rsrs!Pcname = srechnertab
                rsrs.Update
            End If
        
        End If
    Else '** nur Bestandveränderung,kein Preisveränderung  **
        If Check1.value = vbChecked Or Check2.value = vbChecked Then
            cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
            cSQL = cSQL & " and FILNR = " & gcFilNr
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
               rsrs.Edit
            Else
                rsrs.AddNew
            End If
            rsrs!artnr = cArtNr
            rsrs!BEZEICH = cBezeich
            If cPreis <> "" Then
                rsrs!vkpr = Val(cPreis)
            Else
                rsrs!vkpr = dVkPr
            End If
            If Not IsNull(rsrs!BESTAND) Then
                rsrs!BESTAND = rsrs!BESTAND + dAnzahl
            Else
                rsrs!BESTAND = dAnzahl
            End If
            If Not IsNull(rsrs!ANZAHL) Then
                rsrs!ANZAHL = rsrs!ANZAHL + dAnzahl
            Else
                rsrs!ANZAHL = dAnzahl
            End If
            rsrs!LIBESNR = cLiBesNr
            rsrs!EAN = cEAN
            rsrs!linr = Val(cLinr)
            rsrs!LPZ = lLpz
            rsrs!filnr = gcFilNr
            rsrs!Pcname = srechnertab
            rsrs.Update
            rsrs.Close: Set rsrs = Nothing
        End If
    End If
    
    If Check4.value = vbChecked Then
        schreibeWKEtidru cArtNr, CLng(dAnzahl), CLng(gcFilNr)
    End If
    
    LeereDialogWKL15
    
    If Not gbDrueck5 Then
        Text1(0).SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
    
End Sub
Private Sub SucheArtikelWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim bDebug As Boolean
    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim cSuch As String
    Dim cArtNr As String
    Dim cArtBez As String
    Dim dBestand As Double
    Dim dVkPr As Double
    Dim dKVKPR As Double
    Dim dEkpr As Double
    Dim dLEKPR As Double
    Dim cLinr As String
    Dim cLiefBez As String
    Dim cLiBesNr As String
    Dim bgefunden As Boolean
    Dim cFeld As String
    Dim cLBSatz As String
    Dim lMinBest As Long
    Dim bEAN As Boolean
    Dim cEAN As String
    
    bDebug = False
    bgefunden = True
    bEAN = True
    bfoundauto = False
    
    cSuch = Text1(0).Text
    cSuch = Trim$(cSuch)
    
    If cSuch = "" Then
        anzeigeNew "rot", "Bitte Wert eingeben!", Label5
        Text1(0).SetFocus
        Exit Sub
    End If
    
    anzeigeNew "normal", "", Label5
    
    cLinr = Text1(4).Text
    cLinr = Trim$(cLinr)
    
    If Len(cSuch) > 6 Then
        iRet = fnPruefeEANWert(cSuch)
        Select Case iRet
            Case Is = 0
                'alles okay
            Case Is = 1     'falsche Länge
                bEAN = False

            Case Is = 8     'falscher EAN-8
                bEAN = False

            Case Is = 12    'falscher UPC-A
                bEAN = False

            Case Is = 13    'falscher EAN-13
                bEAN = False

        End Select
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where A.ARTNR = B.ARTNR "
    End If
    
    If Len(cSuch) <= 6 Then
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where B.ARTNR = " & cSuch & " and A.ARTNR = B.ARTNR "
    Else
        If Len(cSuch) <= 8 And (Left(cSuch, 1) = "2") Then  'Or Left(cSuch, 1) = "0"
            cSuch = Mid(cSuch, 2, 6)
            cSQL = cSQL & "and B.ARTNR = " & cSuch & " "
        Else
            If bEAN Then
                cSQL = cSQL & "and (A.EAN = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN2 = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN3 = '" & cSuch & "' )"
            Else
                cSQL = cSQL & "and A.LIBESNR = '" & cSuch & " ' "
            End If
        End If
    End If
    
    If Len(cLinr) > 0 Then
        cSQL = cSQL & "and B.LINR = " & cLinr & " "
    End If
    
    bgefunden = False
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
        rsrs.Close: Set rsrs = Nothing
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            If Len(cSuch) = 8 And Left(cSuch, 1) = "2" Then
                cSuch = Mid(cSuch, 2, 6)
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If rsrs.EOF Then
                
                Else
                    bgefunden = True
                End If
            Else
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTLIEF where LIBESNR = '" & cSuch & "' "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveLast
                    If rsrs.RecordCount > 1 Then
                        List1.Clear
                        List2.Clear
                        List1.AddItem "ArtNr. Artikelbezeichnung                  EAN-Code      LiefNr LiefBestNr"
                        
                        rsrs.MoveFirst
                        Do While Not rsrs.EOF
                            If Not IsNull(rsrs!artnr) Then
                                cSuch = rsrs!artnr
                            Else
                                cSuch = "-1"
                            End If
                            cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                            Set rsRs2 = gdBase.OpenRecordset(cSQL)
                            If Not rsRs2.EOF Then
                                rsRs2.MoveFirst
                                If Not IsNull(rsRs2!artnr) Then
                                    cFeld = rsRs2!artnr
                                Else
                                    cFeld = ""
                                End If
                                cFeld = cFeld & Space$(6 - Len(cFeld))
                                cLBSatz = cFeld & " "
                                
                                If Not IsNull(rsRs2!BEZEICH) Then
                                    cFeld = rsRs2!BEZEICH
                                Else
                                    cFeld = ""
                                End If
                                cFeld = cFeld & Space$(35 - Len(cFeld))
                                cLBSatz = cLBSatz & cFeld & " "
                                
                                If Not IsNull(rsRs2!EAN) Then
                                    cFeld = rsRs2!EAN
                                Else
                                    cFeld = ""
                                End If
                                cFeld = cFeld & Space$(13 - Len(cFeld))
                                cLBSatz = cLBSatz & cFeld & " "
                                
                                If Not IsNull(rsRs2!linr) Then
                                    cFeld = rsRs2!linr
                                Else
                                    cFeld = ""
                                End If
                                cFeld = cFeld & Space$(6 - Len(cFeld))
                                cLBSatz = cLBSatz & cFeld & " "
                                
                                If Not IsNull(rsRs2!LIBESNR) Then
                                    cFeld = rsRs2!LIBESNR
                                Else
                                    cFeld = ""
                                End If
                                cFeld = cFeld & Space$(13 - Len(cFeld))
                                cLBSatz = cLBSatz & cFeld & " "
                                
                                List2.AddItem cLBSatz
                            End If
                            rsRs2.Close: Set rsRs2 = Nothing
                            rsrs.MoveNext
                        Loop
                        Frame1.Enabled = False
                        Frame3.Enabled = False
                        Frame2.Visible = True
                        Exit Sub
                    Else
                        rsrs.MoveFirst
                        If Not IsNull(rsrs!artnr) Then
                            Text1(0).Text = rsrs!artnr
                            Command1_Click
                            rsrs.Close: Set rsrs = Nothing
                            Exit Sub
                        End If
                    End If
                Else
                
                End If
            
            End If
        Else
            bgefunden = True
        End If
    Else
        bgefunden = True
    End If
    
    If bgefunden = True Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
        Else
            cArtNr = ""
        End If
        cArtNr = Trim$(cArtNr)
        Text1(0).Text = cArtNr
        
        If Not IsNull(rsrs!BEZEICH) Then
            cArtBez = rsrs!BEZEICH
        Else
            cArtBez = ""
        End If
        cArtBez = Trim$(cArtBez)
        
        If Not IsNull(rsrs!EAN) Then
            cEAN = rsrs!EAN
        Else
            cEAN = ""
        End If
        
        If Not IsNull(rsrs!linr) Then
            cLinr = rsrs!linr
        Else
            cLinr = "-1"
        End If
        cLinr = Trim$(cLinr)
        
    
        If Not IsNull(rsrs!BESTAND) Then
            dBestand = rsrs!BESTAND
        Else
            dBestand = 0
        End If
    
        If Not IsNull(rsrs!vkpr) Then
            dVkPr = rsrs!vkpr
        Else
            dVkPr = 0
        End If
    
        If Not IsNull(rsrs!KVKPR1) Then
            dKVKPR = rsrs!KVKPR1
        Else
            dKVKPR = 0
        End If
        
        If Not IsNull(rsrs!MINBEST) Then
            lMinBest = rsrs!MINBEST
        Else
            lMinBest = 0
        End If
        
        If Not IsNull(rsrs!LIBESNR) Then
            cLiBesNr = rsrs!LIBESNR
        Else
            cLiBesNr = ""
        End If
        cLiBesNr = Trim$(cLiBesNr)
        Text1(6).Text = cLiBesNr
        
        'Speichern aktivieren
        Command2(15).Enabled = True
    Else
        MsgBox "Artikel nicht gefunden!", vbInformation, "INFO"
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    If bgefunden Then
    
        cLiefBez = ermLiefBez(CLng(cLinr))
        Label2(0).Caption = cArtBez
        Label2(1).Caption = dBestand
        Label2(2).Caption = cArtNr
        Label2(3).Caption = Format$(dVkPr, "##,##0.00") & " " & gcWaehrung
        Label2(5).Caption = Format$(dKVKPR, "##,##0.00") & " " & gcWaehrung
        Text1(2).Text = Format$(dKVKPR, "#####0.00")
        If Trim$(Text1(4).Text) <> "" Then
            If Option1(1).value Then
                Text1(4).Text = cLinr
                Label2(4).Caption = cLiefBez
            End If
        Else
            Text1(4).Text = cLinr
            Label2(4).Caption = cLiefBez
        End If
        Text1(5).Text = Format$(lMinBest, "#####0")
        Text1(0).Text = cEAN
        If Not gbDrueck5 Then
            Text1(1).SetFocus
        End If
    End If
    
    If bgefunden = True Then
        LeseLieferantenPreisWKL15
    End If
    
    If bgefunden = True Then
        bfoundauto = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR

If Check3.value = vbChecked Then
    loeschapp "ZuAusUV"
    Command9.BackColor = Command3.BackColor
    Check3.value = vbUnchecked
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub cmdStandardUp_Click()
    On Error GoTo LOKAL_ERROR
    
    txtZinPfad.Text = gcDBPfad & "\Kissdata.mdb"
    gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    
    speicherpfad
    Dateienladen
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdStandardUp_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdUpdate_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim sPfad   As String

    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Speichern des Wareneingangpfades"
        .Filter = "Access - Dateien (*.mdb)|*.mdb"
        .ShowSave
    End With

    sPfad = cdlopen.FileName
    txtZinPfad.Text = sPfad
    
    gsZinPfad = sPfad
    
    speicherpfad
    Dateienladen
        
err:
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdUpdate_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cFeld As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim bTextSuche As Boolean
    Dim iFilnr  As Integer
    
    gbDrueck5 = False
    Command2(15).Enabled = False
    
    Screen.MousePointer = 11
    
    Command8.Visible = False
    cValid = "1234567890"
    cFeld = Text1(0).Text
    
    bTextSuche = False
    
    For lcount = 1 To Len(cFeld)
        cZeichen = Mid(cFeld, lcount, 1)
        If InStr(cValid, cZeichen) = 0 Then
            bTextSuche = True
            Exit For
        End If
    Next lcount
    
    If bTextSuche Then
        SucheTextArtikelWKL15
    Else
        SucheArtikelWKL15
    End If
    
    
    If bfoundauto And fromMde = False And bscanner Then
        bscanner = False
        bfoundauto = False
        Command2_Click 15
    Else
        fromMde = False
    End If
    
    
    iFilnr = CInt(gcFilNr)
    If iFilnr > 0 And Label2(2).Caption <> "" Then
        Command8.Visible = True
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheTextArtikelWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim cSuch As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim cFeld As String
    Dim cLBSatz As String
    
    List1.Clear
    List2.Clear
    List1.AddItem "ArtNr. Artikelbezeichnung                  EAN-Code      LiefNr LiefBestNr"
    
    cSuch = Text1(0).Text
    cSuch = UCase$(Trim$(cSuch))
    
    cSQL = "Select * from ARTIKEL where BEZEICH like '" & cSuch & "*' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(6 - Len(cFeld))
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(13 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(6 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!LIBESNR) Then
                cFeld = rsrs!LIBESNR
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(13 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    Else
        cSQL = "Select * from ARTLIEF where LIBESNR = '" & cSuch & "' "
        rsrs.Close: Set rsrs = Nothing
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveLast
            If rsrs.RecordCount > 1 Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                    If Not IsNull(rsrs!artnr) Then
                        cSuch = rsrs!artnr
                    Else
                        cSuch = "-1"
                    End If
                    cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                    Set rsRs2 = gdBase.OpenRecordset(cSQL)
                    If Not rsRs2.EOF Then
                        rsRs2.MoveFirst
                        If Not IsNull(rsRs2!artnr) Then
                            cFeld = rsRs2!artnr
                        Else
                            cFeld = ""
                        End If
                        cFeld = cFeld & Space$(6 - Len(cFeld))
                        cLBSatz = cFeld & " "
                        
                        If Not IsNull(rsRs2!BEZEICH) Then
                            cFeld = rsRs2!BEZEICH
                        Else
                            cFeld = ""
                        End If
                        cFeld = cFeld & Space$(35 - Len(cFeld))
                        cLBSatz = cLBSatz & cFeld & " "
                        
                        If Not IsNull(rsRs2!EAN) Then
                            cFeld = rsRs2!EAN
                        Else
                            cFeld = ""
                        End If
                        cLBSatz = cLBSatz & cFeld
                        
                        If Not IsNull(rsRs2!linr) Then
                            cFeld = rsRs2!linr
                        Else
                            cFeld = ""
                        End If
                        cFeld = cFeld & Space$(6 - Len(cFeld))
                        cLBSatz = cLBSatz & cFeld & " "
                        
                        If Not IsNull(rsRs2!LIBESNR) Then
                            cFeld = rsRs2!LIBESNR
                        Else
                            cFeld = ""
                        End If
                        cFeld = cFeld & Space$(13 - Len(cFeld))
                        cLBSatz = cLBSatz & cFeld & " "

                        List2.AddItem cLBSatz
                    End If
                    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
                    rsrs.MoveNext
                Loop
            Else
                rsrs.MoveFirst
                If Not IsNull(rsrs!artnr) Then
                    Text1(0).Text = rsrs!artnr
                    Command1_Click
                    rsrs.Close: Set rsrs = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Frame1.Enabled = False
    Frame3.Enabled = False
    
    Frame2.Visible = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheTextArtikelWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1_KeyUp 4, vbKeyF2, 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command11_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
        Case 0  'schließen
            Unload frmWKL23
        Case 1  'weiter
            Zeigeauswahlframe
        Case Is = 2 'Kundenbestellungen anzeigen
            KB "GELIEFERT", "INFORMIEREN"
            UpdateKuBestKUNDENSTATUS "INFORMIEREN", "GELIEFERT"
            Command11(2).Visible = False
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zeigeauswahlframe()
    On Error GoTo LOKAL_ERROR
    
    Frame7.Visible = False
    
    If Option2(0).value = True Then         'Manuell
        vorbereitungManuell
    ElseIf Option2(1).value = True Then     'Datei
        vorbereitungDatei
    ElseIf Option2(2).value = True Then     'Mde
        vorbereitungMDE
    ElseIf Option2(3).value = True Then     'Datei express
        vorbereitungExpress
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Dim lcount As Long
    Dim ctmp As String
    Dim iRet As Integer
    lcount = Val(Label3.Caption)
    
    Select Case index
        Case 0 To 10
            If lcount >= 0 Then
                Text1(lcount).Text = Text1(lcount).Text & Command2(index).Caption
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
            
        Case Is = 11        '** Plus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "+") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "-") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "+"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "+"
                End If
            End If
            Text1(lcount).SetFocus
            
        Case Is = 12        '** Minus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "-") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "+") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "-"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "-"
                End If
            End If
            Text1(lcount).SetFocus
        Case Is = 13        '** Löschen **
            Text1(lcount).Text = ""
            Text1(lcount).SetFocus
            
        Case Is = 14        '** Rückgängig **
            If Len(Text1(lcount).Text) > 0 Then
                ctmp = Text1(lcount).Text
                ctmp = Left(ctmp, Len(ctmp) - 1)
                Text1(lcount).Text = ctmp
            End If
            Text1(lcount).SetFocus
            
        Case Is = 15        '** Speichern **
        
            If Trim$(Text1(1).Text) = "" Then
                Text1(1).Text = "0"
            End If
            
            If Trim$(Text1(0).Text) = "" Then
                If Label2(2).Caption = "0" Then
                    anzeigeNew "rot", "Bitte einen Artikel festlegen!", Label5
                    Screen.MousePointer = 0
                    Text1(0).SetFocus
                    Exit Sub
                Else
                    Text1(0).Text = Label2(2).Caption
                End If
            End If
            
            If Trim$(Text1(4).Text) = "" Then
                anzeigeNew "rot", "Bitte einen Lieferanten festlegen!", Label5
                Screen.MousePointer = 0
                Text1(4).SetFocus
                Exit Sub
            End If
            
            glBestandNeu = Val(Text1(1).Text)
            If glBestandNeu < 0 Then
'                If glLevel < 7 Then
                    anzeigeNew "rot", "Mengen-Reduzierung nicht möglich!", Label5
                    Screen.MousePointer = 0
                    Exit Sub
'                End If
            ElseIf glBestandNeu = 0 Then
                anzeigeNew "rot", "Bitte Zugang eingeben!", Label5
                Screen.MousePointer = 0
                Text1(1).SetFocus
                Exit Sub
            End If
            gbDrueck5 = False
            
            
            
            If glBestandNeu > 999 Then
                iRet = MsgBox("Mengenangabe von " & glBestandNeu & " fraglich! Trotzdem speichern?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                If iRet = vbNo Then
                    Screen.MousePointer = 0
                    Text1(1).SetFocus
                    Exit Sub
                End If
            End If
            
            ABINFEEDBF Label2(2).Caption, CStr(glBestandNeu)
            SchreibeDatenWKL15
            SchreibeListe
            
            If NewTableSuchenDBKombi("ZuAusUV", gdApp) Then
                If Datendrin("ZuAusUV", gdApp) Then
                    Command9.BackColor = vbRed
                End If
            End If
            
        
        Case Is = 16        'Vorheriges Feld
            If lcount > 0 Then
                Text1(0).SetFocus
            Else
                Text1(lcount).SetFocus
            End If
            
        Case Is = 17        'Nächstes Feld
            If lcount < 1 Then
                Text1(1).SetFocus
            Else
                Text1(lcount).SetFocus
            End If
            
        Case Is = 18        'Komma
            If lcount = 2 Or lcount = 3 Then
                If InStr(Text1(lcount).Text, ",") = 0 Then
                    Text1(lcount).Text = Text1(lcount).Text & Command2(index).Caption
                End If
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
        Case Is = 19        'F2
            Text1_KeyUp Val(Label3.Caption), vbKeyF2, 0
            
        Case Is = 20        'F4
            If Text1(0).Text = "" Then
                MsgBox "Bitte den Artikel eindeutig definieren (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            Else
                Text1_KeyUp Val(Label3.Caption), vbKeyF4, 0
            End If
        Case Is = 21
            If Frame1.Visible Then
                Frame1.Visible = False
            Else
                Frame1.Visible = True
            End If
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeListe()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld       As String
    Dim cSatz       As String
    Dim rsZutemp    As Recordset
    Dim sSQL        As String
    
    
    If tableSuchenDBKombi("ZuAusUV", 2) Then
        If Not SpalteInTabellegefundenNEW("ZuAusUV", "lfnr", gdApp) Then
            SpalteAnfuegenNEW "ZuAusUV", "lfnr", "autoincrement", gdApp
        End If
        
        If Not SpalteInTabellegefundenNEW("ZuAusUV", "KVKPR1", gdApp) Then
            SpalteAnfuegenNEW "ZuAusUV", "KVKPR1", "single", gdApp
        End If
    End If
    
    
    sSQL = "Select * from ZuAusUV order by lfnr"
    
    List10.Clear
    Label11(0).Caption = "insgesamt: "
    Label11(0).Refresh
    
    If tableSuchenDBKombi("ZuAusUV", 2) Then
        Set rsZutemp = gdApp.OpenRecordset(sSQL)
        
        If Not rsZutemp.EOF Then
            Do While Not rsZutemp.EOF
                If Not IsNull(rsZutemp!ADATE) Then
                    cFeld = rsZutemp!ADATE
                End If
                cSatz = cFeld & Space(1)
                If Not IsNull(rsZutemp!Uhrzeit) Then
                    cFeld = rsZutemp!Uhrzeit
                End If
                cSatz = cSatz & cFeld & Space(1)
                
                If Not IsNull(rsZutemp!artnr) Then
                    cFeld = rsZutemp!artnr
                End If
                cSatz = cSatz & cFeld & Space(1)
                
                If Not IsNull(rsZutemp!BEZEICH) Then
                    cFeld = rsZutemp!BEZEICH
                End If
                cSatz = cSatz & cFeld & Space(40 - Len(cFeld))
                
                If Not IsNull(rsZutemp!BEWEGUNG) Then
                    cFeld = rsZutemp!BEWEGUNG
                End If
                cSatz = cSatz & Space(7 - Len(cFeld)) & cFeld & Space(2)
                
                cFeld = ""
                If Not IsNull(rsZutemp!KVKPR1) Then
                    cFeld = Format(rsZutemp!KVKPR1, "#####0.00")
                End If
                cSatz = cSatz & cFeld
                
                
                
                List10.AddItem cSatz, 0
                rsZutemp.MoveNext
                
            Loop
        End If
        rsZutemp.Close: Set rsZutemp = Nothing
        Label11(0).Caption = "insgesamt: " & List10.ListCount
        Label11(0).Refresh
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeListe"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame3.Visible = False
    Frame7.Visible = True
    
    check_ex
    
    anzeigeNew "normal", "", Label5
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    Select Case index
        Case Is = 0
            If List2.ListIndex < 0 Then
                MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbCritical, "STOP!"
            Else
                cLBSatz = List2.list(List2.ListIndex)
                cLBSatz = Trim$(cLBSatz)
                cLBSatz = Left(cLBSatz, 6)
                cLBSatz = Trim$(cLBSatz)
                If Len(cLBSatz) >= 13 Then
                    cLBSatz = Left(cLBSatz, 13)
                End If
                Text1(0).Text = Trim$(cLBSatz)
                Command4_Click 1
                Command1_Click
            End If
        Case Is = 1
            Frame3.Enabled = True
            Frame1.Enabled = True
            Frame2.Visible = False
            Text1(0).SetFocus
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub durchsuchen()
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad As String
    
    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Speichern des Pfades"
        .Filter = "Access - Dateien (*.mdb)|*.mdb"
        .ShowSave
    End With
            
    sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
    
    gsZinPfad = sPfad

err:

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "durchsuchen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub andererPfad()
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad As String
    
    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Speichern des Pfades"
        
        .Filter = "Access - Dateien (*.mdb)|*.mdb"
        .ShowSave
    End With
            
    sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
    
    gsZinPfad = sPfad
    speicherpfad

err:

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "andererPfad"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim rsMDE As Recordset
    Dim cSQL As String
    
    Select Case index
        Case Is = 0 'zurück Dateien Zentrale 1
            If Option2(3).value = True Then
                AlleZugriffeLöschen
                Frame4.Visible = False
                Frame9.Visible = True
                anzeigeNew "normal", "", Label5
                lbl6(6).Caption = ""
                lbl6(6).Refresh
            Else
                Frame4.Visible = False
                Frame6.Visible = True
                anzeigeNew "normal", "", Label5
                lbl6(6).Caption = ""
                lbl6(6).Refresh
            End If
        Case Is = 1
            einlesen gcdatei
            AlleZugriffeLöschen
        Case Is = 2
            reportbildschirm "umv1a", "aWKL23a"
        Case Is = 23
            reportbildschirm "umv1a", "aWKL23f"
        Case Is = 3     'zurück Dateien Zentrale
            Frame7.Visible = True
            check_ex
            Frame6.Visible = False
            anzeigeNew "normal", "", Label5
        Case Is = 4     'löschen
            löschen
        Case Is = 5     'auswählen
            auswählen
        Case Is = 7
            Unload frmWKL23
        Case Is = 8
            reportbildschirm "umv1a", "aWKL23c"
        Case Is = 9
        
            If MDEeinlesenOhneLinr(Label5, txtStatus, picprogress, frmWKL23) = False Then
                anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label5
            Else
                Frame5.Visible = False
                Frame8.Visible = True
                anzeigeNew "normal", "", Label5
        
                MdeVerarbeitung1
            End If
            
        Case Is = 10 'zurück aus MDE
            Frame5.Visible = False
            Frame7.Visible = True
            check_ex
            anzeigeNew "normal", "", Label5
            
        Case Is = 11 'zurück Dateien Zentrale 1
            Frame8.Visible = False
            Frame5.Visible = True
            anzeigeNew "normal", "", Label5
        Case Is = 12
        
            einlesenausMDE
        Case Is = 13 'Dateien holen
            giKissFtpMode = 13 ' FTPMODE= 13 Warenverteilungen holen
            frmWKL38.Show 1
            
            Dateienverarbeiten
            Dateienladen
        Case Is = 14
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "WarenVerteilungen.txt"
            Screen.MousePointer = 0
            
        Case Is = 15
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "WarenExpress.txt"
            Screen.MousePointer = 0
            
        Case Is = 16 'Dateien holen
        
            If gbWVNOT = False Then
                giKissFtpMode = 15 '  Expressverteilungen holen
                frmWKL38.Show 1
                
                DateienverarbeitenX
                DateienladenEX
            End If
            
            
        Case Is = 18     'auswählen x
        
            auswählenX
            
        Case Is = 19    'löschen von x
            löschenX
        Case Is = 20     'zurück Dateien Express
            Frame7.Visible = True
            check_ex
            Frame9.Visible = False
            Text3.Text = ""
            Text3.SetFocus
            anzeigeNew "normal", "", Label5
        Case 21 'Proto löschen
            
            cSQL = "Delete from Proto"
            gdBase.Execute cSQL, dbFailOnError
        Case Is = 22 ' Drucken der Liste
            drucken_Auflistung
    End Select
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub drucken_Auflistung()
    On Error GoTo LOKAL_ERROR
    
    Dim cZeile As String
    Dim lcount As Long
    Dim cSQL As String

    loeschNEW "PRINT_EXPRESSDAT", gdBase
    CreateTableT2 "PRINT_EXPRESSDAT", gdBase
    
    If List13.ListCount > 0 Then
    For lcount = 0 To List13.ListCount - 1
    
        cZeile = List13.list(lcount)
        
        cSQL = "Insert into PRINT_EXPRESSDAT (Zeile) values ('" & cZeile & "')"
        gdBase.Execute cSQL, dbFailOnError
        
    Next lcount
    
    reportbildschirm "WKL029", "aWKL23h"
    
End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken_Auflistung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub check_ex()
    On Error GoTo LOKAL_ERROR

    If Dat_vorhanden("N") Then
        Option2(3).ForeColor = vbRed
    Else
        If ermAnz_Expressdateien > 0 Then
            Option2(3).ForeColor = vbRed
        Else
            Option2(3).ForeColor = glS1
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "check_ex"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Dateienverarbeiten()
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cPfad       As String
    Dim dbwv        As Database
    Dim sSQL        As String
    Dim sName       As String
    Dim cdabapfad   As String
    Dim iStufe      As Integer
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    iStufe = 31
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "In"
    
    iStufe = 32
    
    File1.Path = cPfad
    File1.Pattern = "WV*.mdb"
    File1.Refresh
    
    cPfad = cPfad & "\"
    
    iStufe = 33
    
    
    If File1.ListCount = 0 Then Exit Sub
    
    For i = 0 To File1.ListCount - 1
    
        iStufe = i
        
        Set dbwv = Nothing
        
        iStufe = 77
        Set dbwv = OpenDatabase(cPfad & File1.list(i), False)
        
        'export in den Dabapfad
        sName = Left(File1.list(i), Len(File1.list(i)) - 4)
        
        iStufe = 88
        
        loeschNEW sName, gdBase
        
        iStufe = 44
        
        sSQL = "Select " & sName & ".* INTO " & sName & " IN '" & cdabapfad & "' from " & sName & " "
        dbwv.Execute sSQL, dbFailOnError
        
        iStufe = 55
            
        dbwv.Close
        
        Kill cPfad & File1.list(i)
    Next i
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Then
        iStufe = 91
        Resume Next
    ElseIf err.Number = 3343 Then
        iStufe = 92
        Kill cPfad & File1.list(i)
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Dateienverarbeiten"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten. " & iStufe
        
        Fehlermeldung1
    End If
End Sub
Private Sub DateienverarbeitenX()
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cPfad       As String
    Dim dbwv        As Database
    Dim sSQL        As String
    Dim sName       As String
    Dim cdabapfad   As String
    Dim iStufe      As Integer
    Dim cDatum      As String
    Dim ctmp        As String
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    iStufe = 31
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "In"
    
    iStufe = 32
    
    File2.Path = cPfad
    File2.Pattern = "N*.mdb"
    File2.Refresh
    
    cPfad = cPfad & "\"
    
    iStufe = 33
    
    
    If File2.ListCount = 0 Then Exit Sub
    
    For i = 0 To File2.ListCount - 1
    
        iStufe = i
        
        Set dbwv = Nothing
        sName = Left(File2.list(i), Len(File2.list(i)) - 4)
        
        iStufe = 77
        Set dbwv = OpenDatabase(cPfad & File2.list(i), False)
        
        'export in den Dabapfad
        
        cDatum = FileDateTime(cPfad & sName & ".mdb")
        
        sSQL = "Delete from ProtoEin where Datname = '" & sName & "'"
        gdBase.Execute sSQL, dbFailOnError
        sSQL = "Insert into ProtoEin (Datname,Datum) values ('" & sName & "','" & cDatum & "') "
        gdBase.Execute sSQL, dbFailOnError
        
        
        iStufe = 88
        
        loeschNEW sName, gdBase
        
        iStufe = 44
        
        sSQL = "Select " & sName & ".* INTO " & sName & " IN '" & cdabapfad & "' from " & sName & " "
        dbwv.Execute sSQL, dbFailOnError
        
        iStufe = 55
            
        dbwv.Close
        
        Kill cPfad & File2.list(i)
    Next i

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Then
        iStufe = 91
        Resume Next
    ElseIf err.Number = 3343 Then
    
        ctmp = "Die Datei: " & sName & " " & vbCrLf & vbCrLf
        ctmp = ctmp & "aus der Filiale: " & Val(Mid(sName, 4, 2)) & vbCrLf & vbCrLf
        ctmp = ctmp & "ist beschädigt." & vbCrLf
        ctmp = ctmp & "Sie müssen die Datei aus der Filiale: " & Val(Mid(sName, 4, 2)) & " nochmals anfordern!" & vbCrLf
        ctmp = ctmp & "___________________________________________________________" & vbCrLf & vbCrLf
        ctmp = ctmp & "Was ist in der Filiale " & Val(Mid(sName, 4, 2)) & " zu tun?" & vbCrLf
        ctmp = ctmp & "Winkiss: Im Filialtausch an der Kasse über die Schaltfläche 'nochmals Senden' diese Datei: (" & sName & ") erneut anfordern." & vbCrLf & vbCrLf
        ctmp = ctmp & "Zentrale: Im Menü über 'Filialen/Sicherungsdateien kopieren' Hier wählen Sie 'Expressverteilungen für die Filialen' und fordern diese Datei: (" & sName & ") erneut an."
        
        MsgBox ctmp, vbInformation + vbOKOnly, "Winkiss Hinweis:"
        
        Kill cPfad & sName & ".mdb"
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DateienverarbeitenX"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten. " & iStufe
        
        Fehlermeldung1
    End If
End Sub
Private Function ermAnz_Expressdateien() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    
    ermAnz_Expressdateien = 0
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "In"
    
    File2.Path = cPfad
    File2.Pattern = "N*.mdb"
    File2.Refresh
    
    ermAnz_Expressdateien = File2.ListCount

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAnz_Expressdateien"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Function

Private Sub Dateienladen()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    Screen.MousePointer = 11
    
    If gsZinPfad = "" Then
        gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    End If
    
    If Not Right(gsZinPfad, 3) = "mdb" Then
        gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    End If
    
    cPfad = gsZinPfad
    
    Set dbwv = Nothing
    Set dbwv = OpenDatabase(cPfad, False)
    
    NewListeFuellAnfangsbuch "WV", frmWKL23.List9, dbwv
    
    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 68 Or err.Number = 3043 Then
        List9.Clear
        anzeigeNew "rot", "Das Öffnen von der Diskette ist gescheitert.", Label5
        
        Screen.MousePointer = 0
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Dateienladen"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub NewListeFuellAnfangsbuchX(anfangsbuch As String, list As Object, daba As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim cDatum      As String
    Dim lcount      As Long
    Dim sdat        As String
    
    Dim rsrs        As Recordset
    Dim lLief       As Long
    Dim LenAnfang   As Integer
    Dim cQuellfil   As String
    
    LenAnfang = Len(anfangsbuch)
    
    List15.Clear
    List15.AddItem "Datei              Fil Datum"
    
    list.Clear
    daba.TableDefs.Refresh
    lAnzTable = daba.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = daba.TableDefs(lcount).name
        If UCase(Left(name, LenAnfang)) = UCase(anfangsbuch) Then
            cdatei = UCase$(name)
            If IsNumeric(Right(cdatei, Len(cdatei) - 1)) Then
                sdat = cdatei
            
                If datumvergleichen("Warenver", sdat) Then
                    loeschNEW sdat, daba
                Else
                    cQuellfil = Mid(sdat, 4, 2)
                    cDatum = datum_aus_Proto(sdat)
                    cdatei = cdatei & Space$(18 - Len(cdatei)) & " " & cQuellfil & "  " & cDatum
                    list.AddItem cdatei
                End If
            End If
         End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NewListeFuellAnfangsbuchX"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Dat_vorhanden(anfangsbuch As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim lcount      As Long
    Dim sdat        As String
    
    Dim rsrs        As Recordset
    Dim lLief       As Long
    Dim LenAnfang   As Integer
    
    Dat_vorhanden = False
    
    LenAnfang = Len(anfangsbuch)
    
    gdBase.TableDefs.Refresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If UCase(Left(name, LenAnfang)) = UCase(anfangsbuch) Then
            cdatei = UCase$(name)
            If IsNumeric(Right(cdatei, Len(cdatei) - 1)) Then
            
                Dat_vorhanden = True
                Exit For
                
            End If
         End If
    Next lcount
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Dat_vorhanden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub DateienladenEX()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    Screen.MousePointer = 11
    
    If gsZinPfad = "" Then
        gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    End If
    
    If Not Right(gsZinPfad, 3) = "mdb" Then
        gsZinPfad = gcDBPfad & "\Kissdata.mdb"
    End If
    
    cPfad = gsZinPfad
    
    Set dbwv = Nothing
    Set dbwv = OpenDatabase(cPfad, False)
    
    NewListeFuellAnfangsbuchX "N", frmWKL23.List13, dbwv

    Screen.MousePointer = 0
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 68 Or err.Number = 3043 Then
        List9.Clear
        anzeigeNew "rot", "Das Öffnen von der Diskette ist gescheitert.", Label5
        
        Screen.MousePointer = 0
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DateienladenEX"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub löschen()
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei      As String
    Dim lRet        As Long
    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lcount      As Long
    Dim ctmp        As String
    Dim cDatum      As String
    
    If List9.ListIndex < 0 Then
        anzeigeNew "rot", "Bitte einen Eintrag auswählen!", Label5
        List9.SetFocus
    Else
        cdatei = List9.list(List9.ListIndex)
        cdatei = UCase$(cdatei)
        cdatei = Left(cdatei, 8)
        cdatei = Trim$(cdatei)
        
        ctmp = "Wollen Sie die markierte Datei wirklich löschen?"
        
        dlgAbfrage.BCaptioneins = "Löschen"
        dlgAbfrage.BCaptionzwei = "Abbrechen"
        dlgAbfrage.Überschrift = "Winkiss Frage:"
        dlgAbfrage.Beschriftung = ctmp
        dlgAbfrage.Show vbModal
        
        If dlgAbfrage.Back = 1 Then
            Screen.MousePointer = 11
            loeschNEW cdatei, dbwv
            
            dateiloeschen
            schreibeWVEingangProtokoll cdatei & " gelöscht"
            
            NewListeFuellAnfangsbuch "WV", frmWKL23.List9, dbwv
            
            Screen.MousePointer = 0
        Else
            
        End If
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "löschen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub löschenX()
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei      As String
    Dim lRet        As Long
    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lcount      As Long
    Dim ctmp        As String
    Dim cDatum      As String
    Dim cSQL        As String
    
    
    If List13.ListIndex < 0 Then
        anzeigeNew "rot", "Bitte einen Eintrag auswählen!", Label5
        List13.SetFocus
    Else
        cdatei = List13.list(List13.ListIndex)
        cdatei = UCase$(cdatei)
        cdatei = Left(cdatei, 18)
        cdatei = Trim$(cdatei)
        
        ctmp = "Wollen Sie die markierte Datei wirklich löschen?"
        
        dlgAbfrage.BCaptioneins = "Löschen"
        dlgAbfrage.BCaptionzwei = "Abbrechen"
        dlgAbfrage.Überschrift = "Winkiss Frage:"
        dlgAbfrage.Beschriftung = ctmp
        dlgAbfrage.Show vbModal 'vbModal wieder reingenommem 22.05.15
        
        If dlgAbfrage.Back = 1 Then
            Screen.MousePointer = 11
            loeschNEW cdatei, dbwv
            
'            dateiloeschen
            schreibeWVExpressProtokoll cdatei & " gelöscht"
            
            cSQL = "Delete from ProtoEin where Datname = '" & cdatei & "'"
            gdBase.Execute cSQL, dbFailOnError
            
            NewListeFuellAnfangsbuchX "N", frmWKL23.List13, dbwv
            
            Screen.MousePointer = 0
        Else
            
        End If
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "löschenX"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub auswählen()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cdatei As String
    Dim cQuelle As String
    Dim cZiel As String
    Dim cSQL As String
    Dim lfail As Long
    Dim iRet As Integer
    Dim lRet As Long
    Dim cdabapfad As String
    
    Screen.MousePointer = 11
    cPfad = gsZinPfad
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    Set dbwv = OpenDatabase(cPfad, False)
    
    If List9.ListIndex < 0 Then
        anzeigeNew "rot", "Bitte einen Eintrag auswählen!", Label5
        
        List9.SetFocus
        Screen.MousePointer = 0
    Else
        cdatei = List9.list(List9.ListIndex)
        cdatei = UCase$(cdatei)
        cdatei = Left(cdatei, 8)
        cdatei = Trim$(cdatei)
        gcdatei = cdatei
        
        lbl6(6).Caption = gcdatei
        lbl6(6).Refresh
        'Import
        loeschNEW "filbeste", gdBase
        
        cSQL = "Select " & cdatei & ".* INTO " & cdatei & " IN '" & cdabapfad & "' from " & cdatei & " "
        dbwv.Execute cSQL, dbFailOnError

        cSQL = " Select * into filbeste from " & cdatei
        gdBase.Execute cSQL, dbFailOnError
        
        Screen.MousePointer = 0
        
        schreibeWVEingangProtokoll gcdatei & " ausgewählt"
        
        Speichernfilbeste gcdatei
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "auswählen"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub auswählenX()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cdatei As String
    Dim cQuelle As String
    Dim cZiel As String
    Dim cSQL As String
    Dim lfail As Long
    Dim iRet As Integer
    Dim lRet As Long
    Dim cdabapfad As String
    
    Screen.MousePointer = 11
    cPfad = gsZinPfad
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    Set dbwv = OpenDatabase(cPfad, False)
    
    If List13.ListIndex < 0 Then
        anzeigeNew "rot", "Bitte einen Eintrag auswählen!", Label5
        
        List13.SetFocus
        Screen.MousePointer = 0
    Else
        cdatei = List13.list(List13.ListIndex)
        cdatei = UCase$(cdatei)
        cdatei = Left(cdatei, 18)
        cdatei = Trim$(cdatei)
        gcdatei = cdatei
        
        Dim sZugriffsrechner As String
        sZugriffsrechner = Ist_Datei_im_Zugriff(gcdatei)
        If sZugriffsrechner <> "" Then
            anzeigeNew "rot", "Die Datei(" & gcdatei & ") ist im Zugriff(" & sZugriffsrechner & ")!", Label5
            Exit Sub
        
        End If
        InZugriffsetzen gcdatei
        
        lbl6(6).Caption = gcdatei
        lbl6(6).Refresh
        'Import
        loeschNEW "filbeste", gdBase
        
        cSQL = "Select " & cdatei & ".* INTO " & cdatei & " IN '" & cdabapfad & "' from " & cdatei & " "
        dbwv.Execute cSQL, dbFailOnError

        cSQL = " Select * into filbeste from " & cdatei
        gdBase.Execute cSQL, dbFailOnError
        
        Screen.MousePointer = 0
        
        schreibeWVExpressProtokoll gcdatei & " ausgewählt"
        
        SpeichernfilbesteX gcdatei
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "auswählenX"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub InZugriffsetzen(cdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    cdat = Trim(cdat)
    
    sSQL = "Delete from ZugriffDat where Dat = '" & cdat & "'"
    sSQL = sSQL & " and rechner = '" & srechnertab & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ZugriffDat (Dat,Rechner) values ('" & cdat & "','" & srechnertab & "')"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InZugriffsetzen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Function Ist_Datei_im_Zugriff(cdat As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    Ist_Datei_im_Zugriff = ""
    
    cdat = Trim(cdat)
    
    sSQL = "Select Rechner from ZugriffDat where Dat = '" & cdat & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Rechner) Then
            Ist_Datei_im_Zugriff = rsrs!Rechner
        Else
            Ist_Datei_im_Zugriff = "Workstation"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Function
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Ist_Datei_im_Zugriff"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
   
End Function
Private Sub auswählenXmEAN()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cdatei As String
    Dim cQuelle As String
    Dim cZiel As String
    Dim cSQL As String
    Dim lfail As Long
    Dim iRet As Integer
    Dim lRet As Long
    Dim cdabapfad As String
    Dim bFound As Boolean
    Dim lcount As Long
    bFound = False
    
    Screen.MousePointer = 11
    cPfad = gsZinPfad
    
    cdabapfad = gcDBPfad
    If Right(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "kissdata.mdb"
    
    Set dbwv = OpenDatabase(cPfad, False)
    
    bFound = False
    
    For lcount = 0 To List13.ListCount - 1
        cdatei = List13.list(lcount)
        cdatei = UCase$(cdatei)
        cdatei = Left(cdatei, 18)
        cdatei = Trim$(cdatei)
        gcdatei = cdatei
        
        
        
        
        
        
        cdatei = Right(cdatei, Len(cdatei) - 5)
        If cdatei = Trim(Text3.Text) Then
            bFound = True
            Exit For
        End If
        
    Next lcount
    
    If bFound Then
    
        Dim sZugriffsrechner As String
        sZugriffsrechner = Ist_Datei_im_Zugriff(gcdatei)
        If sZugriffsrechner <> "" Then
            anzeigeNew "rot", "Die Datei(" & gcdatei & ") ist im Zugriff(" & sZugriffsrechner & ")!", Label5
            Exit Sub
        
        End If
        InZugriffsetzen gcdatei
        
        lbl6(6).Caption = gcdatei
        lbl6(6).Refresh
        'Import
        loeschNEW "filbeste", gdBase
        
        cSQL = "Select " & gcdatei & ".* INTO " & gcdatei & " IN '" & cdabapfad & "' from " & gcdatei & " "
        dbwv.Execute cSQL, dbFailOnError

        cSQL = " Select * into filbeste from " & gcdatei
        gdBase.Execute cSQL, dbFailOnError
        
'        gcDateidatum = List13.list(lcount)
'        gcDateidatum = UCase$(gcDateidatum)
'        gcDateidatum = Right(gcDateidatum, 17)
        Screen.MousePointer = 0
        
        schreibeWVExpressProtokoll gcdatei & " ausgewählt"
        
        SpeichernfilbesteX gcdatei
    Else
    
        Unload Me
        frmWKL23.Show
        
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "auswählenXmEAN"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub MdeVerarbeitung1()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsFilB      As Recordset
    Dim rsArt       As Recordset
    Dim seekEAN     As String
    
    Screen.MousePointer = 11
    
    Check2.Visible = False
    Command6(12).Visible = False
    
    loeschNEW "Umvteil", gdBase
    CreateTable "Umvteil", gdBase
    
    Set rsFilB = gdBase.OpenRecordset("Umvteil")
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
            If Not IsNull(rsMDE!eancode) Then
                seekEAN = Trim(rsMDE!eancode)
                seekEAN = checkean(seekEAN)
                
                If Len(seekEAN) = 11 Then
                    seekEAN = "0" & seekEAN
            
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                ElseIf Len(seekEAN) = 8 Then
                    If Left(seekEAN, 1) = "2" Then
                        seekEAN = Mid$(seekEAN, 2, 6)
                        sSQL = "select * from artikel where artnr = " & seekEAN
                    Else
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    End If
                
                Else
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                End If
                
               
                
                Set rsArt = gdBase.OpenRecordset(sSQL)
                
                If Not rsArt.EOF Then 'hier die bekannten
                    rsFilB.AddNew
                    
                    rsFilB!artnr = rsArt!artnr
                    rsFilB!BEZEICH = rsArt!BEZEICH
                    rsFilB!linr = rsArt!linr
                    rsFilB!LIBESNR = rsArt!LIBESNR
                    rsFilB!LPZ = rsArt!LPZ
                    rsFilB!KVKPR1 = rsArt!KVKPR1
                    rsFilB!BESTVOR = rsMDE!Menge
                    rsFilB!FILIALE = CByte(gcFilNr)
                    rsFilB!Status = "vorhanden"

                    rsFilB.Update
                Else 'hier die unbekannten
                
                    rsFilB.AddNew
                    rsFilB!BEZEICH = seekEAN
                    rsFilB!BESTVOR = rsMDE!Menge
                    rsFilB!Status = "nicht vorhanden"
                    rsFilB!FILIALE = CByte(gcFilNr)
                    rsFilB.Update
                    
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsMDE.MoveNext
        Loop
    
    End If
    
    rsMDE.Close: Set rsMDE = Nothing
    rsFilB.Close: Set rsFilB = Nothing
    
    anzeigeMDE
    
    anzeigeNew "normal", "Wollen Sie die eingelesenen Artikel jetzt in den Bestand übernehmen?", Label5

    Command6(12).Visible = True 'Einlese Button aktiv
    Check2.Visible = True

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "MdeVerarbeitung1"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command7_Click()
    On Error GoTo LOKAL_ERROR
    
    If Label2(2).Caption <> "0" Then
        Screen.MousePointer = 11
        gsARTNR = Label2(2).Caption
        If gsARTNR <> "" Then
            frmWKL10.Show 1
            Me.Refresh
        End If
        gsARTNR = ""
    
        Command1_Click
    Else
        MsgBox "Bitte einen Artikel festlegen!", vbInformation, "Winkiss Hinweis:"
        Text1(0).SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    gcArtNrFiliale = Trim(Label2(2).Caption)
    frmWKLae.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_Click()
On Error GoTo LOKAL_ERROR

Dim lLoeschen As Long

'Grund.cfg
Dim cpfaddb As String
Dim sLieferschein As String

cpfaddb = gcDBPfad
If Right$(cpfaddb, 1) <> "\" Then
    cpfaddb = cpfaddb & "\"
End If

If tableSuchenDBKombi("ZuAusUV", 2) Then

    If FileExists(cpfaddb & "Grund.cfg") Then
    
        sLieferschein = Trim(Text1(7).Text)
        
        If sLieferschein = "" Then
            MsgBox "Bitte einen Lieferschein-Nr angeben!", vbInformation, "Winkiss Hinweis:"
            Text1(7).SetFocus
            Exit Sub
        End If
        
        Schreibe_Lieferschein_Navision sLieferschein
    End If
    'Ende Grund.cfg
    
End If




If tableSuchenDBKombi("ZuAusUV", 2) Then
    lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
    
    If FileExists(App.Path & "\aWKL23e.rpt") Then
        reportbildschirmApp "dWKL15", "aWKL23e"
    Else
        reportbildschirmApp "dWKL15", "aWKL23b"
    End If
    
Else
    MsgBox "Es sind keine Druckdaten vorhanden.", vbInformation, "Winkiss Hinweis:"
End If
    
If lLoeschen = vbYes Then
    loeschapp "ZuAusUV"
End If

SchreibeListe

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Schreibe_Lieferschein_Navision(sLieferschein As String)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim cSatz       As String
    Dim rsrs        As Recordset
    Dim cSQL        As String
    Dim sTime       As String
    Dim sDate       As String
    
    Dim cPfad2      As String
    
    cPfad2 = gcDBPfad
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    sTime = Format$(TimeValue(Now), "HHMMSS")
    sDate = Format$(DateValue(Now), "DDMMYYYY")
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "STAT\"
    
    Kill cPfad & "WV_" & gcFilNr & "_" & sLieferschein & "_" & sDate & "_" & sTime & ".csv"
    
    iFileNr = FreeFile
    Open cPfad & "WV_" & gcFilNr & "_" & sLieferschein & "_" & sDate & "_" & sTime & ".csv" For Binary As #iFileNr
    
    loeschNEW "tart", gdApp
    cSQL = "Select '" & gcFilNr & "' as Filiale, '" & sLieferschein & "' as LS, 'Warenumverteilung' as Art "
    cSQL = cSQL & " ,artnr,BEZEICH,EAN,'' as EAN2 , '' as EAN3, BEWEGUNG into tArt from ZuAusUV order by LFNR  "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "update tart as A inner join [;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].artikel as B on A.artnr = B.Artnr "
    cSQL = cSQL & " set a.EAN = b.EAN  "
    cSQL = cSQL & " , a.EAN2 = b.EAN2  "
    cSQL = cSQL & " , a.EAN3 = b.EAN3  "
    gdApp.Execute cSQL, dbFailOnError
    
    
    cSQL = " select * from tart "
    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
            
                cSatz = ""
                cSatz = cSatz & rsrs!FILIALE & vbTab
                cSatz = cSatz & rsrs!LS & vbTab
                cSatz = cSatz & rsrs!art & vbTab
                cSatz = cSatz & rsrs!artnr & vbTab
                cSatz = cSatz & rsrs!BEZEICH & vbTab
                
                cSatz = cSatz & rsrs!BEWEGUNG & vbTab
                
                cSatz = cSatz & rsrs!EAN & vbTab
                cSatz = cSatz & rsrs!EAN2 & vbTab
                cSatz = cSatz & rsrs!EAN3
                cSatz = cSatz & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr

    Dim bmerke  As Boolean
    bmerke = gbFTPautomatic

    If gbFtpYes Then

        gbFTPautomatic = True
        giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
        frmWKL38.Show 1
        gbFTPautomatic = bmerke

    End If

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Schreibe_Lieferschein_Navision"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    PositionierenWKL23
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    lbl6(6).Caption = ""
    lbl6(6).Refresh
    
    
    bfoundauto = False
    fromMde = False
    bscanner = False
    
    If NewTableSuchenDBKombi("ZuAusUV", gdApp) Then
        If Datendrin("ZuAusUV", gdApp) Then
            Command9.BackColor = vbRed
        End If
    End If
    
    If NewTableSuchenDBKombi("PROTOEIN", gdBase) = False Then
        CreateTableT2 "PROTOEIN", gdBase
    End If

    
    Option2(Leselast23Einstellung).value = True
    Option2(2).Caption = Option2(2).Caption & " (" & gsMDEGERAET & ")"
    
    check_ex
    
'    Option2(3).Caption = Option2(3).Caption & " (" & ermAnz_Expressdateien & ")"
    
'    füllefil cboFil
'    füllefil cbofil3
    Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungManuell()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    Frame3.Visible = True
    Text1(1).Text = gsWeEinzMe
    Text1(0).SetFocus
    Option1(1).value = True
    LeereDialogWKL15
    gF2Prompt.lLastPos = -1
    List14.AddItem "Datum     Zeit ArtNr  Artikelbezeichnung                      Zu/Abgang  KVK"
    SchreibeListe

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungManuell"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungDatei()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    gbDrueck5 = True
    Frame1.Visible = False
    Frame6.Visible = True
    txtZinPfad.Text = gsZinPfad
    
    Dateienverarbeiten
    Dateienladen
    
    If gbFtpYes Then
        Command6(13).Enabled = True
    End If
    
    

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungDatei"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungExpress()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    gbDrueck5 = True
    
'    Text2.Text = gsZinPfad
    
    DateienverarbeitenX
    DateienladenEX
    
    If gbFtpYes Then
        If gbWVNOT = False Then
            Command6(16).Enabled = True
        End If
    End If
    
    Frame1.Visible = False
    
    If Text3.Text <> "" Then
        If IsNumeric(Text3.Text) Then
        
            If Len(Trim(Text3.Text)) = 8 And Left(Trim(Text3.Text), 1) = "2" Then
                Text3.Text = Trim(Text3.Text)
                Text3.Text = Val(Mid(Text3.Text, 2, 6))
            End If
            
            auswählenXmEAN
            Exit Sub
        End If
    End If
    
    Frame9.Visible = True

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
'    If err.Number = 53 Then
'
'        Resume Next
'    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "vorbereitungExpress"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'    End If
End Sub
Private Sub vorbereitungMDE()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    Dim ctmp As String

    gbDrueck5 = True
    Frame5.Visible = True
    If UCase(gsMDEGERAET) = "SCANPAL" Then
        ctmp = "Gerät richtig einstellen! "
        ctmp = ctmp & "Am Scanpal 2 auf 'Daten senden' navigieren. "
        ctmp = ctmp & "Dann mit der Enter - Taste auf dem Scanpal 2 bestätigen. "
        ctmp = ctmp & "Wenn dann auf dem Display des Scanpal 2 'Verbindung....' steht, "
        ctmp = ctmp & "können Sie auf den hier unten aufgeführten Button 'Einlesen' klicken."
        Command6(9).Enabled = True
    ElseIf UCase(gsMDEGERAET) = "FORCOM" Then
        ctmp = "Das Formula in die Station stecken - dann im Menü 'Übertragen' anwählen"
        ctmp = ctmp & " dann Enter auf dem Formula drücken und danach hier im Programm auf 'Einlesen' klicken."
       
    End If
    
    lbl6(5).Caption = ctmp
    lbl6(5).Refresh

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungMDE"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub einlesen(sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dTranspack  As Double
    
    Screen.MousePointer = 11
    
    sSQL = "Select * from UMVERTEIL_EXPRESS where Status = 'vorhanden' "
    Set rs = gdBase.OpenRecordset(sSQL)
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Verarbeitung vorhanden", Label5
        
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    
    pbr.Max = 50
    pbr.Visible = True
    
    lCounter = 0
    rs.MoveFirst
    If Not rs.EOF Then
        anzeigeNew "normal", "Die Warenlieferung wird jetzt eingelesen...", Label5
        Do While Not rs.EOF
            If lCounter = 50 Then
                lCounter = 0
            End If
            lCounter = lCounter + 1
            pbr.value = lCounter
            
            If Not IsNull(rs!artnr) Then
                 Text1(0).Text = rs!artnr
            End If
            
            SucheArtikelWKL15
            
            If Not IsNull(rs!BESTVOR) Then
                 Text1(1).Text = rs!BESTVOR
            End If
            
            If Option2(3).value = True Then
                dTranspack = CDbl(Right(gcdatei, Len(gcdatei) - 5))
                
                Dim cFilvon As String
                cFilvon = Mid(gcdatei, 4, 2)
                
                cFilvon = Val(cFilvon)
                
                If DatendrinSQL("select * from " & sdat & " where AENART = 'Filialtausch WK'", gdBase) Then
                
                    ABINFEEDBF rs!artnr, rs!BESTVOR
                    
                    'nur für kisslive
                    ABINFEEDB rs!artnr, rs!BESTVOR, dTranspack, cFilvon
                    
                Else
                
                    ABINFEEDB rs!artnr, rs!BESTVOR, dTranspack, cFilvon
                End If

            ElseIf Option2(1).value = True Then

            End If
            
            SchreibeDatenWKL15
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    pbr.Visible = False
    Screen.MousePointer = 0
    anzeigeNew "normal", "Die Aktualisierung wurde erfolgreich durchgeführt.", Label5
    
    Check1.Visible = False
    Command6(1).Visible = False
    
    If Option2(3).value = True Then
        
        schreibeWVExpressProtokoll sdat & " eingelesen: " & iPos & " versch. Artikel(" & iSum & " Gesamtmenge)"
        If iNeg > 0 Then
            schreibeWVExpressProtokoll sdat & " NICHT eingelesen: " & iNeg & " Artikel (Artikel waren in der Datenbank nicht vorhanden)"
        End If
        loeschNEW sdat, dbwv
        schreibeWVExpressProtokoll sdat & " gelöscht"
        
        sSQL = "Delete from ProtoEin where Datname = '" & sdat & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        NewListeFuellAnfangsbuchX "N", frmWKL23.List13, dbwv
        
    ElseIf Option2(1).value = True Then
        schreibeWVEingangProtokoll sdat & " eingelesen: " & lCounter & " Artikel"
    End If
    
    datumschreiben "Warenver", sdat
    
'    loeschNEW "KO23", gdBase
'    CreateTable "KO23", gdBase

    sSQL = "Update UMVERTEIL_EXPRESS set Dat = '" & sdat & "' "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Insert into KO23 (Dat) values ('" & sdat & "')"
'    gdBase.Execute sSQL, dbFailOnError
    
    reportbildschirm "umv1", "aWKL23"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "einlesen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub einlesenausMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    
    Screen.MousePointer = 11
    
    sSQL = "Select * from umvteil where Status = 'vorhanden' "
    Set rs = gdBase.OpenRecordset(sSQL)
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Verarbeitung vorhanden", Label5
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    
    pbr.Max = 50
    pbr.Visible = True
    
    lCounter = 0
    rs.MoveFirst
    If Not rs.EOF Then
        anzeigeNew "normal", "Die Warenlieferung wird jetzt eingelesen...", Label5
        Do While Not rs.EOF
            If lCounter = 50 Then
                lCounter = 0
            End If
            lCounter = lCounter + 1
            pbr.value = lCounter
            
            If Not IsNull(rs!artnr) Then
                 Text1(0).Text = rs!artnr
            End If
            SucheArtikelWKL15
            If Not IsNull(rs!BESTVOR) Then
                 Text1(1).Text = rs!BESTVOR
            End If
            
            ABINFEEDBF rs!artnr, rs!BESTVOR
            SchreibeDatenWKL15
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    pbr.Visible = False
    Screen.MousePointer = 0
    anzeigeNew "normal", "Die Aktualisierung wurde erfolgreich durchgeführt.", Label5
    
    Check2.Visible = False
    Command6(12).Visible = False
    
    reportbildschirm "umv1", "aWKL23d"
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "einlesenausmde"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub anzeigeN()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim ckPr        As String
    Dim cMenge      As String
    Dim cLinr       As String
    Dim iZaehler    As Integer
    
    List6.Clear
    List5.Clear
    List6.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from UMVERTEIL_EXPRESS where Status = 'vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    iSum = 0
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            iSum = iSum + CInt(cMenge)
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List5.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        
        
        Label7(1).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(1).Refresh
        
        Label7(8).Caption = "Summe (Menge): " & iSum & " Artikel"
        Label7(8).Refresh
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    List8.Clear
    List7.Clear
    List8.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from UMVERTEIL_EXPRESS where Status = 'nicht vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List7.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        
        Command6(2).Visible = True
        
        Label7(3).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(3).Refresh
    Else
        Command6(2).Visible = False
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Frame4.Visible = True
    Frame6.Visible = False
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "anzeigen"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub anzeigenX()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim ckPr        As String
    Dim cMenge      As String
    Dim cLinr       As String
    Dim iZaehler    As Integer
    
    List6.Clear
    List5.Clear
    List6.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from UMVERTEIL_EXPRESS where Status = 'vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    iSum = 0
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            iSum = iSum + CInt(cMenge)
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List5.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        iPos = 0
        iPos = iZaehler
        Label7(1).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(1).Refresh
        
        Label7(8).Caption = "Summe (Menge): " & iSum & " Artikel"
        Label7(8).Refresh
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    List8.Clear
    List7.Clear
    List8.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from UMVERTEIL_EXPRESS where Status = 'nicht vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List7.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        
        Command6(2).Visible = True
        iNeg = 0
        iNeg = iZaehler
        Label7(3).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(3).Refresh
    Else
        Command6(2).Visible = False
        
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "anzeigenX"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub anzeigeMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim ckPr        As String
    Dim cMenge      As String
    Dim cLinr       As String
    Dim iZaehler    As Integer
    
    List12.Clear
    List11.Clear
    List12.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from umvteil where Status = 'vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List11.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Label7(5).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(5).Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    List4.Clear
    List3.Clear
    List4.AddItem "EANCODE       Menge Scanreihenfolge"
    
    sSQL = "Select * from umvteil where Status = 'nicht vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!lfnr), "0", rsrs!lfnr)
        
            cLBSatz = cBez & Space$(14 - Len(cBez))
            cLBSatz = cLBSatz & cMenge & Space$(6 - Len(cMenge)) & cLinr
            List3.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Command6(8).Visible = True
        Label7(4).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(4).Refresh
    Else
        Command6(8).Visible = False
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame8.Visible = True
    Frame5.Visible = False
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "anzeigeMDE"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub Speichernfilbeste(sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim rs          As Recordset
    Dim rsrs        As Recordset
    Dim Rsk         As Recordset
    Dim cArttemp    As String
    Dim sSQL        As String
    Dim cPfad       As String
    Dim ctmp        As String
    
    cPfad = gcDBPfad        'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    loeschNEW "vkpro", gdBase
    
    sSQL = "Select * into vkpro from filbeste where Filiale = " & gcFilNr
    gdBase.Execute sSQL, dbFailOnError
       
    
    loeschNEW "UMVERTEIL_EXPRESS", gdBase
    
    sSQL = "Create Table UMVERTEIL_EXPRESS"
    sSQL = sSQL & "(ARTNR Double"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LINR Double"
    sSQL = sSQL & ", LPZ Double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", BESTVOR Long"
    sSQL = sSQL & ", FILIALE BYTE"
    sSQL = sSQL & ", Status Text(30)"
    sSQL = sSQL & ", dat TEXT(20)"
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("UMVERTEIL_EXPRESS", dbOpenTable)
    Set Rsk = gdBase.OpenRecordset("vkpro", dbOpenTable)
    
    If Not Rsk.EOF Then
        Rsk.MoveFirst
        Do While Not Rsk.EOF
            rsrs.AddNew
            rsrs!artnr = Rsk!artnr
            rsrs!BEZEICH = Rsk!BEZEICH
            rsrs!linr = Rsk!linr
            rsrs!LPZ = Rsk!LPZ
            rsrs!LIBESNR = Rsk!LIBESNR
            rsrs!KVKPR1 = Rsk!KVKPR1
            rsrs!BESTVOR = Rsk!BESTVOR
            rsrs!FILIALE = Rsk!FILIALE
            
            sSQL = "Select * from artikel where artikel.artnr = " & Rsk!artnr
            Set rs = gdBase.OpenRecordset(sSQL)
            If rs.EOF Then
                rsrs!Status = "nicht vorhanden"
            Else
                rsrs!Status = "vorhanden"
            End If
            rs.Close: Set rs = Nothing
            
            rsrs!dat = ""
            
            rsrs.Update
            Rsk.MoveNext
        Loop
        
        anzeigeN
        
        rsrs.Close: Set rsrs = Nothing
        Rsk.Close
        
        If datumvergleichen("Warenver", sdat) Then
            anzeigeNew "rot", "Dieser Wareneingang wurde schon eingelesen!", Label5
            Exit Sub
        End If
        
        anzeigeNew "normal", "Wollen Sie die Warenlieferung jetzt einlesen?", Label5

        Command6(1).Visible = True
        Check1.Visible = True
    Else
        anzeigeNew "Rot", "Für Ihre Filiale liegen keine Artikeldaten vor.", Label5
        
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Speichernfilbeste"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub SpeichernfilbesteX(sdat As String)
    On Error GoTo LOKAL_ERROR
    
    Dim rs          As Recordset
    Dim rsrs        As Recordset
    Dim Rsk         As Recordset
    Dim cArttemp    As String
    Dim sSQL        As String
    Dim cPfad       As String
    Dim ctmp        As String
    Dim iRet        As Integer
    
    cPfad = gcDBPfad        'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    loeschNEW "vkpro", gdBase
    
    sSQL = "Select * into vkpro from filbeste  where Filiale = " & gcFilNr
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "UMVERTEIL_EXPRESS", gdBase
    
    sSQL = "Create Table UMVERTEIL_EXPRESS"
    sSQL = sSQL & "(ARTNR Double"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LINR Double"
    sSQL = sSQL & ", LPZ Double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", BESTVOR Long"
    sSQL = sSQL & ", FILIALE BYTE"
    sSQL = sSQL & ", Status Text(30)"
    sSQL = sSQL & ", dat TEXT(20)"
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("UMVERTEIL_EXPRESS", dbOpenTable)
    Set Rsk = gdBase.OpenRecordset("vkpro", dbOpenTable)
    
    If Not Rsk.EOF Then
        Rsk.MoveFirst
        Do While Not Rsk.EOF
            rsrs.AddNew
            rsrs!artnr = Rsk!artnr
            rsrs!BEZEICH = Rsk!BEZEICH
            rsrs!linr = Rsk!linr
            rsrs!LPZ = Rsk!LPZ
            rsrs!LIBESNR = Rsk!LIBESNR
            rsrs!KVKPR1 = Rsk!KVKPR1
            rsrs!BESTVOR = Rsk!BESTVOR
            rsrs!FILIALE = Rsk!FILIALE
            
            sSQL = "Select * from artikel where artikel.artnr = " & Rsk!artnr
            Set rs = gdBase.OpenRecordset(sSQL)
            If rs.EOF Then
                rsrs!Status = "nicht vorhanden"
            Else
                rsrs!Status = "vorhanden"
            End If
            rs.Close: Set rs = Nothing
            
            rsrs!dat = ""
            
            rsrs.Update
            Rsk.MoveNext
        Loop
        
        anzeigenX
        
        rsrs.Close: Set rsrs = Nothing
        Rsk.Close
        
        Frame4.Visible = True
        Frame9.Visible = False
        Me.Refresh
        
        Screen.MousePointer = 11
        
        If datumvergleichen("Warenver", sdat) Then
            anzeigeNew "rot", "Dieser Wareneingang wurde schon eingelesen!", Label5
            iRet = MsgBox("Dieser Wareneingang wurde schon eingelesen, möchten Sie ihn löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
            
                loeschNEW sdat, dbwv
                schreibeWVExpressProtokoll sdat & " gelöscht"
                
                sSQL = "Delete from ProtoEin where Datname = '" & sdat & "'"
                gdBase.Execute sSQL, dbFailOnError
                
                NewListeFuellAnfangsbuchX "N", frmWKL23.List13, dbwv
            
                
            End If
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Screen.MousePointer = 0
        anzeigeNew "normal", "Wollen Sie die Warenlieferung jetzt einlesen?", Label5

        Command6(1).Visible = True
        Check1.Visible = True
        
        
    Else
        anzeigeNew "Rot", "Für Ihre Filiale liegen keine Artikeldaten vor.", Label5
        
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SpeichernfilbesteX"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    dbwv.Close
    
    AlleZugriffeLöschen
    
    loeschNEW "KO23", gdBase
    loeschNEW "vkpro", gdBase
    loeschNEW "umvteil", gdBase
    loeschNEW "filbeste", gdBase
    loeschNEW "PRINT_EXPRESSDAT", gdBase
    LogtoEnd Me
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Or err.Number = 3420 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Unload"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub List2_dblClick()
    On Error GoTo LOKAL_ERROR

    Command4_Click 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = 13 Then
        Command4_Click 0
    End If
    
    If KeyCode = 27 Then
        Command4_Click 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List4_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command4_Click 2
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List4_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = 13 Then
        Command4_Click 2
    End If
    
    If KeyCode = 27 Then
        Command4_Click 3
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    speicherlast23Einstellung index
    
    If Option2(3).value = True Then
        Text3.Visible = True
        Text3.Text = ""
'        Text3.SetFocus
        
    Else
        Text3.Visible = False
    End If
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherlast23Einstellung(i As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschapp "UMVENX"
    CreateTable "UMVENX", gdApp
    
    sSQL = "Insert into UMVENX (Ind) values (" & i & ")"
    gdApp.Execute sSQL, dbFailOnError
    
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherlast23Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Leselast23Einstellung() As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    Leselast23Einstellung = 0
    
    If Not NewTableSuchenDBKombi("UMVENX", gdApp) Then
        CreateTable "UMVENX", gdApp
        
        sSQL = "Insert into UMVENX (Ind) values (0)"
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    
    Set rsrs = gdApp.OpenRecordset("UMVENX")
    If Not rsrs.EOF Then
        Leselast23Einstellung = rsrs!ind
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
     
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Leselast23Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Option2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command11_Click 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(index).BackColor = glSelBack1
    Label3.Caption = Format$(index, "##0")
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case index
        Case Is = 0
            'wegen Volltextsuche nicht mehr gültig
            cValid = "1234567890" & Chr$(8)
        Case Is = 1
            cValid = "1234567890+-" & Chr$(8)
        Case Is = 2
            cValid = "1234567890," & Chr$(8)
        Case Is = 3
            cValid = "1234567890," & Chr$(8)
        Case Is = 4
            cValid = "1234567890" & Chr$(8)
        Case Is = 5
            cValid = "1234567890" & Chr$(8)
        Case Is = 7
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!?"
    End Select
    
    If index <> 0 And index <> 6 Then
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    End If
    If index = 2 And cZeichen = "," Then
        If InStr(Text1(index).Text, ",") > 0 Then
            KeyAscii = 0
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    bscanner = False
    
    If KeyCode = vbKeyReturn Then
        If index = 0 Then
            If gbscanmodi Then
                bscanner = True
            Else
                bscanner = False
            End If
            Command1_Click
        End If
        If index >= 1 Then
            Command2_Click 15
        End If
    End If
    
    If KeyCode = vbKeyF4 Then
        If index = 0 Then
            ctmp = Trim$(Text1(4).Text)
            If ctmp = "" Then
                MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                Text1(4).SetFocus
                Exit Sub
            End If
            
            gF2Prompt.cFeld = "ARTNRPOS"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            Command1_Click
            ctmp = Trim$(Text1(0).Text)
            If ctmp = "" Then
                MsgBox "Bitte den Artikel eindeutig bestimmen (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            End If
            gF2Prompt.cWert2 = ctmp
        
        If gF2Prompt.cFeld <> "" Then
            
            frmWK00a.Show 1
        
            If gF2Prompt.cWahl <> "" Then
                Text1(index).Text = gF2Prompt.cWahl
                If index = 0 Then
                    Command1_Click
                End If
            End If
            
        End If
        
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False

        Select Case index
            Case Is = 0     'Artikel
                ctmp = Trim$(Text1(4).Text)
                If ctmp = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Text1(4).SetFocus
                    Exit Sub
                Else
                    gF2Prompt.cFeld = "ARTNRPOS"
                    gF2Prompt.cWert = ctmp
                End If
            
            Case Is = 4     'Lieferant
                gF2Prompt.cFeld = "LINR"
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            
            frmWK00a.Show 1
        
            If gF2Prompt.cWahl <> "" Then
                Text1(index).Text = gF2Prompt.cWahl
                If index = 0 Then
                    Command1_Click
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
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    
    If index = 4 Then
        ctmp = Text1(4).Text
        ctmp = Trim$(Str$(Val(ctmp)))
        
        cSQL = "Select * from LISRT where LINR = " & ctmp & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!LIEFBEZ) Then
                Label2(4).Caption = rsrs!LIEFBEZ
            Else
                Label2(4).Caption = ""
            End If
        Else
            Label2(4).Caption = ""
        End If
        rsrs.Close: Set rsrs = Nothing
        
        LeseLieferantenPreisWKL15
    End If
    
    Text1(index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus()
On Error GoTo LOKAL_ERROR

    Text3.BackColor = glSelBack1
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command11_Click 1
    End If
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text3_LostFocus()
On Error GoTo LOKAL_ERROR

    Text3.BackColor = vbWhite
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Umverteilung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtStatus_Change()
    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh
    
End Sub
