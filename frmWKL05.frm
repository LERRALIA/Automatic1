VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL05 
   Caption         =   "Exportoptionen"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frmWKL05.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10080
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame4 
      Caption         =   "Zeitpunkt festlegen"
      Height          =   5295
      Left            =   1200
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   7335
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   35
         Top             =   3120
         Width           =   2295
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
         Caption         =   "alle Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   5
         Left            =   3840
         TabIndex        =   34
         Top             =   2640
         Width           =   2295
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
         Caption         =   "Übernehmen vom Vortag"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   3840
         TabIndex        =   31
         Top             =   720
         Width           =   2295
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   4
         Left            =   7680
         TabIndex        =   29
         Top             =   3120
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131923970
         CurrentDate     =   38457.8333333333
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   58
         Top             =   3600
         Width           =   2295
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
         Caption         =   "Eintrag Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
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
         Left            =   3840
         TabIndex        =   36
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Auswertungstag"
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
         Index           =   9
         Left            =   3840
         TabIndex        =   32
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Auswertungstag"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Einstellungen"
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CheckBox Check4 
         Caption         =   "Normalexport"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   2295
      End
      Begin VB.Frame Frame5 
         Caption         =   "Normalexport"
         Height          =   4215
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   3255
         Begin VB.CheckBox Check9 
            Caption         =   "mit Spaltenüberschriften"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3840
            Width           =   2895
         End
         Begin VB.CheckBox Check8 
            Caption         =   "+ Shop Preis"
            Height          =   255
            Left            =   1800
            TabIndex        =   56
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox Check7 
            Caption         =   "nur Shopartikel"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1575
         End
         Begin VB.Frame Frame6 
            Caption         =   "Feldtrenner"
            Height          =   1215
            Left            =   120
            TabIndex        =   50
            Top             =   2520
            Width           =   2895
            Begin VB.OptionButton Option2 
               Caption         =   "Komma"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   53
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Tab"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Semikolon"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   51
               Top             =   840
               Width           =   1455
            End
         End
         Begin VB.OptionButton Option1 
            Caption         =   "csv"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   2160
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "txt"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   47
            Top             =   2160
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.CheckBox Check6 
            Caption         =   "+ Bezeichnung"
            Height          =   255
            Left            =   1320
            TabIndex        =   46
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "+ EAN"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "nur Artikel mit Bestand > 0"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Dateiendung"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   49
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "VEDES"
         Height          =   3495
         Left            =   3720
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   40
            Top             =   600
            Width           =   2175
         End
         Begin sevCommand3.Command Command36 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
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
            Caption         =   "FTP"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Kundenkennung"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "VEDES DSL"
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         ToolTipText     =   "VEDES Digitale Shopping Lösung"
         Top             =   240
         Width           =   2295
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   7
         Left            =   7680
         TabIndex        =   37
         Top             =   240
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame3 
         Caption         =   "Zeitpunkte"
         Height          =   3615
         Left            =   7440
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   6
            Left            =   1440
            TabIndex        =   19
            Top             =   3120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   5
            Left            =   1440
            TabIndex        =   18
            Top             =   2640
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   4
            Left            =   1440
            TabIndex        =   17
            Top             =   2160
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   3
            Left            =   1440
            TabIndex        =   16
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   2
            Left            =   1440
            TabIndex        =   15
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   1
            Left            =   1440
            TabIndex        =   14
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   0
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   370
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.CheckBox Check1 
            Caption         =   "sonntags"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "samstags"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "freitags"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "donnerstags"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "mittwochs"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "dienstags"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "montags"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Dateiendung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Test"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label lblAnzeige 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   7455
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelbestandsexport"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmWKL05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Check1(Index).Value = vbChecked Then
        Command2(Index).Enabled = True
    Else
        Command2(Index).Enabled = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR

    If Check3.Value = vbChecked Then
        Frame2.Visible = True
    Else
        Frame2.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check4_Click()
On Error GoTo LOKAL_ERROR

    If Check4.Value = vbChecked Then
        Frame5.Visible = True
    Else
        Frame5.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check7_Click()
On Error GoTo LOKAL_ERROR

    If Check7.Value = vbChecked Then
        Check8.Visible = True
    Else
        Check8.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check7_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim sVTAG       As String
Dim cPfad       As String
Dim bmerke As Boolean

Select Case Index
    Case 0
        anzeigeNew "normal", "", lblAnzeige
        
        Command1_Click 7
        
        lese_Ex_Steu
        
        If gbEXNOR Then
            If Export_Artikelbestände Then
                If gbFtpYes Then
                    giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
                    frmWKL38.Show 1
                Else
                    cPfad = gcDBPfad
                    If Right$(cPfad, 1) <> "\" Then
                        cPfad = cPfad & "\"
                    End If
                    cPfad = cPfad & "STAT\"
                
                    gsAnzeigeText = "Die Datei 'Best.txt' ist unter: " & cPfad & " erstellt. Bitte übertragen Sie diese."
                    frmWK21l.Show 1
                End If
            End If
        End If
        
        If gbBL Then 'Jetzt Vedes komplett Export
            If Export_Artikelbestände_Komplett_Vedes Then
                If gbFtpYes Then
                    gbFTPautomatic = True
                    giKissFtpMode = 25 'FTPMODE= 25 , VEDESDSL - Ordner leeren abschicken
                    frmWKL38.Show 1
                    gbFTPautomatic = bmerke
                Else

                    cPfad = gcDBPfad
                    If Right$(cPfad, 1) <> "\" Then
                        cPfad = cPfad & "\"
                    End If
                    cPfad = cPfad & "VEDESDSL\"

                    gsAnzeigeText = "Die Datei 'retailer_*.txt' ist unter: " & cPfad & " erstellt. Bitte übertragen Sie diese."
                    frmWK21l.Show 1
                End If

            End If
        End If
        
        
        
        
        anzeigeNew "normal", "Artikelbestände erfolgreich exportiert", lblAnzeige
    Case 1
        Unload frmWKL05
    Case 2
        If List1.ListIndex < 0 Then
'                anzeigeNew "rot", "Bitte einen Eintrag in der Liste markieren!", lbl1
        Else
            LoescheEinzeln List1.list(List1.ListIndex)
            Lese_AEA_Zeiten List1, Label1(9).Caption
        End If
        
        
    Case 3
        speicher_AEA_ZEITEN
        Lese_AEA_Zeiten List1, Label1(9).Caption
    Case 4
        Frame4.Visible = False
        ZeigeZeitenSummen
        Label1(10).Caption = ""
    Case 5
        Loeschetag
        Select Case Label1(9).Caption
            Case "Montag"
                sVTAG = "Sonntag"
            Case "Dienstag"
              sVTAG = "Montag"
            Case "Mittwoch"
              sVTAG = "Dienstag"
            Case "Donnerstag"
              sVTAG = "Mittwoch"
            Case "Freitag"
              sVTAG = "Donnerstag"
            Case "Samstag"
              sVTAG = "Freitag"
            Case "Sonntag"
              sVTAG = "Samstag"
        End Select
        speicher_AEA_ZEITENvomVtag sVTAG
        
        Lese_AEA_Zeiten List1, Label1(9).Caption
    Case 6
        Loeschetag
    Case 7
        speicher_AEA_TAGE
        speicherExEinstellungen
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherExEinstellungen()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    Dim bo7 As Integer
    Dim bo8 As Integer
    Dim cBlKennung As String
    Dim cEndung As String
    Dim cFeldtrenner As String
    
    If Check2.Value = vbChecked Then
        bo1 = -1
    Else
        bo1 = 0
    End If
    
    If Check3.Value = vbChecked Then
        bo2 = -1
    Else
        bo2 = 0
    End If
    
    If Check4.Value = vbChecked Then
        bo3 = -1
    Else
        bo3 = 0
    End If
    
    If Check5.Value = vbChecked Then
        bo4 = -1
    Else
        bo4 = 0
    End If
    
    If Check6.Value = vbChecked Then
        bo5 = -1
    Else
        bo5 = 0
    End If
    
    If Check7.Value = vbChecked Then
        bo6 = -1
    Else
        bo6 = 0
    End If
    
    If Check8.Value = vbChecked Then
        bo7 = -1
    Else
        bo7 = 0
    End If
    
    If Check9.Value = vbChecked Then
        bo8 = -1
    Else
        bo8 = 0
    End If
    
    If Option1(1).Value = True Then
        cEndung = "txt"
    Else
        cEndung = "csv"
    End If
    
    If Option2(0).Value = True Then
        cFeldtrenner = "Tab"
    ElseIf Option2(1).Value = True Then
        cFeldtrenner = "Komma"
    ElseIf Option2(2).Value = True Then
        cFeldtrenner = "Semikolon"
    End If
    
    cBlKennung = Text2.Text

    loeschNEW "EXSTEU", gdApp
    CreateTableT3 "EXSTEU", gdApp
    
    sSQL = "Insert into EXSTEU (NURMITBESTAND,bl,EXNOR,blkennung,PLUSEAN,PLUSBEZEICH,DATEIENDUNG,FELDTRENNER,Shopartikel,ShopPreis,MITUEBERSCHRIFT) Values (" & bo1 & "," & bo2 & "," & bo3 & ",'" & cBlKennung & "'," & bo4 & "," & bo5 & ",'" & cEndung & "','" & cFeldtrenner & "'," & bo6 & "," & bo7 & "," & bo8 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherExEinstellungen"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Lese_AEA_Zeiten(Listx As ListBox, cDay As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    Dim lcount As Long
    
    lcount = 0
    Listx.Clear
    
    sSQL = "select * from ZEITAEA where ART = 'AEA' and TAG = '" & cDay & "' order by zeit"
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!zeit) Then
                cFeld = rsrs!zeit
            Else
                cFeld = 0
            End If
            cFeld = Format$(TimeValue(cFeld), "HH:MM")
            lcount = lcount + 1
            Listx.AddItem cFeld
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    Listx.Refresh
    
    If lcount > 0 Then
        anzeige "normal", lcount & " Zeiten", Label1(10)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "Lese_AEA_Zeiten"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Lese_AEA_TAGE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim i As Integer
    
    For i = 0 To 6
        Check1(i).Value = vbUnchecked
        Command2(i).Enabled = False
        Label2(i).Caption = ""
    Next i
    
    sSQL = "select * from TAGAEA where art = 'AEA'"
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Tag) Then
                Select Case rsrs!Tag
                    Case gcWochentag(1)
                        Check1(0).Value = vbChecked
                        Command2(0).Enabled = True
                    Case gcWochentag(2)
                        Check1(1).Value = vbChecked
                        Command2(1).Enabled = True
                    Case gcWochentag(3)
                        Check1(2).Value = vbChecked
                        Command2(2).Enabled = True
                    Case gcWochentag(4)
                        Check1(3).Value = vbChecked
                        Command2(3).Enabled = True
                    Case gcWochentag(5)
                        Check1(4).Value = vbChecked
                        Command2(4).Enabled = True
                    Case gcWochentag(6)
                        Check1(5).Value = vbChecked
                        Command2(5).Enabled = True
                    Case gcWochentag(7)
                        Check1(6).Value = vbChecked
                        Command2(6).Enabled = True
                End Select
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "Lese_AEA_TAGE"
    Fehler.gsFehlertext = "Im Programmteil Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeZeitenSummen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim lcount  As Long
    Dim i       As Integer
    
    For i = 0 To 6
        Label2(i).Caption = ""
    Next i
    
    For i = 1 To 7
    
        lcount = 0
        sSQL = "select * from ZEITAEA where TAG = '" & gcWochentag(i) & "' and art = 'AEA' "
        Set rsrs = gdApp.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveLast
            lcount = rsrs.RecordCount
            If lcount = 1 Then
                Label2(i - 1).Caption = lcount & " Zeit"
                Label2(i - 1).Refresh
            Else
                Label2(i - 1).Caption = lcount & " Zeiten"
                Label2(i - 1).Refresh
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "ZeigeZeitenSummen"
    Fehler.gsFehlertext = "Im Programmteil Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher_AEA_ZEITEN()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim czeit As String
    Dim cTag As String
    
    czeit = Format(DTPicker1.Value, "HH:MM")
    cTag = Label1(9).Caption
    
    sSQL = "Delete from ZEITAEA where Tag = '" & cTag & "' "
    sSQL = sSQL & " and zeit =  '" & czeit & "' "
    sSQL = sSQL & " and ART =  'AEA' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ZEITAEA (Tag,zeit,ART) values  "
    sSQL = sSQL & " ( "
    sSQL = sSQL & "  '" & cTag & "' "
    sSQL = sSQL & " , '" & czeit & "' "
    sSQL = sSQL & " , 'AEA' "
    sSQL = sSQL & " ) "
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_AEA_ZEITEN"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher_AEA_TAGE()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim i As Integer
    
    For i = 1 To 7
        If Check1(i - 1).Value = vbChecked Then
            sSQL = "Delete from TAGAEA where Tag = '" & gcWochentag(i) & "' and Art = 'AEA'"
            gdApp.Execute sSQL, dbFailOnError
            
            sSQL = "Insert into TAGAEA (Tag,ART) values  "
            sSQL = sSQL & " ( "
            sSQL = sSQL & "  '" & gcWochentag(i) & "' , 'AEA'"
            sSQL = sSQL & " ) "
            gdApp.Execute sSQL, dbFailOnError
        Else
            sSQL = "Delete from TAGAEA where Tag = '" & gcWochentag(i) & "' and Art = 'AEA'"
            gdApp.Execute sSQL, dbFailOnError
            
            sSQL = "Delete from ZEITAEA where Tag = '" & gcWochentag(i) & "' and Art = 'AEA'"
            gdApp.Execute sSQL, dbFailOnError
        End If
    Next i

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_AEA_TAGE"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher_AEA_ZEITENvomVtag(cVtag As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim cTag As String
    
    cTag = Label1(9).Caption
    
    sSQL = "Insert into ZEITAEA select zeit,'" & cTag & "' as tag ,Art from ZEITAEA "
    sSQL = sSQL & " where tag =  '" & cVtag & "' "
    sSQL = sSQL & " and art =  'AEA' "
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_AEA_ZEITENvomVtag"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Loeschetag()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim cTag As String
    
    List1.Clear
    cTag = Label1(9).Caption
    
    sSQL = "Delete from ZEITAEA "
    sSQL = sSQL & " where tag =  '" & cTag & "' "
    sSQL = sSQL & " and art =  'AEA' "
    gdApp.Execute sSQL, dbFailOnError
    
    List1.Refresh

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Loeschetag"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheEinzeln(czeit As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim cTag As String
    
    cTag = Label1(9).Caption
    
    sSQL = "Delete from ZEITAEA where Tag = '" & cTag & "' "
    sSQL = sSQL & " and zeit =  '" & czeit & "' "
    sSQL = sSQL & " and ART =  'AEA' "
    gdApp.Execute sSQL, dbFailOnError
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheEinzeln"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Frame4.Visible = True
    Label1(8).Caption = "Zeitpunkte für " & Check1(Index).Caption & " festlegen"
    Label1(8).Refresh
    
    Label1(9).Caption = gcWochentag(Index + 1)
    Label1(9).Refresh
    
    DTPicker1.Value = Format$(TimeValue(Now), "HH:MM:00")
    
    Lese_AEA_Zeiten List1, Label1(9).Caption
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command36_Click()
On Error GoTo LOKAL_ERROR


frmWKL216.Show 1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command36_Click"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    PositionierenWKL05
    alternativFarbform Me, lblUeberschrift
    Modul6.Skalieren Me, True, True: Schrift Me
    
    
    
    lese_Ex_Steu
    
    If gbNMB = True Then
        Check2.Value = vbChecked
    Else
        Check2.Value = vbUnchecked
    End If
    
    If gbBL = True Then
        Check3.Value = vbChecked
    Else
        Check3.Value = vbUnchecked
    End If
    
    If gbEXNOR = True Then
        Check4.Value = vbChecked
    Else
        Check4.Value = vbUnchecked
    End If
    
    If gbSHOPARTIKEL = True Then
        Check7.Value = vbChecked
    Else
        Check7.Value = vbUnchecked
    End If
    
    If gbPlusEAN = True Then
        Check5.Value = vbChecked
    Else
        Check5.Value = vbUnchecked
    End If
    
    If gbPlusBezeich = True Then
        Check6.Value = vbChecked
    Else
        Check6.Value = vbUnchecked
    End If
    
    If gbPlusShopPreis = True Then
        Check8.Value = vbChecked
    Else
        Check8.Value = vbUnchecked
    End If
    
    If gbMITUEBERSCHRIFT = True Then
        Check9.Value = vbChecked
    Else
        Check9.Value = vbUnchecked
    End If
    
    
    Text2.Text = gsBLKENNUNG
    
    If gsDATEIENDUNG = "txt" Then
        Option1(1).Value = True
    Else
        Option1(0).Value = True
    End If
    
    If gsFELDTRENNER = "Tab" Then
        Option2(0).Value = True
    ElseIf gsFELDTRENNER = "Komma" Then
        Option2(1).Value = True
    ElseIf gsFELDTRENNER = "Semikolon" Then
        Option2(2).Value = True
    End If
    
    If Not NewTableSuchenDBKombi("ZEITAEA", gdApp) Then
        CreateTableT2 "ZEITAEA", gdApp
    End If
    
    If Not NewTableSuchenDBKombi("TAGAEA", gdApp) Then
        CreateTableT2 "TAGAEA", gdApp
    End If
    
    If gbAuto_Export_Artikelbestand Then
        Frame1.Visible = True
        Lese_AEA_TAGE
        ZeigeZeitenSummen
    Else
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Bei den Exporteinstellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL05()
On Error GoTo LOKAL_ERROR

    With Frame4
        .Top = 720
        .Left = 120
        .Width = 9855
        .Height = 4935
    End With

    With Frame1
        .Top = 720
        .Left = 120
        .Width = 9855
        .Height = 4935
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL05"
    Fehler.gsFehlertext = "Bei den Zeitungsauswertung ist ein Fehler aufgetreten."
    
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




