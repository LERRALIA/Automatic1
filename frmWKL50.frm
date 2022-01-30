VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmWKL50 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Druckereinstellungen"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   11910
   Icon            =   "frmWKL50.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   4
      Left            =   5880
      TabIndex        =   7
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Caption         =   "Fax Drucker"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Konfiguration Kunden-Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   6615
      Left            =   6360
      TabIndex        =   13
      Top             =   840
      Width           =   5415
      Begin VB.Frame Frame4 
         BackColor       =   &H00800000&
         Caption         =   "Zeileninhalt"
         ForeColor       =   &H0000FFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   34
         Top             =   4920
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   20
            TabIndex        =   42
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   20
            TabIndex        =   41
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   20
            TabIndex        =   40
            Top             =   1080
            Width           =   1455
         End
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   6
            Left            =   3240
            TabIndex        =   39
            Top             =   960
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Einzelwert, Zwischensumme"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   7
            Left            =   1920
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   3135
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Menge x Einzelpreis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   6
            Left            =   1920
            TabIndex        =   35
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00800000&
         Caption         =   "Display-Typ"
         ForeColor       =   &H0000FFFF&
         Height          =   2415
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   5175
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Sango"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   44
            Top             =   1680
            Width           =   855
         End
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   38
            Top             =   1440
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
            Caption         =   "Test"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Peacock (alt)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Peacock (Aures)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "JarlTech"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Epson"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00800000&
            Caption         =   "Aures neu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00800000&
            Caption         =   """AT"" auf dem Display - Eintrag in der LADECOM.CFG beachten"
            Height          =   495
            Index           =   0
            Left            =   2400
            TabIndex        =   37
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Caption         =   "Verbindung"
         ForeColor       =   &H0000FFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   5175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   8
            Left            =   3720
            TabIndex        =   27
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   26
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   25
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   3720
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   22
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   21
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "COM 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800000&
            Caption         =   "am Kassenbon-Drucker angeschlossen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Display vorhanden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "kein Display vorhanden bzw. dauerhaft ausgeschaltet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   5175
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   3
      Left            =   9720
      TabIndex        =   12
      Top             =   7560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
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
      Height          =   735
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Caption         =   "Etiketten Drucker"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Caption         =   "Listen Drucker"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Caption         =   "Kassenbon Drucker"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6135
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6135
   End
   Begin sevCommand3.Command Command1 
      Height          =   735
      Index           =   7
      Left            =   7800
      TabIndex        =   48
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Caption         =   "Gutschein Drucker"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   11160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   46
      Top             =   7080
      Width           =   5415
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Druckereinstellungen"
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
      TabIndex        =   33
      Top             =   120
      Width           =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   17
      Top             =   6360
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Fax-Drucker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   5640
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Etiketten-Drucker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   4920
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   4200
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Listen-Drucker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Kassenbon-Drucker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Gutschein-Drucker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   47
      Top             =   6720
      Width           =   3255
   End
End
Attribute VB_Name = "frmWKL50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub SchreibeIniPrinterWKL50()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim cDrucker As String
    
    iFileNr = FreeFile
    Open gcPfad & "PRINTWKL.INI" For Binary As #iFileNr
    Close iFileNr
    Kill gcPfad & "PRINTWKL.INI"
    
    cDrucker = Label2(0).Caption
    gcBonDrucker = Label2(0).Caption
    cDrucker = cDrucker & vbCrLf
    
    cDrucker = cDrucker & Label2(1).Caption
    gcListenDrucker = Label2(1).Caption
    cDrucker = cDrucker & vbCrLf
    
    cDrucker = cDrucker & Label2(2).Caption
    gcEtikettenDrucker = Label2(2).Caption
    cDrucker = cDrucker & vbCrLf
    
    cDrucker = cDrucker & Label2(3).Caption
    gcFaxDrucker = Label2(3).Caption
    cDrucker = cDrucker & vbCrLf
    
    cDrucker = cDrucker & Label2(4).Caption
    gcGutscheinDrucker = Label2(4).Caption
    cDrucker = cDrucker & vbCrLf
    
    iFileNr = FreeFile
    Open gcPfad & "PRINTWKL.INI" For Binary As #iFileNr
    Put #iFileNr, 1, cDrucker
    Close iFileNr
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeIniPrinterWKL50"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
End Sub
Private Sub SchreibeIniKasseWKL50()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim DSatz As KASSEINI_
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim iLfdNr As Integer
    Dim iWert1 As Integer
    Dim iWert2 As Integer
    Dim lcount As Long
    Dim lCount2 As Long
    Dim cDisplay As String
    Dim cSQL As String
    
    iLfdNr = -1
    iWert2 = 0
    cDisplay = ""
    
    If Option1(0).Value = True Then
        iWert1 = 0
    Else
        iWert1 = 1
        For lcount = 0 To 8
            If Option2(lcount).Value = True Then
                iWert2 = lcount
                If iWert2 > 0 Then
                    For lCount2 = 0 To 5
                        If Option3(lCount2).Value = True Then
                            Select Case lCount2
                                Case 0
                                    cDisplay = "Epson"
                                Case 1
                                    cDisplay = "JarlTech"
                                Case 2
                                    cDisplay = "Peacock"
                                Case 3
                                    cDisplay = "Peacock (alt)"
                                Case 4
                                    cDisplay = "Aures neu"
                                Case 5
                                    cDisplay = "Sango"
                            End Select
                        End If
                    Next lCount2
                End If
                Exit For
            End If
        Next lcount
    End If
    
    
    'Struktur der Datei:
    'LFDNR (Integer), WERT1 (Integer), WERT2 (Integer) = 6 Bytes je Satz
    
    iFileNr = FreeFile
    Open gcPfad & "KASSEWKL.INI" For Random As #iFileNr Len = Len(DSatz)
    If LOF(iFileNr) > 0 Then
        lAnzSatz = LOF(iFileNr) / Len(DSatz)
        For lAktSatz = 1 To lAnzSatz
            Get #iFileNr, lAktSatz, DSatz
            If DSatz.LFDNR = -1 Then
                DSatz.Wert1 = iWert1
                DSatz.WERT2 = iWert2
                Put #iFileNr, lAktSatz, DSatz
                Exit For
            End If
        Next lAktSatz
    Else
        lAktSatz = 1
        DSatz.LFDNR = -1
        DSatz.Wert1 = iWert1
        DSatz.WERT2 = 0
        Put #iFileNr, lAktSatz, DSatz
    End If
    
    Close iFileNr
    
    If cDisplay <> "" Then
        Kill gcPfad & "DISPLAY.CFG"
        iFileNr = FreeFile
        Open gcPfad & "DISPLAY.CFG" For Binary As #iFileNr
        Put #iFileNr, 1, cDisplay
        Close iFileNr
    End If
    
    
    If Frame4.Visible = True Then
    
        
        loeschNEW "DISINH", gdApp
        cSQL = "Create Table DISINH (KDEXM BIT )"
        gdApp.Execute cSQL, dbFailOnError

        If Option3(6).Value = True Then
            gbKDEXM = True
            cSQL = "Insert into DISINH (KDEXM) values (-1)"
            gdApp.Execute cSQL, dbFailOnError
        Else
            gbKDEXM = False
            cSQL = "Insert into DISINH (KDEXM) values (0)"
            gdApp.Execute cSQL, dbFailOnError
        End If
    End If
    
    If Text1(3).Text <> "" Then
        loeschNEW "DISPAUSE", gdApp
        cSQL = "Create Table DISPAUSE (DISPAUSE SINGLE )"
        gdApp.Execute cSQL, dbFailOnError

        cSQL = "Insert into DISPAUSE (DISPAUSE) values ('" & Text1(3).Text & "')"
        gdApp.Execute cSQL, dbFailOnError
       
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeIniKasseWKL50"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
        Fehlermeldung1
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    Dim cZeile1 As String
    Dim cZeile2 As String
    
    Select Case Index
        Case 0 To 2, 4, 7
            If List2.ListIndex < 0 Then
                MsgBox "Bitte einen Eintrag auswählen!", vbCritical, "STOP!"
                List2.SetFocus
            Else
                cDrucker = List2.list(List2.ListIndex)
                If Index = 0 Then
                    Label2(0).Caption = cDrucker
                ElseIf Index = 1 Then
                    Label2(1).Caption = cDrucker
                ElseIf Index = 2 Then
                    Label2(2).Caption = cDrucker
                ElseIf Index = 4 Then
                    Label2(3).Caption = cDrucker
                ElseIf Index = 7 Then
                    Label2(4).Caption = cDrucker
                End If
            End If
        Case Is = 3

            SchreibeIniKasseWKL50
            SchreibeIniPrinterWKL50
            setzedrucker gcListenDrucker

            Unload frmWKL50
        Case Is = 5
        
            SchreibeIniKasseWKL50
            SchreibeIniPrinterWKL50
        
            LeseKonfigurationKasseWKL50
                        
            gbDisplay = True
            If gbDisplay Then
                
                cZeile1 = Chr$(31) & Chr$(67) + "0"
                cZeile2 = ""
                ZeigeKundenDisplay_forTest cZeile1, cZeile2
                
                cZeile1 = gcDisplay
                
                If gbDisplaySeriell Then
                    cZeile2 = "COM: " & giDisplaySeriellComPort
                End If
                ZeigeKundenDisplay_forTest cZeile1, cZeile2
            End If
            
        Case 6
            SpeicherDisplayText Text1(0).Text, Text1(1).Text, Text1(2).Text
            LeseDisplayText
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Dim lAnzDrucker As Long
    Dim lcount As Long
    
    List1.Clear
    List2.Clear
    List1.AddItem "verfügbare Drucker"
    
    lAnzDrucker = Printers.Count
    For lcount = 0 To lAnzDrucker - 1
        Set Printer = Printers(lcount)
        List2.AddItem Printer.DeviceName
    Next lcount
    
    LeseKonfigurationKasseWKL50
    
    Label2(0).Caption = gcBonDrucker
    Label2(1).Caption = gcListenDrucker
    Label2(2).Caption = gcEtikettenDrucker
    Label2(3).Caption = gcFaxDrucker
    Label2(4).Caption = gcGutscheinDrucker
    
    LeseDISINH
    
    If gbKDEXM = True Then
        Option3(6).Value = True
        Option3(7).Value = False
    Else
        Option3(7).Value = True
        Option3(6).Value = False
    End If
    
    LeseDisplayText
    
    Text1(0).Text = gsMORGENTEXT
    Text1(1).Text = gsMITTAGTEXT
    Text1(2).Text = gsABENDTEXT
    
    LeseDIsPause
    
    Text1(3).Text = gsiDisPause

    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
    
End Sub
Private Sub LeseKonfigurationKasseWKL50()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim DSatz As KASSEINI_
    Dim iFileNr As Integer
    Dim cDisplay As String
    
    iFileNr = FreeFile
    
    Open gcPfad & "KASSEWKL.INI" For Random As #iFileNr Len = Len(DSatz)
    If LOF(iFileNr) > 0 Then
        lAnzSatz = LOF(iFileNr) / Len(DSatz)
        For lAktSatz = 1 To lAnzSatz
            Get #iFileNr, lAktSatz, DSatz
            If DSatz.LFDNR = -1 Then
                If DSatz.Wert1 = 1 Then
                    Option1(1).Value = True
                    Option2(DSatz.WERT2).Value = True
                    Exit For
                Else
                    Option1(0).Value = True
                    Exit For
                End If
            End If
        Next lAktSatz
        
    Else
        Option1(0).Value = True
    End If
    Close iFileNr
    
'    If DSatz.WERT2 = 9 Then
'        gbZweitMoni = True
'    End If
'
    If DSatz.Wert1 = 1 And DSatz.WERT2 > 0 Then
    
        gbDisplaySeriell = True
        giDisplaySeriellComPort = DSatz.WERT2
                    
        iFileNr = FreeFile
        
        Open App.Path & "\DISPLAY.CFG" For Binary As #iFileNr
        cDisplay = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cDisplay
        gcDisplay = cDisplay
        
        Close iFileNr
        Select Case cDisplay
            Case Is = "Epson"
                Option3(0).Value = True
            Case Is = "JarlTech"
                Option3(1).Value = True
            Case Is = "Peacock"
                Option3(2).Value = True
            Case Is = "Peacock (alt)"
                Option3(3).Value = True
            Case Is = "Aures neu"
                Option3(4).Value = True
            Case Is = "Sango"
                Option3(5).Value = True
        End Select
    Else
        gbDisplaySeriell = False
    End If
    
'    If gbDisplaySeriell Then
'        Select Case gcDisplay
'            Case Is = "Aures neu"
'                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
'                frmWKL20!MSComm3.Settings = "9600,N,8,1"
'                frmWKL20!MSComm3.InputLen = 0
'                frmWKL20!MSComm3.PortOpen = True
'            Case Is = "Sango"
'                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
'                frmWKL20!MSComm3.Settings = "9600,N,8,1"
'                frmWKL20!MSComm3.InputLen = 0
'                frmWKL20!MSComm3.PortOpen = True
'        End Select
'    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Then
        Resume Next
    ElseIf err.Number = 8002 Then
        MsgBox "Der COM - Port " & giDisplaySeriellComPort & " steht nicht zur Verfügung.", vbInformation, "Winkiss Hinweis:"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "LeseKonfigurationKasseWKL50"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
        Fehlermeldung1
    End If
    
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            
        Case Is = 1
            
            Option2(0).Value = True
            Frame2.Visible = True
            Frame4.Visible = True
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
End Sub


Private Sub Option2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            
            Frame3.Visible = False
            
        Case Else
            
'            Option3(0).Value = True
            Frame3.Visible = True
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
    

End Sub
Private Sub Positionieren()
    On Error GoTo LOKAL_ERROR
    
    
           
           
            
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Druckereinstellungen auf. "
    Fehlermeldung1
    
    

End Sub


