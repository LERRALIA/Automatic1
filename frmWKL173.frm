VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL173 
   BackColor       =   &H00C0C000&
   Caption         =   "Kundenanalyse"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL173.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame fraSerienB 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   975
      Left            =   11160
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   2295
      Begin sevCommand3.Command cmdSUeber 
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtSerienBHaupt 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   37
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00808000&
         Caption         =   "Text erstellen"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame fraEmail 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   7080
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Index           =   3
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   21
         Top             =   360
         Width           =   3735
      End
      Begin sevCommand3.Command cmdSenden 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Senden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "an Emailadresse"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "Betreff"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "Mitteilung"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   33
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame fraAusgabe 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   -2520
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraSort 
         Appearance      =   0  '2D
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   3720
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Postleitzahl"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Geburtsdatum"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Nachname"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   43
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackColor       =   &H00808000&
            Caption         =   "sortiert nach"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame fraFormat 
         Appearance      =   0  '2D
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4680
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Steuerdatei erw."
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Steuerdatei ein."
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "als Word-Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "als Excel-Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdListen 
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   28
         Top             =   2040
         Width           =   1935
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
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "schließen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Frame fraExport 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         Height          =   1455
         Left            =   2520
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
         Begin sevCommand3.Command cmdExport 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "in Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdExport 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "auf Diskette"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdExport 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "per Email"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Exportieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame fraEtiketten 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         Height          =   975
         Left            =   2760
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   495
         Begin sevCommand3.Command cmdEtikett 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Format: Zweckform 3475"
            Top             =   600
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "7,0 cm x 3,6 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Format: Zweckform 3653"
            Top             =   120
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "10,5 cm x 4,24 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin VB.Frame fraListen 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   3960
         TabIndex        =   10
         Top             =   -360
         Visible         =   0   'False
         Width           =   2295
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Kundenliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Bonusliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Telefon/Fax Liste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Geburtstagsliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   1935
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
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Adressliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Serienbriefvorlage"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Adressetiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Listen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   2
         X1              =   20
         X2              =   20
         Y1              =   0
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   6
         X1              =   6830
         X2              =   6830
         Y1              =   0
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   4
         X1              =   0
         X2              =   6840
         Y1              =   20
         Y2              =   20
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   3
         X1              =   0
         X2              =   6840
         Y1              =   2510
         Y2              =   2510
      End
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   11280
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "Kundenliste.doc"
      PrintFileType   =   17
      PrintFileLinesPerPage=   60
   End
   Begin sevCommand3.Command cmdEnd 
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   7920
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin sevCommand3.Command cmdPrint 
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   7440
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Ausgabe"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   5655
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   12632256
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   8880
      TabIndex        =   51
      Top             =   3120
      Width           =   2295
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
      Caption         =   "zurücksetzen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   5
      Left            =   10800
      TabIndex        =   50
      ToolTipText     =   "Kalender"
      Top             =   2640
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
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   49
      Top             =   4320
      Width           =   2295
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
      Caption         =   "alle zurücksetzen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   48
      Top             =   4800
      Width           =   2295
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
      Caption         =   "Verkaufsdetails"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   47
      Top             =   5280
      Width           =   2295
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
      Caption         =   "Kundendaten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label18 
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   53
      Top             =   6720
      Width           =   4575
   End
   Begin VB.Label lblAnzeige 
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   4
      Top             =   7080
      Width           =   10815
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunden - Analyse"
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
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   11280
      Y1              =   840
      Y2              =   840
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
      Left            =   8880
      TabIndex        =   52
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmWKL173"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPruef  As Integer
Dim bDat1   As Boolean
Dim bDat2   As Boolean
Dim bKauf   As Boolean
Dim bVorhanden As Boolean
Dim bEmail As Boolean
Dim bDis As Boolean
Dim bDat As Boolean
Dim bExcel As Boolean
Dim bWord As Boolean

Dim lAusgewählt As Long

Dim sdateiname As String
Dim sErstelldatum As String
Dim bAender As Boolean
Dim bNotAll As Boolean
Dim bClickAusgabe As Boolean
Private Sub Check1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkmannl_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkmannl_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkOKr_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkOKr_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkweibl_GotFocus()
    On Error GoTo LOKAL_ERROR
    If bVorhanden Then
        bAender = True
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkweibl_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdAusgabe_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim rsrs As Recordset
    Dim sHaupt As String
    
    Select Case Index
        Case Is = 0
            fraListen.Visible = True
            fraEtiketten.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
        Case Is = 1
            fraEtiketten.Visible = True
            fraSort.Visible = True
            fraListen.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
        Case Is = 2
            'Serienbriefvorlage
            fraSerienB.Visible = True
            
            If Not NewTableSuchenDBKombi("haupt", gdBase) Then
            
            Else
                Set rsrs = gdBase.OpenRecordset("Haupt", dbOpenTable)
                If Not rsrs.RecordCount = 0 Then
                    sHaupt = rsrs!texthaupt
                    txtSerienBHaupt.Text = sHaupt
                End If
                rsrs.Close: Set rsrs = Nothing
            End If
        
            fraEtiketten.Visible = False
            fraListen.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
        Case Is = 3
            fraExport.Visible = True
            fraListen.Visible = False
            fraEtiketten.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdAusgabe_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL173
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdEtikett_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
        loeschNEW "KUTTEN", gdBase
        CreateTableT2 "KUTTEN", gdBase
        
        sSQL = "Insert into KUTTEN Select  "
        sSQL = sSQL & " Knummer"
        sSQL = sSQL & ", KUERZEL"
        sSQL = sSQL & ", FIRMA"
        sSQL = sSQL & ", TITEL"
        sSQL = sSQL & ", NAME"
        sSQL = sSQL & ", VORNAME"
        sSQL = sSQL & ", STRASSE"
        sSQL = sSQL & ", PLZ"
        sSQL = sSQL & ", STADT"
        sSQL = sSQL & ", DATUM1"
        sSQL = sSQL & ", ANREDE"
        
        sSQL = sSQL & " from KUTEILME "
        
        If Option1(0).Value = True Then
    '        Sortierung 1
            sSQL = sSQL & " order by Month(Datum1),Day(Datum1)"
        ElseIf Option1(1).Value = True Then
    '        Sortierung 2
            sSQL = sSQL & " order by Name"
        ElseIf Option1(2).Value = True Then
    '        Sortierung 3
            sSQL = sSQL & " order by Plz"
        End If
        gdBase.Execute sSQL, dbFailOnError
        
        Select Case Index
            Case Is = 0
                'Etiketten groß
                If Modul6.FindFile(gcDBPfad, "aWKLavas.rpt") Then
                    reportbildschirm "spezial", "aWKLavas"
                Else
                    reportbildschirm "WKL017", "aWKLava"
                End If
                
               
            Case Is = 1
                'Etiketten klein
                
                If Modul6.FindFile(gcDBPfad, "aWKLavbs.rpt") Then
                    reportbildschirm "spezial", "aWKLavbs"
                Else
                    reportbildschirm "WKL017", "aWKLavb"
                End If
           
        End Select
    Else
        anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblAnzeige
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEtikett_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdExport_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            bEmail = True
            bDis = False
            bDat = False
            fraFormat.Visible = True
        Case Is = 1
            bDis = True
            bEmail = False
            bDat = False
            fraFormat.Visible = True
        Case Is = 2
            bDat = True
            bDis = False
            bEmail = False
            fraFormat.Visible = True
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdExport_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdListen_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Index <> 5 Then
        If NewTableSuchenDBKombi("KUTEILME", gdBase) = False Then
        
            anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblAnzeige
            Exit Sub
        End If
    End If
        Select Case Index
            Case Is = 0     'Adressenliste
                reportbildschirm "kaali", "aWKLavc"
            Case Is = 1     'Geburtstagsliste
            
                sSQL = "Update KUTEILME set datum1 = '31.12.2010' where datum1 is null"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update KUTEILME set datum1 = '31.12.2010' where datum1 = ''"
                gdBase.Execute sSQL, dbFailOnError
            
                reportbildschirm "kagli", "aWKLavd"
            Case Is = 2     'Telefonliste
                reportbildschirm "katli", "aWKLave"
            Case Is = 3     'Bonusliste
                reportbildschirm "kaboli", "aWKLavf"
            Case Is = 4     'Kundenliste
                reportbildschirm "kakuli", "aWKLavg"
            Case Is = 5
                fraAusgabe.Visible = False
                fraEmail.Visible = False
                fraExport.Visible = False
                fraSort.Visible = False
                fraSerienB.Visible = False
                bClickAusgabe = False
        End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdListen_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo LOKAL_ERROR
    
    If bClickAusgabe Then
        fraAusgabe.Visible = False
        bClickAusgabe = False
    Else
    
        KUTEILMEupdate
        
        If NewTableSuchenDBKombi("KUTEILME", gdBase) = False Then
            anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblAnzeige
            Exit Sub
        End If
    
        fraAusgabe.Visible = True
        fraListen.Visible = False
        fraExport.Visible = False
        fraEtiketten.Visible = False
        fraEmail.Visible = False
        fraFormat.Visible = False
        fraSerienB.Visible = False
        
        bClickAusgabe = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPrint_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdSenden_Click()
    On Error GoTo LOKAL_ERROR

    Dim sMailadress As String
    Dim sMessage As String
    Dim sBetreff As String
    Dim sPunkt As String
    Dim lPos As Long
    Dim lpos1 As Long
    Dim cPfad1 As String
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    lPos = InStr(Text1(0).Text, "@")
    
    If lPos > 0 Then
        sPunkt = Right(Text1(0).Text, Len(Text1(0)) - lPos)
        lpos1 = InStr(sPunkt, ".")
    End If
    
    
    
    If Text1(0).Text = "" Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf lPos = 0 Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf lpos1 <= 1 Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf Right(Text1(0).Text, 1) = "." Then
        Text1(0).SetFocus
        
        Exit Sub
    
    Else
        If bExcel Then
        
            Dim Result      As String
            Dim Buff        As String
            Dim sZeitung    As String
            
            sZeitung = cPfad1 & "BOX\Kunden.xls"
        
            
            Buff = "mailto:" & Trim(Text1(0).Text)
            Buff = Buff & "?Subject=" & Trim(Text1(1).Text)
            Buff = Buff & "&Body=" & Trim(Text1(3).Text)
            Buff = Buff & "&Attach=" + Chr$(34) & sZeitung + Chr$(34)
            
        
            Result = ShellExecute(0&, "open", Buff, "", "", 6)
    
        ElseIf bWord Then
        
            cr2.ReportFileName = cPfad1 & "aWKLavg.rpt"
            cr2.PrintFileName = cPfad1 & "BOX\Kundenliste.doc"
            cr2.PrintFileType = crptRTF
            cr2.Destination = 3
            
            sMailadress = Text1(0).Text
            sBetreff = Text1(1).Text
            sMessage = Text1(3).Text
            
            cr2.EMailToList = sMailadress
            cr2.EMailMessage = sMessage
            cr2.EMailSubject = sBetreff
            cr2.Action = 1
            
        End If
        
        fraEmail.Visible = False
    End If
    

    bExcel = False
    bWord = False
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdSenden_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zusammenstellunganzeigen()
    On Error GoTo LOKAL_ERROR
    
    Tabelleerstellen
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
        Tabellefuellen
        
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zusammenstellunganzeigen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Long
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
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabelleerstellen()
    On Error GoTo LOKAL_ERROR

    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 8
        .FixedCols = 0
        .FixedRows = 1
        
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 620
        .Text = "OK"
   
        
        .Col = 1
        .ColWidth(1) = 800
        .Text = "Kundennr"
        
        .Col = 2
        .ColWidth(2) = 1500
        .Text = "Vorname"
        
        .Col = 3
        .ColWidth(3) = 1600
        .Text = "Name"
        
        .Col = 4
        .ColWidth(4) = 1600
        .Text = "Straße"
        
        .Col = 5
        .ColWidth(5) = 600
        .Text = "Plz"
        
        .Col = 6
        .ColWidth(6) = 1600
        .Text = "Ort"
        
        .Col = 7
        .ColWidth(7) = 1000
        .Text = "Geburtstag"

    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabelleerstellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellefuellen()
    On Error GoTo LOKAL_ERROR

    Dim rsKUTEILME      As Recordset
    Dim lrow            As Long
    Dim lWert           As Long
    Dim sWert           As String
    Dim lCounter        As Long
    
    Set rsKUTEILME = gdBase.OpenRecordset("KUTEILME", dbOpenTable)
    
    lrow = 1
    If Not rsKUTEILME.EOF Then
        rsKUTEILME.MoveFirst
        
        MSHFLEX1.Redraw = False
        
        anzeige "normal", "Kunden werden ermittelt...", lblAnzeige
        
        pbrZeit.Visible = True
        pbrZeit.Max = 300
        
        Do While Not rsKUTEILME.EOF
            
            lrow = lrow + 1
            lCounter = lCounter + 1
            
            If lCounter = 300 Then
                lCounter = 0
            End If
            pbrZeit.Value = lCounter
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = "X"
            
            If Not IsNull(rsKUTEILME!knummer) Then
                lWert = rsKUTEILME!knummer
            Else
                lWert = 0
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = lWert
            
            Dim sKUNDNR     As String
            Dim cAWM        As String
            
            
            sKUNDNR = lWert
            cAWM = ""
            If sKUNDNR <> "" Then
                cAWM = WhatIsAwmKU(sKUNDNR)
            Else
                
            End If
            
            If cAWM = "" Then cAWM = "0"
            FaerbenFlexHKunde cAWM, MSHFLEX1, 1, lrow
            
            If Not IsNull(rsKUTEILME!vorname) Then
                sWert = rsKUTEILME!vorname
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!name) Then
                sWert = rsKUTEILME!name
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!strasse) Then
                sWert = rsKUTEILME!strasse
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 4
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!Plz) Then
                sWert = rsKUTEILME!Plz
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 5
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!STADT) Then
                sWert = rsKUTEILME!STADT
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 6
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!Datum1) Then
                sWert = rsKUTEILME!Datum1
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 7
            MSHFLEX1.Text = Trim$(sWert)
    
            rsKUTEILME.MoveNext
        Loop
        pbrZeit.Visible = False
    End If
    rsKUTEILME.Close
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    lAusgewählt = lrow
    
    If lrow > 1 Then
        anzeige "normal", lrow & " Kunden wurden ermittelt.", lblAnzeige
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
    ElseIf lrow = 1 Then
        anzeige "normal", lrow & " Kunde wurden ermittelt.", lblAnzeige
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
    Else
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblAnzeige
        anzeige "normal", "", Label18
        
        pbrZeit.Visible = False
        Exit Sub
    End If
    
    
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellefuellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdFormat_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    cmdFormat(Index).Enabled = False
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim i           As Integer
    Dim cDatname    As String
    Dim rsrs        As Recordset
    Dim dUmsatzLJ   As Double
    Dim dUmsatzVJ   As Double
    
    cDatname = "KundenA" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
    
    Select Case Index
        Case Is = 0     'EXCEL

            
                loeschNEW "KunExc", gdBase
                
                gsZSpalte = ""
                gstab = "KUEX"
                frmWKL36.Show 1
                
                'dannach Tablay auswerten
                
                Tabcheck "KUEX"
                FormatGridOverTablay "KUEX"
            
                Set rsrs = gdBase.OpenRecordset("Select * from Kuteilme")
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    Do While Not rsrs.EOF
                        
                        If Not IsNull(rsrs!knummer) Then
                        
                            dUmsatzLJ = umsatzLFJahr1(rsrs!knummer)

                            dUmsatzVJ = umsatzVJJahr(rsrs!knummer)
                
                        End If
                        rsrs.Edit
                        
                        rsrs!UMSLJ = dUmsatzLJ
                        rsrs!UMSVJ = dUmsatzVJ
                        rsrs.Update
                        rsrs.MoveNext
                    Loop
            
                End If
                rsrs.Close: Set rsrs = Nothing
                
                sSQL = " Update KUTEILME set Datum1 = 0 where Datum1 = '' "
                gdBase.Execute sSQL, dbFailOnError
                sSQL = " Update KUTEILME set Datum1 = 0 where Datum1 is null "
                gdBase.Execute sSQL, dbFailOnError
            
                If byAnzahlSpalten > 0 Then
                    sSQL = "Select " & sSpaltenbez(0) & " "
                    
                    If byAnzahlSpalten > 1 Then
                        For i = 1 To byAnzahlSpalten - 1
                            sSQL = sSQL & " , " & sSpaltenbez(i) & " "
                        Next i
                    End If
                Else
                    Exit Sub
                End If
                sSQL = sSQL & " into KunExc from KUTEILME"
                
'                MsgBox sSQL
                gdBase.Execute sSQL, dbFailOnError
                
                
                
                
                If bDat Then
                    cdatei = cPfad1 & "BOX\" & cDatname
                    cPfad = cPfad1 & "BOX"
                    
                    sSQL = "Select * "
                    sSQL = sSQL & " into KunExc IN '" & cdatei & "' 'Excel 8.0;' from KunExc "
                    gdBase.Execute sSQL, dbFailOnError

                    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
                
                ElseIf bDis Then
                
                    Screen.MousePointer = 11
                    cdatei = "a:\" & cDatname
                    cPfad = "a:"
                   
                    sSQL = "Select * into KunExc IN '" & cdatei & "' 'Excel 8.0;' from KunExc "
                    gdBase.Execute sSQL, dbFailOnError
                     
                    Screen.MousePointer = 0
                    MsgBox "Diese Datei ist auf der Diskette mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
                
                ElseIf bEmail Then
                
                    cdatei = cPfad1 & "BOX\" & cDatname
                    cPfad = cPfad1 & "BOX"
                    
                    sSQL = "Select * into KunExc IN '" & cdatei & "' 'Excel 8.0;' from KunExc "
                    gdBase.Execute sSQL, dbFailOnError
                
                
                    bExcel = True
    
                    fraEmail.Visible = True
                    Text1(0).SetFocus
    
                End If
        Case Is = 1     'Word bzw RTF
            If bDat Then
            
'                Pause (2)

                cr2.ReportFileName = cPfad1 & "aWKLavg.rpt"
                cr2.PrintFileName = cPfad1 & "BOX\Kundenliste.doc"
                cr2.PrintFileType = crptRTF
                cr2.Destination = 2
                cr2.Action = 1
                
                MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: Kundenliste.doc abgespeichert", vbInformation, "Winkiss Information:"
            ElseIf bDis Then

'                Pause (2)
                
                cr2.ReportFileName = cPfad1 & "aWKLavg.rpt"
                cr2.PrintFileName = "a:\Kundenliste.doc"
                cr2.PrintFileType = crptRTF
                cr2.Destination = 2
                cr2.Action = 1
                
                
                MsgBox "Diese Datei ist auf der Diskette mit dem Namen: Kundenliste.doc abgespeichert", vbInformation, "Winkiss Information:"
            ElseIf bEmail Then
                bWord = True
                fraEmail.Visible = True
                Text1(0).SetFocus
               
            End If
            
        Case Is = 3     'Steuerdatei einfach
            
            loeschNEW "stdatei", gdBase
        
            sSQL = "Select KNUMMER, KUERZEL, FIRMA, TITEL, NAME, VORNAME, STRASSE, PLZ, STADT, TEL, FAXNR "
            sSQL = sSQL & ", ANREDE,Kurztext1,Datum1,Geschlecht into Stdatei from KUTEILME"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            If bDat Then
                cdatei = cPfad1 & "BOX\StDatei.dbf"
                cPfad = cPfad1 & "BOX"
                Kill cdatei
                
                If NewTableSuchenDBKombi("StDatei", gdBase) Then
                    sSQL = "Select * into StDatei IN '" & cPfad & "' 'dbase IV;' from StDatei "
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

                    Screen.MousePointer = 0
                    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: StDatei.dbf abgespeichert", vbInformation, "Winkiss Information:"
                End If

            ElseIf bDis Then
                Screen.MousePointer = 11
                cdatei = "a:\StDatei.dbf"
                cPfad = "a:"
                Kill cdatei
                
                If NewTableSuchenDBKombi("StDatei", gdBase) Then
                    sSQL = "Select * into StDatei IN '" & cPfad & "' 'dbase IV;' from StDatei "
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                
                    Screen.MousePointer = 0
                    MsgBox "Diese Datei ist auf der Diskette mit dem Namen: StDatei.dbf abgespeichert", vbInformation, "Winkiss Information:"
                End If
            End If
            
        Case Is = 2     'Steuerdatei erweitert
        
            loeschNEW "stdater", gdBase
            
            sSQL = "Select * into Stdater from KUTEILME"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            If bDat Then
                cdatei = cPfad1 & "BOX\StDater.dbf"
                cPfad = cPfad1 & "BOX"
                Kill cdatei
                

                If NewTableSuchenDBKombi("StDater", gdBase) Then
                    sSQL = "Select * into StDater IN '" & cPfad & "' 'dbase IV;' from StDater "
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

                    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: StDater.dbf abgespeichert", vbInformation, "Winkiss Information:"
                End If

            ElseIf bDis Then
                Screen.MousePointer = 11
                cdatei = "a:\StDater.dbf"
                cPfad = "a:"
                Kill cdatei
                
                If NewTableSuchenDBKombi("StDater", gdBase) Then
                    sSQL = "Select * into StDater IN '" & cPfad & "' 'dbase IV;' from StDater "
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

                    Screen.MousePointer = 0
                    MsgBox "Diese Datei ist auf der Diskette mit dem Namen: StDater.dbf abgespeichert", vbInformation, "Winkiss Information:"
                End If
            End If
            
    End Select
    
    Else
        anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblAnzeige
    End If
    
    cmdFormat(Index).Enabled = True

    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 75 Then
        Resume Next
    ElseIf err.Number = 20530 Or err.Number = 3051 Then
        Screen.MousePointer = 0
        MsgBox "Sie haben keine Diskette eingelegt", vbInformation, "Winkiss Hinweis"
    ElseIf err.Number = 20999 Then
        Screen.MousePointer = 0
        MsgBox "Bitte nutzen Sie ein anderes Ausgabeformat! Die Ausgabe in diesem Format ist nicht möglich. ", vbInformation, "Winkiss Hinweis"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdFormat_Click"
        Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten. " & Index
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub cmdSUeber_Click()
    On Error GoTo LOKAL_ERROR
    Dim sHaupt As String
    Dim sSQL As String
    Dim sPfad As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    fraSerienB.Visible = False
    
    sHaupt = txtSerienBHaupt.Text
    
    loeschNEW "Haupt", gdBase
    
    sSQL = "create table haupt ("
    sSQL = sSQL & " texthaupt memo )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into haupt"
    sSQL = sSQL & "(texthaupt) "
    sSQL = sSQL & "values ("
    sSQL = sSQL & "'" & sHaupt & "' "
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    Pause (1)
    reportbildschirm "kaser", "aWKLavh"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdSUeber_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim lrow As Long
Dim cFarbkenn As String
Dim iRet As Integer
Dim ctmp As String

Select Case Index

    Case 0 'kundendaten
        
        lrow = Val(MSHFLEX1.Row)
        If lrow > 0 Then
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = 1
            gcKundenNr = MSHFLEX1.Text
            iKasse = 2
            frmWKL13.Show 1
        End If
    Case 1
        Screen.MousePointer = 0
                
        gsBackcolor = Label15.BackColor
        gsForecolor = Label15.ForeColor
        gsKundenFarbe = Label15.Tag
        
        frmWKL65.Show 1
        
        Label15.BackColor = gsBackcolor
        Label15.ForeColor = glS1
        Label15.Tag = gsKundenFarbe
        If gsKundenFarbe <> "" Then
            Label15.Caption = "Farbauswahl"
        Else
            Label15.Caption = "alle Farben"
        End If
    Case 2 'historie
        lrow = Val(MSHFLEX1.Row)
        If lrow > 0 Then
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = 1
            gckundnr = MSHFLEX1.Text
            
            gckundnr = Trim$(gckundnr)
            gsARTNR = ""
            
            If gckundnr <> "" Then
                frmWKL74.Show 1
            End If
        End If
    Case 3
        If Command1(3).Caption = "alle zurücksetzen" Then
        
            SchalteKunden (2)
            Command1(3).Caption = "alle auswählen"
        ElseIf Command1(3).Caption = "alle auswählen" Then
            SchalteKunden (3)
            Command1(3).Caption = "alle zurücksetzen"
        End If
    Case 5
        Screen.MousePointer = 0
        
        gsBackcolor = Label4(32).BackColor
        gsForecolor = Label4(32).ForeColor
        gsKundenFarbe = Label4(32).Tag
        
        frmWKL65.Show 1
        
        Label4(32).BackColor = gsBackcolor
        Label4(32).ForeColor = gsForecolor
        Label4(32).Tag = gsKundenFarbe
        If gsKundenFarbe <> "" Then
            Label4(32).Caption = "Farbauswahl"
        Else
            Label4(32).Caption = "alle Farben"
        End If
        
    Case 6
        ctmp = Trim$(Label4(32).Tag)
        If ctmp <> "" Then
            cFarbkenn = ermFarbeKU(ctmp)
        Else
            cFarbkenn = "alle Farben"
            SchalteKunden (2)
            Exit Sub
            ctmp = "0"
        End If
        
        If cFarbkenn = "" Then cFarbkenn = "ohne Kennzeichen"
        
        iRet = MsgBox("Möchten Sie jetzt alle Kunden aus der Tabelle mit dem Farbkennzeichen '" & cFarbkenn & "' zurücksetzen?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            Screen.MousePointer = 11
            SchalteKunden (4)
            Screen.MousePointer = 0
            
        End If
        
End Select
            
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchalteKunden(iSchaltung As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    Dim ctmp    As String
    Dim cAWM    As String
    Dim sKUNDNR As String
    
    If iSchaltung = 3 Then
        lAusgewählt = 0
    End If
    
    If iSchaltung = 2 Then
        lAusgewählt = 0
    End If
    
    
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = 0
    MSHFLEX1.Redraw = False
    For lrow = 1 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If iSchaltung = 2 Then
            MSHFLEX1.Text = ""
        End If
        If iSchaltung = 4 Then
        
            'ja aber hat der kunden bestimmte farbe
            
           
            anzeige "normal", lrow & "...", lblAnzeige
                
            
            ctmp = Trim$(Label4(32).Tag)
            If ctmp = "" Then ctmp = "0"
            
            MSHFLEX1.Col = 1
            sKUNDNR = MSHFLEX1.Text
            
            cAWM = ""
            If sKUNDNR <> "" Then
                cAWM = WhatIsAwmKU(sKUNDNR)
            End If
            
            If cAWM = ctmp Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = lcol
                MSHFLEX1.Text = ""
                lAusgewählt = lAusgewählt - 1
            End If
        End If
        
        If iSchaltung = 3 Then
            MSHFLEX1.Text = "X"
            lAusgewählt = lAusgewählt + 1
        End If
    Next lrow
    
    MSHFLEX1.Redraw = True
    
    If lAusgewählt > 1 Then
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
    ElseIf lAusgewählt = 1 Then
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
    Else
        anzeige "normal", "", Label18
    End If
    
    With MSHFLEX1
        .Row = 1
        .Col = 0
        .SetFocus
    End With
    
Exit Sub
LOKAL_ERROR:
    

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchalteKunden"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Sub WKLavPositionieren()
    On Error GoTo LOKAL_ERROR
    
    With MSHFLEX1
        .Height = 5655
        .Left = 480
        .Top = 960
        .Width = 8175
    End With
    
    With fraAusgabe
        .Top = 3840
        .Left = 3120
        .Height = 2535
        .Width = 6855
    End With
    
    With fraListen
        .Top = 120
        .Left = 2520
        .Height = 2295
        .Width = 2175
    End With
    
    With fraEtiketten
        .Top = 400
        .Left = 2520
        .Height = 975
        .Width = 2175
    End With
    
    With fraEmail
        .Top = 1200
        .Left = 3120
        .Height = 2415
        .Width = 6855
    End With
    
    With fraSerienB
        .Top = 1200
        .Left = 3120
        .Height = 2415
        .Width = 6855
    End With
    
    With fraFormat
        .Top = 120
        .Left = 4680
        .Height = 1815
        .Width = 2055
    End With
    
    With fraSort
        .Top = 120
        .Left = 4680
        .Height = 1815
        .Width = 2055
    End With
    
    With fraExport
        .Top = 960
        .Left = 2520
        .Height = 1455
        .Width = 2175
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKLavPositionieren"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    WKLavPositionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    bKauf = False
    bDat1 = False
    bDat2 = False
    bVorhanden = False
    bAender = False
    bNotAll = False
    bClickAusgabe = False
    bEmail = False
    bDis = False
    bDat = False
    bExcel = False
    bWord = False
    
    sdateiname = ""
    sErstelldatum = ""
    
    Zusammenstellunganzeigen
    
    anzeige "normal", "", lblAnzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    loeschNEW "KUTEILME", gdBase 'Kundenanalyse
    loeschNEW "Kuteil", gdBase
    loeschNEW "KUTTEN", gdBase

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




Private Sub MSHFLEX1_DblClick()
 On Error GoTo LOKAL_ERROR
    Dim lcol As Long
    Dim sSQL As String
    Dim rs As Recordset
    Dim sSortKrit As String
    
    If MSHFLEX1.Row = 1 Then
    
        lcol = MSHFLEX1.Col
        Select Case lcol
            Case Is = 1
            sSortKrit = " order by  Knummer"
            Case Is = 2
            sSortKrit = " order by  Vorname"
            Case Is = 3
            sSortKrit = " order by  Name"
            Case Is = 4
            sSortKrit = " order by  Strasse"
            Case Is = 5
            sSortKrit = " order by  Plz"
            Case Is = 6
            sSortKrit = " order by  stadt"
            Case Is = 7
            sSortKrit = " order by  datum1"
        End Select
        loeschNEW "Kutte", gdBase
        
        
        sSQL = "select * into kutte from KUTEILME " & sSortKrit
        
        If byteSortReihen = 1 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " desc"
            End If
            byteSortReihen = 2
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 1
        ElseIf byteSortReihen = 2 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " asc"
            End If
            byteSortReihen = 1
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 2
        End If
        
        gdBase.Execute sSQL
        
        loeschNEW "KUTEILME", gdBase
        sSQL = "select * into KUTEILME from KUTTE "
        gdBase.Execute sSQL
        loeschNEW "Kutte", gdBase
    Else
    
        MSHFLEX1.Col = 0
        If MSHFLEX1.Text = "X" Then
            MSHFLEX1.Text = ""
            lAusgewählt = lAusgewählt - 1
        Else
            MSHFLEX1.Text = "X"
            lAusgewählt = lAusgewählt + 1
        End If
        
        If lAusgewählt > 1 Then
            anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
        ElseIf lAusgewählt = 1 Then
            anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
        Else
            anzeige "normal", "", Label18
        End If
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KUTEILMEupdate()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    Dim cKdnr As String
    Dim sSQL As String
    
    
    MSHFLEX1.Redraw = False
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = 0
    
    For lrow = 2 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If MSHFLEX1.Text = "" Then
            MSHFLEX1.Col = 1
            cKdnr = MSHFLEX1.Text
            If IsNumeric(cKdnr) Then
                sSQL = "Delete from KUTEILME where knummer = " & cKdnr
                gdBase.Execute sSQL, dbFailOnError
            End If
        End If
    Next lrow
    
    MSHFLEX1.Redraw = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KUTEILMEupdate"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

