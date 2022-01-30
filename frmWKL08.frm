VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL08 
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frmWKL08.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10080
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   57
      Top             =   360
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
      Caption         =   "Zeitungscheck"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame4 
      Caption         =   "Zeitpunkt festlegen"
      Height          =   1455
      Left            =   7920
      TabIndex        =   46
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   50
         Top             =   720
         Width           =   2295
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   4
         Left            =   7680
         TabIndex        =   48
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
         TabIndex        =   47
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
         Format          =   225837058
         CurrentDate     =   38457.8333333333
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
         TabIndex        =   55
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
         TabIndex        =   51
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Auswertungstag"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   3495
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   13
      Top             =   840
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
      Caption         =   "Einstellungen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Einstellungen"
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   9855
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   7
         Left            =   7680
         TabIndex        =   56
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
         Height          =   2055
         Left            =   4560
         TabIndex        =   24
         Top             =   1440
         Width           =   5175
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   6
            Left            =   4080
            TabIndex        =   38
            Top             =   1680
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   5
            Left            =   4080
            TabIndex        =   37
            Top             =   1440
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   4
            Left            =   4080
            TabIndex        =   36
            Top             =   1200
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   3
            Left            =   4080
            TabIndex        =   35
            Top             =   960
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   2
            Left            =   4080
            TabIndex        =   34
            Top             =   720
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   1
            Left            =   4080
            TabIndex        =   33
            Top             =   480
            Width           =   855
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
            Caption         =   "Uhrzeit"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   210
            Index           =   0
            Left            =   4080
            TabIndex        =   32
            Top             =   240
            Width           =   855
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
            TabIndex        =   31
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "samstags"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "freitags"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "donnerstags"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "mittwochs"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "dienstags"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "montags"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
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
            Left            =   1560
            TabIndex        =   45
            Top             =   1680
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   44
            Top             =   1440
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   43
            Top             =   1200
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   42
            Top             =   960
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   41
            Top             =   720
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   40
            Top             =   480
            Width           =   2415
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
            Left            =   1560
            TabIndex        =   39
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Auswertungsmodus"
         Height          =   1095
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "Zeitpunkt bezogen"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "tageweise"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "txt"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "aktuelle Tageszahl"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "LieferantenNr"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Dateiendung"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Emailbetreff"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Emailadresse"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kundennummer"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   240
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
      Caption         =   "Auswertung"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   3
      Left            =   1800
      TabIndex        =   58
      ToolTipText     =   "Kalender"
      Top             =   840
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
   Begin VB.Label Label3 
      Caption         =   "Sind keine FTP-Einstellungen vorgenommen, so bitte nur an einem Rechner im Netzwerk die 'ständige Internetverbindung' aktivieren."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   59
      Top             =   5040
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "666669 = ZEITSCHRIFTEN volle MwSt."
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "666668 = ZEITSCHRIFTEN erm. MwSt."
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Auswertungstag"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1335
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
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   4455
   End
   Begin VB.Label lblUeberschrift 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Zeitungsauswertung"
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
Attribute VB_Name = "frmWKL08"
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
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Select Case Index
        Case Is = 3
            Text1(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim lDatum      As Long
Dim sFText      As String
Dim sZeitung    As String
Dim ldatum1Jan  As Long
Dim lex         As Long
Dim cex         As String
Dim cPfad       As String
Dim sVTAG       As String

Dim cKdnr       As String

cPfad = gcDBPfad
If Right$(cPfad, 1) <> "\" Then
    cPfad = cPfad & "\"
End If
cPfad = cPfad & "ZEITUNG"

Select Case Index
    Case 0
        anzeigeNew "normal", "", lblAnzeige
        
        If Frame1.Visible = True Then
            speicherZeitungsEinstellungen
            speicherVMPTAGE
            
        End If
        
        If Text1(1).Text = "" Then
            anzeigeNew "rot", "Bitte eine gültige Emailadresse eingeben!", lblAnzeige
            Frame1.Visible = True
            Text1(1).SetFocus
            Exit Sub
        End If
        
        If Text1(2).Text = "" Then
            anzeigeNew "rot", "Bitte Kundennummer eingeben!", lblAnzeige
            Frame1.Visible = True
            Text1(2).SetFocus
            Exit Sub
        End If
        
        If IsDate(DateValue(Text1(0).Text)) Then
            lDatum = DateValue(Text1(0).Text)
            ldatum1Jan = DateValue("01.01." & Year(DateValue(Text1(0).Text)))
            lex = lDatum - ldatum1Jan + 1
            cex = CStr(lex)
            
            If Len(cex) = 1 Then
                cex = "00" & cex
            ElseIf Len(cex) = 2 Then
                cex = "0" & cex
            End If
        Else
            anzeigeNew "rot", "Bitte ein gültiges Datum eingeben!", lblAnzeige
            Text1(0).SetFocus
            Exit Sub
        End If
        
        cKdnr = Trim(Text1(2).Text)
        
        If Option1(0).Value = True Then
        
        ElseIf Option1(1).Value = True Then
            cex = "txt"
        End If
        
        Kill cPfad & "\" & cKdnr & ".*"
        sZeitung = cPfad & "\" & cKdnr & "." & cex
        
        If Option2(0).Value = True Then
            sFText = AUSwertungZP(lDatum, sZeitung, cKdnr, Val(Text1(4).Text))
        Else
            sFText = AUSwertungZP(0, sZeitung, cKdnr, Val(Text1(4).Text))
        End If
        
        If sFText <> "" Then
            anzeigeNew "rot", sFText, lblAnzeige
        Else
            gcBestellEmail.Attachment1 = sZeitung
            gcBestellEmail.Subject = Text1(3).Text
            gcBestellEmail.Message = "Zeitungsdaten"
            gcBestellEmail.Recipient = Text1(1).Text
            
            frmWKL129.Show 1
            
            gcBestellEmail.Attachment1 = ""
            gcBestellEmail.Subject = ""
            gcBestellEmail.Message = ""
            gcBestellEmail.Recipient = ""
        
            anzeigeNew "normal", sZeitung & " erfolgreich erstellt", lblAnzeige
        End If
    Case 1
        Unload frmWKL08
    Case 2
        Frame1.Visible = True
        LeseVMPTAGE
        ZeigeZeitenSummen
    Case 3
        speicherVMPZEITEN
        LeseVMPZeiten List1, Label1(9).Caption
        
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
        speicherVMPZEITENvomVtag sVTAG
        
        LeseVMPZeiten List1, Label1(9).Caption
        
    Case 6
        Loeschetag
    Case 7
        speicherZeitungsEinstellungen
        speicherVMPTAGE
        leseZeitSteu
    Case 8
        frmWKL174.Show 1
End Select

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub LeseVMPZeiten(Listx As ListBox, cDay As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    Dim lcount As Long
    
    lcount = 0
    Listx.Clear
    
    sSQL = "select * from ZEITVMP where TAG = '" & cDay & "' order by zeit"
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
    
    If lcount > 0 Then
        anzeige "normal", lcount & " Zeiten", Label1(10)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "LeseVMPZeiten"
    Fehler.gsFehlertext = "Im Programmteil Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseVMPTAGE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim i As Integer
    
    For i = 0 To 6
        Check1(i).Value = vbUnchecked
        Command2(i).Enabled = False
        Label2(i).Caption = ""
    Next i
    
    sSQL = "select * from TAGVMP"
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
    Fehler.gsFunktion = "LeseVMPTAGE"
    Fehler.gsFehlertext = "Im Programmteil Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeZeitenSummen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim lcount  As Long
    Dim i       As Integer

    For i = 1 To 7
    
        lcount = 0
        sSQL = "select * from ZEITVMP where TAG = '" & gcWochentag(i) & "' "
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
    Fehler.gsFehlertext = "Im Programmteil Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherZeitungsEinstellungen()
On Error GoTo LOKAL_ERROR

    Dim byEndung As Byte
    Dim iArt As Integer

    loeschNEW "ZEITSTEU", gdApp
    CreateTable "ZEITSTEU", gdApp
    
    If Option1(0).Value = True Then
        byEndung = 1
    ElseIf Option1(1).Value = True Then
        byEndung = 2
    End If
    
    If Option2(0).Value = True Then
        iArt = 1
    ElseIf Option2(1).Value = True Then
        iArt = 2
    End If
    
    InsertZeitsteu Text1(1).Text, Text1(3).Text, Text1(2).Text, byEndung, Val(Text1(4).Text), iArt
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherZeitungsEinstellungen"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherVMPZEITEN()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim czeit As String
    Dim cTag As String
    
    czeit = Format(DTPicker1.Value, "HH:MM")
    cTag = Label1(9).Caption
    
    sSQL = "Delete from ZEITVMP where Tag = '" & cTag & "' "
    sSQL = sSQL & " and zeit =  '" & czeit & "' "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ZEITVMP (Tag,zeit) values  "
    sSQL = sSQL & " ( "
    sSQL = sSQL & "  '" & cTag & "' "
    sSQL = sSQL & " , '" & czeit & "' "
    sSQL = sSQL & " ) "
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherVMPZEITEN"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherVMPTAGE()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim i As Integer
    
    For i = 1 To 7
        If Check1(i - 1).Value = vbChecked Then
            sSQL = "Delete from TAGVMP where Tag = '" & gcWochentag(i) & "' "
            gdApp.Execute sSQL, dbFailOnError
            
            sSQL = "Insert into TAGVMP (Tag) values  "
            sSQL = sSQL & " ( "
            sSQL = sSQL & "  '" & gcWochentag(i) & "' "
            sSQL = sSQL & " ) "
            gdApp.Execute sSQL, dbFailOnError
        Else
            sSQL = "Delete from TAGVMP where Tag = '" & gcWochentag(i) & "' "
            gdApp.Execute sSQL, dbFailOnError
            
            sSQL = "Delete from ZEITVMP where Tag = '" & gcWochentag(i) & "' "
            gdApp.Execute sSQL, dbFailOnError
        End If
    Next i

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherVMPTAGE"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherVMPZEITENvomVtag(cVtag As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim cTag As String
    
    cTag = Label1(9).Caption
    
    sSQL = "Insert into ZEITVMP select zeit,'" & cTag & "' as tag from ZEITVMP "
    sSQL = sSQL & " where tag =  '" & cVtag & "' "
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherVMPZEITENvomVtag"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Loeschetag()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim cTag As String
    
    List1.Clear
    cTag = Label1(9).Caption
    
    sSQL = "Delete from ZEITVMP "
    sSQL = sSQL & " where tag =  '" & cTag & "' "
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherVMPZEITENvomVtag"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertZeitsteu(cAdress As String, cbet As String, cKdnr As String, byEnd As Byte, lLinr As Long, iArt As Integer)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
    
cSQL = "Insert into Zeitsteu ( "
cSQL = cSQL & " adresse  "
cSQL = cSQL & ", betreff  "
cSQL = cSQL & ", kdnr "
cSQL = cSQL & ", Endung "
cSQL = cSQL & ", zlinr "
cSQL = cSQL & ", ART "
cSQL = cSQL & " ) "
cSQL = cSQL & " values  "
cSQL = cSQL & " ( '" & cAdress & "' "
cSQL = cSQL & ", '" & cbet & "' "
cSQL = cSQL & ", '" & cKdnr & "' "
cSQL = cSQL & ", " & byEnd & " "
cSQL = cSQL & ", " & lLinr & " "
cSQL = cSQL & ", " & iArt & " "
cSQL = cSQL & " ) "
gdApp.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertZeitsteu"
    Fehler.gsFehlertext = "Im Programmteil Zeitungsauswertung ist ein Fehler aufgetreten."
    
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
    
    LeseVMPZeiten List1, Label1(9).Caption
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    PositionierenWKL08
    alternativFarbform Me, lblUeberschrift
    Modul6.Skalieren Me, True, True: Schrift Me
    
    Text1(0).Text = DateValue(Now)
    
    If Not NewTableSuchenDBKombi("ZEITSTEU", gdApp) Then
        Frame1.Visible = True
    Else

        Text1(1).Text = gsVMPadresse
        Text1(3).Text = gsVMPbetreff
        Text1(2).Text = gsVMPKdNr
        Text1(4).Text = gsVMPzLinr

        If Val(gsVMPEndung) = 1 Then
            Option1(0).Value = True
        ElseIf Val(gsVMPEndung) = 2 Then
            Option1(1).Value = True
        End If

        If Val(gsVMPArt) = 1 Then
            Option2(0).Value = True
        ElseIf Val(gsVMPArt) = 2 Then
            Option2(1).Value = True
        End If

        If gsVMPArt = "2" Then
            Frame3.Visible = True

            LeseVMPTAGE
            ZeigeZeitenSummen
        Else
            Frame3.Visible = False
        End If

    End If

    If Not NewTableSuchenDBKombi("ZEITVMP", gdApp) Then
        CreateTable "ZEITVMP", gdApp
    End If

    If Not NewTableSuchenDBKombi("TAGVMP", gdApp) Then
        CreateTable "TAGVMP", gdApp
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL08()
On Error GoTo LOKAL_ERROR

    With Frame4
        .Top = 1320
        .Left = 120
        .Width = 9855
        .Height = 3615
    End With

    With Frame1
        .Top = 1320
        .Left = 120
        .Width = 9855
        .Height = 3615
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL08"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
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

Private Sub Option2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Option2(1).Value = True Then
        Frame3.Visible = True
        
        
    Else
        Frame3.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Bei der Zeitungsauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




