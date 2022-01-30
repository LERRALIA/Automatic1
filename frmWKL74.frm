VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL74 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   " - Kunden Verkauf"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   9960
      MouseIcon       =   "frmWKL74.frx":0000
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL74.frx":030A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   30
      Top             =   240
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   10200
      MouseIcon       =   "frmWKL74.frx":0394
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL74.frx":069E
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   31
      Top             =   240
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   10440
      MouseIcon       =   "frmWKL74.frx":0728
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL74.frx":0A32
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   32
      Top             =   240
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   10680
      MouseIcon       =   "frmWKL74.frx":0ABC
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL74.frx":0DC6
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   33
      Top             =   240
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   10920
      MouseIcon       =   "frmWKL74.frx":0E50
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL74.frx":115A
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   34
      Top             =   240
      Width           =   285
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   16
      Top             =   240
      Width           =   375
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
   Begin sevCommand3.Command Command3 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
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
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11775
      Begin VB.ListBox List7 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2760
         TabIndex        =   43
         Top             =   6240
         Width           =   4815
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   2760
         TabIndex        =   38
         Top             =   4680
         Width           =   4815
      End
      Begin sevCommand3.Command Command3 
         Height          =   285
         Index           =   3
         Left            =   10560
         TabIndex        =   36
         Top             =   2280
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
         Caption         =   "mehr"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   285
         Index           =   2
         Left            =   10560
         TabIndex        =   35
         Top             =   120
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
         Caption         =   "mehr"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   7680
         TabIndex        =   28
         Top             =   2640
         Width           =   3975
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   7680
         TabIndex        =   18
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   1695
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   7455
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7455
         Begin VB.OptionButton Option4 
            Caption         =   "Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   960
            TabIndex        =   45
            Tag             =   "bezeich desc"
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Fil Datum"
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
            Index           =   8
            Left            =   3480
            TabIndex        =   8
            Tag             =   "Filiale , adate desc"
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Menge"
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
            Index           =   7
            Left            =   4920
            TabIndex        =   7
            Tag             =   "menge desc"
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale"
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
            Left            =   2520
            TabIndex        =   6
            Tag             =   "Filiale desc"
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
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
            Index           =   4
            Left            =   6000
            TabIndex        =   5
            Tag             =   "Bednr desc"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Datum"
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
            Left            =   0
            TabIndex        =   4
            Tag             =   "adate desc"
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Sortierung nach"
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
            Index           =   7
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   2
         Top             =   6360
         Width           =   2055
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
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   7455
      End
      Begin sevCommand3.Command Command3 
         Height          =   285
         Index           =   4
         Left            =   6000
         TabIndex        =   41
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Caption         =   "Terminkalender"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   285
         Index           =   5
         Left            =   6000
         TabIndex        =   44
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Caption         =   "Einlösen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Bonusauszahlungen"
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
         Index           =   14
         Left            =   2760
         TabIndex        =   42
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "+ mehr anzeigen"
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
         Index           =   13
         Left            =   120
         MouseIcon       =   "frmWKL74.frx":11E4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   40
         Top             =   6480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Termine"
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
         Left            =   2760
         TabIndex        =   39
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Schwerpunkt/Artikelgruppe"
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
         Index           =   12
         Left            =   7680
         TabIndex        =   29
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Ø Stückzahl pro Bon"
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
         Index           =   11
         Left            =   9600
         TabIndex        =   27
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kaufvorgänge insgesamt:"
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
         Index           =   10
         Left            =   9600
         TabIndex        =   26
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kaufvorgänge insges:"
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
         Index           =   9
         Left            =   9600
         TabIndex        =   25
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Ø Euro pro Bon"
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
         Index           =   8
         Left            =   9600
         TabIndex        =   24
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Jahressummen"
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
         Index           =   6
         Left            =   1320
         TabIndex        =   23
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "insgesamt:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ø Euro pro Bon"
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
         Left            =   9600
         TabIndex        =   21
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Ø Stückzahl pro Bon"
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
         Left            =   9600
         TabIndex        =   20
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Schwerpunkt/Lieferant"
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
         Left            =   7680
         TabIndex        =   19
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Jahressummen"
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
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Klicken Sie auf die Sterne!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   37
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
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
      Left            =   5280
      TabIndex        =   17
      Top             =   240
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblanzeige 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunden Verkauf"
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
      TabIndex        =   12
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmWKL74"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOrder As String
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    With Frame5
        .Height = 6855
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cDatum As String
    Dim lLFNR As Long
    
    Select Case index
    
    Case 0
        Unload frmWKL74
    Case 1
        drucken sOrder
    Case 2
        If IsNumeric(gckundnr) Then
            frmWKL158.Show 1
            'hier die gckundnr nicht löschen
        End If
    Case 3
        If IsNumeric(gckundnr) Then
        
            frmWKL158.Show 1
            'hier die gckundnr nicht löschen
        End If
    Case 4
        
        If List6.ListIndex < 0 Then

        Else
            cLBSatz = List6.list(List6.ListIndex)
            gcTerm_Datum = Mid(cLBSatz, 1, 8)
        End If
    
        Screen.MousePointer = 11
        
        If gsfrmComeFrom = "Terminkalender" Then
            Unload frmWKL74 'Historie
            Unload frmWKL134 'Kunde suchen
            Unload frmWKL82 'Terminkalender
        End If
        
        frmWKL82.Show 1
        
        gcTerm_Datum = ""
        Screen.MousePointer = 0
        
    Case 5
        
        If List7.ListIndex < 0 Then
            MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbInformation, "Winkiss Hinweis:"
            List7.SetFocus
        Else
            
             cLBSatz = List7.list(List7.ListIndex)
            
            If InStr(1, cLBSatz, "offen") = 0 Then
                MsgBox "Bitte nur offene Bonusbeträge auswählen!", vbInformation, "Winkiss Hinweis:"
                List7.SetFocus
            Else
            
               
                lLFNR = Trim(Right(cLBSatz, 8))
                
                
                
'                Zeige_Bonus_Auszahlungen gckundnr
            
                'einlösen + in den Warenkorb legen
            
                gdBonusAusKundenhistorie = ermBonusbetrag(lLFNR)  'CDbl(Trim(Mid(cLBSatz, 17, 7)))
                
                If gdBonusAusKundenhistorie > 0 Then
                    setze_Eingelöst lLFNR
                End If
            
                Unload frmWKL74
            
            End If
            
            
        End If
    
        
    Case 11
        gsHelpstring = "Kundenverkauf"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermBonusbetrag(lNr As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermBonusbetrag = 0
    
     
    sSQL = "Select sum(auszahlbonus) as maxi from Kundenbonus"
    sSQL = sSQL & " where lfnr = " & lNr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermBonusbetrag = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBonusbetrag"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub setze_Eingelöst(lNr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    sSQL = "Update Kundenbonus set"
    sSQL = sSQL & " Sendok = False "
    sSQL = sSQL & " , Eingeloest_Datum = " & CLng(DateValue(Now))
    sSQL = sSQL & " where lfnr = " & lNr
    gdBase.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "setze_Eingelöst"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub drucken(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "KUNDD1", gdBase
    CreateTable "KUNDD1", gdBase
    
    sSQL = "Insert into KUNDD1 select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", FAXNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", MOBILTEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", STADT as ORT "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & " from KUNDEN where KUNDNR = " & gckundnr
    gdBase.Execute sSQL, dbFailOnError
    
   
    loeschNEW "KUNDDRUCK", gdBase
    CreateTable "KUNDDRUCK", gdBase
    
    sSQL = "Insert into KUNDDRUCK select "
    sSQL = sSQL & " ARTNR  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", MENGE  "
    sSQL = sSQL & ", PREIS  "
    sSQL = sSQL & ", ADATE  "
    sSQL = sSQL & ", AZEIT  "
    sSQL = sSQL & ", KUNDNR  "
    sSQL = sSQL & ", FILIALE  "
    sSQL = sSQL & ", KASNUM  "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", AGN  "
    sSQL = sSQL & ", EAN "
    sSQL = sSQL & ", MWST "
    sSQL = sSQL & ", EKPR  "
    sSQL = sSQL & ", VKPR  "
    sSQL = sSQL & ", MOPREIS  "
    sSQL = sSQL & ", BELEGNR  "
    sSQL = sSQL & ", BEST1  "
    sSQL = sSQL & ", RABKENN"
    sSQL = sSQL & ", KK_ART "
    sSQL = sSQL & ", BEDIENER  "
    sSQL = sSQL & ", UMS_OK "
    
    sSQL = sSQL & " from Kassjour where KUNDNR = " & gckundnr
'    sSQL = sSQL & " order by " & sOrder
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into KUNDDRUCK select "
    sSQL = sSQL & " ARTNR  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", MENGE  "
    sSQL = sSQL & ", PREIS  "
    sSQL = sSQL & ", ADATE  "
    sSQL = sSQL & ", AZEIT  "
    sSQL = sSQL & ", KUNDNR  "
    sSQL = sSQL & ", FILIALE  "
    sSQL = sSQL & ", KASNUM  "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", AGN  "
    sSQL = sSQL & ", EAN "
    sSQL = sSQL & ", MWST "
    sSQL = sSQL & ", EKPR  "
    sSQL = sSQL & ", VKPR  "
'    sSQL = sSQL & ", MOPREIS  "
    sSQL = sSQL & ", BELEGNR  "
    sSQL = sSQL & ", BEST1  "
    sSQL = sSQL & ", RABKENN"
    sSQL = sSQL & ", 'KO' as KK_ART "
    sSQL = sSQL & ", BEDIENER  "
'    sSQL = sSQL & ", UMS_OK "
'
    sSQL = sSQL & " from Kollverk where KUNDNR = " & gckundnr
'    sSQL = sSQL & " order by " & sOrder
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KUNDDRUCK_TEMP", gdBase
    
    
    sSQL = "Select * into KUNDDRUCK_TEMP from KUNDDRUCK "
    gdBase.Execute sSQL, dbFailOnError
    

    loeschNEW "KUNDDRUCK", gdBase
    CreateTable "KUNDDRUCK", gdBase
    
    sSQL = "Insert into KUNDDRUCK Select * from KUNDDRUCK_TEMP "
    sSQL = sSQL & " order by " & sOrder
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KUNDDRUCK_TEMP", gdBase
   
    reportbildschirm "", "aWKL74"

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeSummen(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    If cKundnr = "" Then
        Exit Sub
    End If
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    List2.Clear
    sSQL = "select distinct(year(adate)) as jahr,sum(preis) as sumPreis from KUNDAZE where KUNDNR = " & cKundnr
    sSQL = sSQL & " group by year(adate) order by year(adate) desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!jahr) Then
            sSatz = rsrs!jahr & Space(4 - Len(rsrs!jahr))
        End If
        
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(11 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
        Else
            sSatz = sSatz & Space(11)
        End If
        
        List2.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeSummen"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Zeige_Bonus_Auszahlungen(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sSatz As String
    Dim sFeld As String
    
    List7.Clear
    sSQL = "select  "
    sSQL = sSQL & " lfnr, DATUM "
    sSQL = sSQL & ", AUSZAHLBONUS"
    sSQL = sSQL & ", EINGELOEST_DATUM"
    sSQL = sSQL & " from KUNDENBONUS where KUNDNR = " & cKundnr
    sSQL = sSQL & " order by Datum desc "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        sFeld = ""
        
        If Not IsNull(rsrs!Datum) Then
            sFeld = rsrs!Datum
        End If
        
        sSatz = Format(sFeld, "DD.MM.YY")
        sSatz = sSatz & Space(10)
        
        
        If Not IsNull(rsrs!AUSZAHLBONUS) Then
            sFeld = rsrs!AUSZAHLBONUS
        End If
        sSatz = sSatz & Format(sFeld, "####0.00") & " € " & Space(4)
        
        If Not IsNull(rsrs!EINGELOEST_DATUM) Then
            sFeld = rsrs!EINGELOEST_DATUM
            sFeld = "eingelöst am: " & Format(sFeld, "DD.MM.YY")
        Else
            sFeld = "offen"
        End If
        
        sSatz = sSatz & sFeld & Space(140)
        
        
        
        If Not IsNull(rsrs!lfnr) Then
            sFeld = rsrs!lfnr
        End If
        
        sSatz = sSatz & sFeld
        
        
        List7.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    Else
        List7.Visible = False
        Label1(14).Visible = False
        Command3(5).Visible = False
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_Bonus_Auszahlungen"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function sumStückInsgesamt(cKundnr As String) As Long
    On Error GoTo LOKAL_ERROR

    sumStückInsgesamt = 0
    Dim rsrs        As Recordset
    Dim sSQL        As String

    sSQL = "select sum(menge) as sumStück  "
    sSQL = sSQL & " from KUNDAZE "
    sSQL = sSQL & "  where Kundnr = " & cKundnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sumStück) Then
            sumStückInsgesamt = rsrs!sumStück
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "sumStückInsgesamt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Command3_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command3_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    AllesinKUNDAZE gckundnr, False
    
    List1.Clear
    List1.AddItem "Datum     Menge   Artnr  Artikelbezeichnung                  Fil Preis    Bed"
    
    
   'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
    sOrder = "adate desc"
   'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
    Option4(3).Tag = "adate asc"
    ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder
    
    ZeigeSummen gckundnr
    ZeigeSchwerpunktLinr gckundnr, List4
    ZeigeSchwerpunktAGN gckundnr, List5
    ZeigeTermine gckundnr
    
    
    
    
    
    Zeige_Bonus_Auszahlungen gckundnr
    
    Zeigdiesterne Picture1(1), Picture1(2), Picture1(3), Picture1(4), Picture1(5), wievieleSterne(gckundnr, "KUNDAZE", 180), wievieleSterneAlt
    
    Label1(10).Caption = Format(glAnzvkKU, "####0")
    Label1(6).Caption = Format(gdumsgesKU, "####0.00")
    Label1(3).Caption = Format(gdEuroproBonKU, "####0.00")
    
    If glAnzvkKU > 0 Then
        Label1(11).Caption = Format(sumStückInsgesamt(gckundnr) / glAnzvkKU, "####0.00")
    Else
        Label1(11).Caption = "0"
    End If
    
    If gbBonusEinloesungHierErlaubt Then
        Command3(5).Visible = True
    Else
        Command3(5).Visible = False
    End If
    
    anzeige "normal", gckundnr & " " & WhatIsXfromKu(gckundnr, "vorName") & " " & WhatIsXfromKu(gckundnr, "Name"), lblanzeige
    Label2.Caption = lblanzeige.Caption
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(13).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "KUNDD1", gdBase
    loeschNEW "KUNDDRUCK", gdBase
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
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(13).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame5_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    If index = 13 Then
    
        Select Case Label1(13).Caption
        
            Case "+ mehr anzeigen"
            
                Label1(13).Caption = "- weniger anzeigen"
                Label1(13).Refresh
    
                loeschNEW "KUNDAZE", gdBase
                AllesinKUNDAZE gckundnr, True
            
                List1.Clear
                List1.AddItem "Datum     Menge   Artnr  Artikelbezeichnung                  Fil Preis    Bed"
                
'                sOrder = "adate desc"
                
                
                If Option4(0).value = True Then
                    sOrder = Option4(0).Tag
                ElseIf Option4(3).value = True Then
                    sOrder = Option4(3).Tag
                ElseIf Option4(4).value = True Then
                    sOrder = Option4(4).Tag
                ElseIf Option4(5).value = True Then
                    sOrder = Option4(5).Tag
                ElseIf Option4(7).value = True Then
                    sOrder = Option4(7).Tag
                ElseIf Option4(8).value = True Then
                    sOrder = Option4(8).Tag
                End If
                
                ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder
                
                ZeigeSummen gckundnr
                ZeigeSchwerpunktLinr gckundnr, List4
                ZeigeSchwerpunktAGN gckundnr, List5
                ZeigeTermine gckundnr
                
                Zeigdiesterne Picture1(1), Picture1(2), Picture1(3), Picture1(4), Picture1(5), wievieleSterne(gckundnr, "KUNDAZE", 180), wievieleSterneAlt
                
                Label1(10).Caption = Format(glAnzvkKU, "####0")
                Label1(6).Caption = Format(gdumsgesKU, "####0.00")
                Label1(3).Caption = Format(gdEuroproBonKU, "####0.00")
                
            
                If glAnzvkKU > 0 Then
                    Label1(11).Caption = Format(sumStückInsgesamt(gckundnr) / glAnzvkKU, "####0.00")
                Else
                    Label1(11).Caption = "0"
                End If
                
                anzeige "normal", gckundnr & " " & WhatIsXfromKu(gckundnr, "vorName") & " " & WhatIsXfromKu(gckundnr, "Name"), lblanzeige
                Label2.Caption = lblanzeige.Caption
                
            Case "- weniger anzeigen"
            
                Label1(13).Caption = "+ mehr anzeigen"
                Label1(13).Refresh
            
                loeschNEW "KUNDAZE", gdBase
                AllesinKUNDAZE gckundnr, False
            
                List1.Clear
                List1.AddItem "Datum     Menge   Artnr  Artikelbezeichnung                  Fil Preis    Bed"
                
'                sOrder = "adate desc"
                
                
                If Option4(0).value = True Then
                    sOrder = Option4(0).Tag
                ElseIf Option4(3).value = True Then
                    sOrder = Option4(3).Tag
                ElseIf Option4(4).value = True Then
                    sOrder = Option4(4).Tag
                ElseIf Option4(5).value = True Then
                    sOrder = Option4(5).Tag
                ElseIf Option4(7).value = True Then
                    sOrder = Option4(7).Tag
                ElseIf Option4(8).value = True Then
                    sOrder = Option4(8).Tag
                End If
                
                ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder
                
                ZeigeSummen gckundnr
                ZeigeSchwerpunktLinr gckundnr, List4
                ZeigeSchwerpunktAGN gckundnr, List5
                ZeigeTermine gckundnr
                
                Zeigdiesterne Picture1(1), Picture1(2), Picture1(3), Picture1(4), Picture1(5), wievieleSterne(gckundnr, "KUNDAZE", 180), wievieleSterneAlt
                
                Label1(10).Caption = Format(glAnzvkKU, "####0")
                Label1(6).Caption = Format(gdumsgesKU, "####0.00")
                Label1(3).Caption = Format(gdEuroproBonKU, "####0.00")
                
            
                If glAnzvkKU > 0 Then
                    Label1(11).Caption = Format(sumStückInsgesamt(gckundnr) / glAnzvkKU, "####0.00")
                Else
                    Label1(11).Caption = "0"
                End If
                
                anzeige "normal", gckundnr & " " & WhatIsXfromKu(gckundnr, "vorName") & " " & WhatIsXfromKu(gckundnr, "Name"), lblanzeige
                Label2.Caption = lblanzeige.Caption
                
            End Select
    
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ZeigeTermine(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sSatz As String
    Dim sFeld As String
    
    List6.Clear
    sSQL = "select  "
    sSQL = sSQL & " DATUM "
    sSQL = sSQL & ", min(Uhrzeit) as AZEIT "
    sSQL = sSQL & ", BEHANDLUNG"
    sSQL = sSQL & ", BEDNAME"
    sSQL = sSQL & ", buchungsnr"
    
    sSQL = sSQL & " from Termine where KUNDNR = " & cKundnr
    sSQL = sSQL & " group by buchungsnr, Datum ,BEDNAME ,BEHANDLUNG "
    sSQL = sSQL & " order by Datum desc "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        sFeld = ""
        
        If Not IsNull(rsrs!Datum) Then
            sFeld = rsrs!Datum
        End If
        
        sSatz = Format(sFeld, "DD.MM.YY")
        sSatz = sSatz & " "
        
        If Not IsNull(rsrs!AZEIT) Then
            sFeld = rsrs!AZEIT
        End If
        
        sSatz = sSatz & Format(sFeld, "HH:MM")
        sSatz = sSatz & " "
        
        
        If Not IsNull(rsrs!Behandlung) Then
            sFeld = rsrs!Behandlung
        End If
        
        sFeld = SwapStr(sFeld, Chr(13), " ")
        sFeld = SwapStr(sFeld, Chr(10), " ")
        sFeld = SwapStr(sFeld, "  ", " ")
        
        If Len(sFeld) > 30 Then
            sSatz = sSatz & Left(sFeld, 27) & "... "
        Else
            sSatz = sSatz & sFeld & Space(31 - Len(sFeld))
        End If
        
        If Not IsNull(rsrs!bedname) Then
            sFeld = rsrs!bedname
        End If
        
        If Len(sFeld) > 20 Then
            sSatz = sSatz & Left(sFeld, 17) & "... "
        Else
            sSatz = sSatz & sFeld & Space(21 - Len(sFeld))
        End If
        
        If Not IsNull(rsrs!BUCHUNGSNR) Then
            sFeld = rsrs!BUCHUNGSNR
        End If
        
        sSatz = sSatz & Space(100) & sFeld
        
        
        List6.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    Else
        List6.Visible = False
        Label1(4).Visible = False
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeTermine"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If index = 13 Then
        Label1(13).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List6_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sBuchungsNr As String
    Dim sSQL As String
    
    
    Select Case KeyCode
        Case Is = 46    'Del
        
            iRet = MsgBox("Möchten Sie diesen Termin wirklich löschen?", vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            
            If iRet = vbYes Then
                If Not List6.ListIndex = -1 Then
                
                
                    cLBSatz = List6.list(List6.ListIndex)
                    sBuchungsNr = Trim(Right(cLBSatz, 6))
                    
                    sSQL = "Delete * from Termine where buchungsnr = " & sBuchungsNr & " "
                    gdBase.Execute sSQL, dbFailOnError
                
                
                    List6.RemoveItem (List6.ListIndex)
                End If
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List6_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option4_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo LOKAL_ERROR
 
    sOrder = Option4(index).Tag

     ' Odayy  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
  
        If sOrder = "adate asc" Then
            Option4(index).Tag = "adate desc"
        ElseIf sOrder = "Bednr asc" Then
            Option4(index).Tag = "Bednr desc"
        ElseIf sOrder = "Filiale asc" Then
            Option4(index).Tag = "Filiale desc"
        ElseIf sOrder = "menge asc" Then
            Option4(index).Tag = "menge desc"
        ElseIf sOrder = "Filiale , adate asc" Then
            Option4(index).Tag = "Filiale , adate desc"
        ElseIf sOrder = "bezeich asc" Then
            Option4(index).Tag = "bezeich desc"
       
        ElseIf sOrder = "adate desc" Then
            Option4(index).Tag = "adate asc"
        ElseIf sOrder = "Bednr desc" Then
            Option4(index).Tag = "Bednr asc"
        ElseIf sOrder = "Filiale desc" Then
            Option4(index).Tag = "Filiale asc"
        ElseIf sOrder = "menge desc" Then
            Option4(index).Tag = "menge asc"
        ElseIf sOrder = "Filiale , adate desc" Then
            Option4(index).Tag = "Filiale , adate asc"
        ElseIf sOrder = "bezeich desc" Then
            Option4(index).Tag = "bezeich asc"
        End If
        
    ' Odayy  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
 
    ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Picture1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    If IsNumeric(gckundnr) Then
        frmWKL157.Show 1
        'hier die gckundnr nicht löschen
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
