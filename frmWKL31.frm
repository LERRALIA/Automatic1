VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL31 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Etiketten eigene Auswahl"
   ClientHeight    =   8625
   ClientLeft      =   3855
   ClientTop       =   2040
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
   Icon            =   "frmWKL31.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboPreisAenderungen 
      Height          =   330
      Left            =   9600
      TabIndex        =   60
      Text            =   "Combo1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   50
      Top             =   720
      Width           =   2055
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
      Caption         =   "MDE auslesen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
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
      Caption         =   "sofort Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   41
      Top             =   3480
      Width           =   2055
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
      Caption         =   "sofort Drucken Netto"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9600
      TabIndex        =   40
      Top             =   4080
      Width           =   2055
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
      Caption         =   "sofort Drucken Brutto"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame3 
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
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   6720
      Width           =   12015
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C000&
         Caption         =   "mit Sofortetikett"
         Height          =   255
         Left            =   2760
         TabIndex        =   59
         Top             =   120
         Width           =   2535
      End
      Begin VB.ComboBox cboStrichEndlos 
         Height          =   330
         Left            =   2760
         TabIndex        =   58
         Text            =   "Combo1"
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   35
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "Schlieﬂen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   6120
         TabIndex        =   34
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "zum Etikettendruck"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
   End
   Begin VB.Frame Frame0 
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
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   7320
      Width           =   12015
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   2
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   3
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   4
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   5
         Left            =   3720
         TabIndex        =   25
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   6
         Left            =   4440
         TabIndex        =   24
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   7
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   8
         Left            =   5880
         TabIndex        =   22
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   9
         Left            =   6600
         TabIndex        =   21
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   10
         Left            =   7320
         TabIndex        =   20
         Top             =   240
         Width           =   720
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
         Caption         =   "<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   11
         Left            =   8040
         TabIndex        =   19
         Top             =   240
         Width           =   720
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
         Caption         =   "C"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   12
         Left            =   9480
         TabIndex        =   18
         Top             =   240
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   13
         Left            =   10200
         TabIndex        =   17
         Top             =   240
         Width           =   720
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
         Caption         =   "<<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   14
         Left            =   10920
         TabIndex        =   16
         Top             =   240
         Width           =   720
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
         Caption         =   ">>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   15
         Left            =   8760
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   720
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
         Caption         =   "."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
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
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   9135
      Begin VB.CheckBox Check4 
         Caption         =   "Original-EAN"
         Height          =   255
         Left            =   2400
         TabIndex        =   56
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "schneller Scanmodus"
         Height          =   255
         Left            =   2400
         TabIndex        =   55
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "halten"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   2400
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3000
         Width           =   2655
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   7
         Left            =   840
         TabIndex        =   52
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
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
         Caption         =   "Ean aus"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         Caption         =   "nur nichtrabattierf‰hige Artikel"
         Height          =   255
         Left            =   1560
         TabIndex        =   49
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   7800
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picprogress 
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   1875
         TabIndex        =   44
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ListBox List3 
         Height          =   900
         ItemData        =   "frmWKL31.frx":0442
         Left            =   2400
         List            =   "frmWKL31.frx":0444
         MultiSelect     =   2  'Erweitert
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   2400
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorgabe Filiale"
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
         Left            =   5400
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "Filialnr.:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorgabe Etiketten-Anzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   3360
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "Anzahl Etiketten je Artikel:"
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
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "gem‰ﬂ Vorgabe"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   3960
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "nach Bestand"
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
         Left            =   2400
         TabIndex        =   10
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2400
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "LibesNr/NAN:"
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
         Left            =   480
         TabIndex        =   53
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   48
         Top             =   2760
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   47
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Anzahl Artikel:"
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
         Index           =   6
         Left            =   0
         TabIndex        =   46
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Linie (F2):"
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
         Left            =   960
         TabIndex        =   39
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Anzahl Etiketten:"
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
         Left            =   0
         TabIndex        =   9
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "ArtNr. / EAN  :"
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
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant (F2):"
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
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   8
      Left            =   9600
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
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
      Caption         =   "Spezialetikett"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   9
      Left            =   9600
      TabIndex        =   61
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Caption         =   "Etiketten drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Preis‰nderungen:"
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
      Index           =   11
      Left            =   9600
      TabIndex        =   62
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "MDE Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   9600
      TabIndex        =   51
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Etikettendruck nach eigener Auswahl"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
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
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmWKL31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bBlau As Boolean
Dim mdeErr As Boolean
Dim bFocusLibesnr As Boolean
Dim isEtidruFree        As Boolean
Private Function fnPruefeEingabeDialogWKL31%()
    On Error GoTo LOKAL_ERROR

    Dim cFeld       As String

    fnPruefeEingabeDialogWKL31% = 0
    
    If Check1.Value = vbChecked Then
        Exit Function
    End If
    
    If Check2.Value = vbChecked And Text1(1).Text = "" And Text1(4).Text = "" And Text1(5).Text = "" Then
        fnPruefeEingabeDialogWKL31% = 4
    End If

    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    If cFeld = "" Then
        cFeld = Text1(1).Text
        cFeld = Trim$(cFeld)
        If cFeld = "" Then
            fnPruefeEingabeDialogWKL31% = 1
            Exit Function
        ElseIf IsNumeric(cFeld) = False Then
            fnPruefeEingabeDialogWKL31% = 3
            Exit Function
        End If
    End If

    If Text1(2).Visible = True Then
        
        cFeld = Text1(2).Text
        cFeld = Trim$(cFeld)
        If cFeld = "" Then
            fnPruefeEingabeDialogWKL31% = 2
        End If
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWKL31%"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub SchreibeEtiDruKurz()
    On Error GoTo LOKAL_ERROR
        
    Dim lartnr      As Long
    Dim cSQL        As String
    Dim cSQL1       As String
    Dim cLinr       As String
    Dim cLiBesNr    As String
    Dim clpz        As String
    Dim cArtNr      As String
    Dim cAnzahl     As String
    Dim bAnd        As Boolean
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    bAnd = False
    
    cLiBesNr = Text1(5).Text
    cLiBesNr = Trim$(cLiBesNr)
    
    cLinr = Text1(0).Text
    cLinr = Trim$(cLinr)
    
    If cLinr = "400001" Then
        If cLiBesNr <> "" Then
            cLiBesNr = Left(cLiBesNr, Len(cLiBesNr) - 1)
            bFocusLibesnr = True
        End If
    End If
    

    clpz = Text1(4).Text
    clpz = Trim$(clpz)
    
    cArtNr = Text1(1).Text
    cArtNr = Trim$(cArtNr)
    
    If cArtNr = "" Then
        cSQL1 = " ARTIKEL.GEFUEHRT = 'J' "
        bAnd = True
    Else
        cSQL1 = ""
    End If
    
    cAnzahl = Val(Trim(Text1(2).Text))
    

    cSQL = "DELETE FROM ETI" & srechnertab & ""
    gdBase.Execute cSQL, dbFailOnError
 
    cSQL = " Select distinct(ARTIKEL.ARTNR) "
    cSQL = cSQL & ", ARTIKEL.BEZEICH "
    cSQL = cSQL & ", ARTIKEL.KVKPR1 as VKPR "
    cSQL = cSQL & ", ARTIKEL.BESTAND "
    cSQL = cSQL & ", ARTIKEL.SYNSTATUS "
    
    If Option1(0).Value = True Then 'nach bestand
        cSQL = cSQL & ", ARTIKEL.BESTAND as Anzahl "
    ElseIf Option1(1).Value = True Then 'nach vorgabe
        cSQL = cSQL & ", " & cAnzahl & " as Anzahl "
    End If
    
    cSQL = cSQL & ", ARTIKEL.LIBESNR "
    cSQL = cSQL & ", ARTIKEL.EAN "
    cSQL = cSQL & ", ARTIKEL.LPZ "
    cSQL = cSQL & ", ARTIKEL.LINR "
    cSQL = cSQL & ", " & gcFilNr & " as Filnr "
    cSQL = cSQL & " from ARTIKEL inner join ARTLIEF on "
    cSQL = cSQL & " ARTIKEL.ARTNR = ARTLIEF.ARTNR where "
    
    cSQL = cSQL & cSQL1
    
    If Check1.Value = vbChecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " ARTIKEL.RABATT_OK = 'N' "
        bAnd = True
    End If
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If

        cSQL = cSQL & " ARTLIEF.LINR = " & cLinr & " "
        bAnd = True
    End If
    
    If cLiBesNr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If


        cSQL = cSQL & " ARTLIEF.Libesnr = '" & cLiBesNr & "' "
        bAnd = True
    End If
    
    If List3.ListCount <> 0 Then
        If bAnd Then
                cSQL = cSQL & " and "
            End If
    
        cSQL = cSQL & "  (ARTIKEL.LPZ = " & Mid(List3.list(0), 1, InStr(1, List3.list(0), " ")) & " "
        For lcount = 1 To List3.ListCount - 1
            cSQL = cSQL & " or ARTIKEL.LPZ = " & Mid(List3.list(lcount), 1, InStr(1, List3.list(lcount), " ")) & " "
        Next lcount
        cSQL = cSQL & ")"
    Else
        If clpz <> "" Then
            If bAnd Then
                cSQL = cSQL & "and "
            End If
            cSQL = cSQL & "ARTIKEL.LPZ = " & clpz & " "
            bAnd = True
        End If
    End If
    
    If cArtNr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        Select Case Len(cArtNr)
            Case Is > 8
            
                If Ist_in_ARTEAN_K(cArtNr) Then
                
                End If
            
            
            
            
                cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cArtNr & "' "
'                cSQL = cSQL & "or Artikel.Artnr in (Select artnr from artean_k where EAN = '" & cArtNr & "') "
            Case Is = 8
                If Left(cArtNr, 1) = "2" Or Left(cArtNr, 1) = "0" Then
                
                    If Check4.Value = vbChecked Then
                    
                        If Ist_in_ARTEAN_K(cArtNr) Then
                
                        End If
                        
                        cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                        cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cArtNr & "' "
                        cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cArtNr & "' "
'                        cSQL = cSQL & "or Artikel.Artnr in (Select artnr from artean_k where EAN = '" & cArtNr & "') "
                    Else
                        cArtNr = Mid(cArtNr, 2, 6)
                        cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
                    End If
                Else
                
                    If Ist_in_ARTEAN_K(cArtNr) Then
                
                    End If
                    
                    cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                    cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cArtNr & "' "
                    cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cArtNr & "' "
'                    cSQL = cSQL & "or Artikel.Artnr in (Select artnr from artean_k where EAN = '" & cArtNr & "') "
                End If
            Case Is = 6
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
            Case Else
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
        End Select
    End If
    
    Dim rsEti As Recordset
    Dim siAnzeige As Single
    
    siAnzeige = 0
    
    Set rsEti = gdBase.OpenRecordset("ETI" & srechnertab)
'    MsgBox cSQL
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            siAnzeige = siAnzeige + 1
            Label2(7).Caption = siAnzeige
            Label2(7).Refresh
            
            rsEti.AddNew
    
            If Not IsNull(rsrs!artnr) Then
                rsEti!artnr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                rsEti!BEZEICH = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                rsEti!vkpr = rsrs!vkpr
            End If
            
            If Not IsNull(rsrs!BESTAND) Then
                rsEti!BESTAND = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!SYNStatus) Then
                rsEti!SYNStatus = rsrs!SYNStatus
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                rsEti!ANZAHL = rsrs!ANZAHL
            End If
            
            If Not IsNull(rsrs!LIBESNR) Then
                rsEti!LIBESNR = rsrs!LIBESNR
            End If
            
            If Not IsNull(rsrs!EAN) Then
                rsEti!EAN = rsrs!EAN
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                rsEti!LPZ = rsrs!LPZ
            End If
            
            If Not IsNull(rsrs!linr) Then
                rsEti!linr = rsrs!linr
            End If
            
            If Not IsNull(rsrs!filnr) Then
                rsEti!filnr = rsrs!filnr
            End If
            
            rsEti.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsEti.Close: Set rsEti = Nothing

    '************************** ab hier gemeinsame Verarbeitung ****************
    
    cSQL = "DELETE FROM ETI" & srechnertab & " where SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "DELETE FROM ETI" & srechnertab & " where anzahl < 1"
    gdBase.Execute cSQL, dbFailOnError
    

    Dim lAnzahl As Long
    Dim lFil As Long
    
    If Text1(3).Text = "" Then
        lFil = CLng(gcFilNr)
    Else
        If IsNumeric(Text1(3).Text) Then
            lFil = Val(Text1(3).Text)
        End If
    End If
    
    
    
    Set rsrs = gdBase.OpenRecordset("ETI" & srechnertab, dbOpenTable)
    
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        siAnzeige = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        siAnzeige = siAnzeige - 1
        Label2(7).Caption = siAnzeige
        Label2(7).Refresh
        
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
            
            If Not IsNull(rsrs!ANZAHL) Then
                lAnzahl = rsrs!ANZAHL
                
                schreibeWKEtidru cArtNr, lAnzahl, lFil
            End If
    
        End If
        
        rsrs.MoveNext
        Loop
        If bFocusLibesnr = True Then
            'libesnr soll leer
            Text1(5).Text = ""
        End If
    Else
        
        If Option1(0).Value = True Then
            MsgBox "Keine Artikelbest‰nde zum Speichern gefunden!", vbInformation, "Winkiss Hinweis:"
            If bFocusLibesnr = True Then
            
                'libesnr soll gef¸llt bleiben
            
            Else
                Text1(5).Text = ""
            
            End If
            
        ElseIf Option1(1).Value = True Then
        
        
            MsgBox "Keine Artikel zum Speichern gefunden!(gelˆscht oder nicht gef¸hrt?)", vbInformation, "Winkiss Hinweis:"
            If bFocusLibesnr = True Then
                'libesnr soll gef¸llt bleiben
            Else
                Text1(5).Text = ""
            
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeEtiDruKurz"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub SchreibeEtiforSofort(srepname As String)
    On Error GoTo LOKAL_ERROR
        
    Dim lartnr      As Long
    Dim cSQL        As String
    Dim cSQL1       As String
    Dim cLinr       As String
    Dim clpz        As String
    Dim cArtNr      As String
    Dim cAnzahl     As String
    Dim bAnd        As Boolean
    ReDim acArtNr(0 To 0) As String
    ReDim acAnzEti(0 To 0) As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    bAnd = False
    
    cLinr = Text1(0).Text
    cLinr = Trim$(cLinr)
    
    clpz = Text1(4).Text
    clpz = Trim$(clpz)
    
    cArtNr = Text1(1).Text
    cArtNr = Trim$(cArtNr)
    If Not IsNumeric(cArtNr) Then
        If Not IsNumeric(cLinr) Then
            Exit Sub
        Else
        
        End If
    End If
    
    If cArtNr = "" Then
        cSQL1 = " ARTIKEL.GEFUEHRT = 'J' "
        bAnd = True
    Else
        cSQL1 = ""
    End If
    
    cAnzahl = Text1(2).Text
    cAnzahl = Trim$(cAnzahl)
    If cAnzahl = "" Then cAnzahl = "1"
    
    cSQL = "DELETE FROM ETI" & srechnertab
    gdBase.Execute cSQL, dbFailOnError

    cSQL = " Select "
    cSQL = cSQL & "  ARTIKEL.ARTNR "
    cSQL = cSQL & ", ARTIKEL.BEZEICH "
    cSQL = cSQL & ", ARTIKEL.KVKPR1 as VKPR "
    cSQL = cSQL & ", ARTIKEL.BESTAND "
    cSQL = cSQL & ", ARTIKEL.SYNSTATUS "
    
    If Option1(0).Value = True Then 'nach bestand
        cSQL = cSQL & ", ARTIKEL.BESTAND as Anzahl "
    ElseIf Option1(1).Value = True Then 'nach vorgabe
        cSQL = cSQL & ", " & cAnzahl & " as Anzahl "
    End If
    
    cSQL = cSQL & ", ARTIKEL.LIBESNR "
    cSQL = cSQL & ", ARTIKEL.EAN "
    cSQL = cSQL & ", ARTIKEL.LPZ "
    cSQL = cSQL & ", ARTIKEL.LINR "
    cSQL = cSQL & ", " & gcFilNr & " as Filnr "
    cSQL = cSQL & " from ARTIKEL inner join ARTLIEF on "
    cSQL = cSQL & " ARTIKEL.ARTNR = ARTLIEF.ARTNR and Artikel.LINR = Artlief.LINR where "
    
    cSQL = cSQL & cSQL1
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If

        cSQL = cSQL & "ARTLIEF.LINR = " & cLinr & " "
        bAnd = True
    End If
    
    If List3.ListCount <> 0 Then
    
        If bAnd Then
                cSQL = cSQL & " and "
            End If
    
        cSQL = cSQL & "  (ARTIKEL.LPZ = " & Mid(List3.list(0), 1, InStr(1, List3.list(0), " ")) & " "
        For lcount = 1 To List3.ListCount - 1
            cSQL = cSQL & " or ARTIKEL.LPZ = " & Mid(List3.list(lcount), 1, InStr(1, List3.list(lcount), " ")) & " "
        Next lcount
        cSQL = cSQL & ")"
    Else
    
        If clpz <> "" Then
            If bAnd Then
                cSQL = cSQL & "and "
            End If
            cSQL = cSQL & "ARTIKEL.LPZ = " & clpz & " "
            bAnd = True
        End If
        
    End If
    
    If cArtNr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        Select Case Len(cArtNr)
            Case Is > 8
                cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cArtNr & "' "
            Case Is = 8
                If Left(cArtNr, 1) = "2" Or Left(cArtNr, 1) = "0" Then
                    cArtNr = Mid(cArtNr, 2, 6)
                    cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
                Else
                    cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                End If
            Case Is = 6
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
            Case Else
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
        End Select
    End If
    
    Dim rsEti As Recordset
    Dim siAnzeige As Single
    
    siAnzeige = 0
    
    Set rsEti = gdBase.OpenRecordset("ETI" & srechnertab)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            siAnzeige = siAnzeige + 1
            Label2(7).Caption = siAnzeige
            Label2(7).Refresh

            rsEti.AddNew
            
            If Not IsNull(rsrs!artnr) Then
                rsEti!artnr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                rsEti!BEZEICH = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                rsEti!vkpr = rsrs!vkpr
            End If
            
            If Not IsNull(rsrs!BESTAND) Then
                rsEti!BESTAND = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!SYNStatus) Then
                rsEti!SYNStatus = rsrs!SYNStatus
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                rsEti!ANZAHL = rsrs!ANZAHL
            End If
            
            If Not IsNull(rsrs!LIBESNR) Then
                rsEti!LIBESNR = rsrs!LIBESNR
            End If
            
            If Not IsNull(rsrs!EAN) Then
                rsEti!EAN = rsrs!EAN
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                rsEti!LPZ = rsrs!LPZ
            End If
            
            If Not IsNull(rsrs!linr) Then
                rsEti!linr = rsrs!linr
            End If
            
            If Not IsNull(rsrs!filnr) Then
                rsEti!filnr = rsrs!filnr
            End If
            
            rsEti.Update
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    rsEti.Close: Set rsEti = Nothing
    

    '************************** ab hier gemeinsame Verarbeitung ****************
    
    cSQL = "DELETE FROM ETI" & srechnertab & " where SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "DELETE FROM ETI" & srechnertab & " where anzahl < 1"
    gdBase.Execute cSQL, dbFailOnError
    
    
    Dim lAnzahl As Long
    Dim lFil As Long
    
    lAnzahl = -1
    
    If Text1(3).Text = "" Then
        lFil = CLng(gcFilNr)
    Else
        If IsNumeric(Text1(3).Text) Then
            lFil = Val(Text1(3).Text)
        End If
    End If
    
    Set rsrs = gdBase.OpenRecordset("ETI" & srechnertab, dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
            cAnzahl = rsrs!ANZAHL
            lAnzahl = lAnzahl + 1
            ReDim Preserve acArtNr(0 To lAnzahl) As String
            ReDim Preserve acAnzEti(0 To lAnzahl) As String
            acArtNr(lAnzahl) = cArtNr
            acAnzEti(lAnzahl) = cAnzahl
            
        End If
        
        rsrs.MoveNext
        Loop
    Else
        
        MsgBox "Keine Artikel bzw. Artikelbest‰nde zum Speichern gefunden!", vbInformation, "INFO"
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If UCase(srepname) = "AWKL30XS" Then
        DruckeGrundPreisEtikettenWKL30kleinspezial acArtNr(), lAnzahl, srepname
    ElseIf UCase(srepname) = "AWKL30YS" Then
        DruckeStrichcodeY acArtNr(), lAnzahl, acAnzEti()
        reportbildschirmToPrinterETI "aWKL30ys", gcEtikettenDrucker, True
    ElseIf UCase(srepname) = "AWKL30ZS" Then
        DruckeStrichcodeY acArtNr(), lAnzahl, acAnzEti()
        reportbildschirmToPrinterETI "AWKL30ZS", gcEtikettenDrucker, True
    Else
        If srepname = "aWOKINE" Then
            DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "NETTO"

        Else
            DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "BRUTTO"

        End If

    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeEtiforSofort"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    Resume Next
End Sub

Private Sub Check5_Click()
On Error GoTo LOKAL_ERROR
    
    If Check5.Value = vbChecked Then
        cboStrichEndlos.Visible = True
        cboStrichEndlos.Refresh
        setzedrucker gcEtikettenDrucker
        
        Command1(0).Caption = "Drucken"
        
        Frame2.Visible = False
        Label2(6).Visible = False
        Label2(7).Visible = False
        Label2(0).Visible = False
        Text1(0).Visible = False
        Label2(10).Visible = False
        Text1(5).Visible = False
        Label2(5).Visible = False
        Text1(4).Visible = False
        Check2.Visible = False
        Check1.Visible = False
        Check3.Visible = False
        Check4.Visible = False
        Label2(2).Visible = False
        Option1(0).Visible = False
        Option1(1).Visible = False
        
        
    Else
        cboStrichEndlos.Visible = False
        cboStrichEndlos.Refresh
        setzedrucker gcListenDrucker
        
        Command1(0).Caption = "Speichern"
        
        Frame2.Visible = True
        Label2(6).Visible = True
        Label2(7).Visible = True
        Label2(0).Visible = True
        Text1(0).Visible = True
        Label2(10).Visible = True
        Text1(5).Visible = True
        Label2(5).Visible = True
        Text1(4).Visible = True
        Check2.Visible = True
        Check3.Visible = True
        Check1.Visible = True
        Check4.Visible = True
        Label2(2).Visible = True
        Option1(0).Visible = True
        Option1(1).Visible = True
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 To 9
            If bBlau Then
            Text1(Label0.Caption).Text = ""
            bBlau = False
            End If
            Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & Command0(Index).Caption
            Text1(Label0.Caption).SetFocus
        Case Is = 10
            If Len(Text1(Label0.Caption).Text) > 0 Then
                Text1(Label0.Caption).Text = Left(Text1(Label0.Caption).Text, Len(Text1(Label0.Caption).Text) - 1)
                Text1(Label0.Caption).SetFocus
            End If
        Case Is = 11
            Text1(Label0.Caption).Text = ""
            Text1(Label0.Caption).SetFocus
        Case Is = 12
            Text1_KeyUp Val(Label0.Caption), vbKeyF2, 0
            
        Case Is = 13
            If Val(Label0.Caption) > 0 Then
                Label0.Caption = Trim$(Str$((Val(Label0.Caption) - 1)))
            End If
            Text1(Label0.Caption).SetFocus
            
        Case Is = 14
            If Val(Label0.Caption) < 2 Then
                Label0.Caption = Trim$(Str$((Val(Label0.Caption) + 1)))
                If Text1(2).Visible = False And Val(Label0.Caption) = 2 Then
                    Label0.Caption = Trim$(Str$((Val(Label0.Caption) - 1)))
                End If
            End If
            Text1(Label0.Caption).SetFocus
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    Dim acArtNr(1) As String
    Dim acAnzEti(1) As String
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
        
            If Command1(0).Caption = "Speichern" Then
        
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiDruKurz
                        leerStandard
                        
                        If gbEtiFokEan Then
                            Text1(1).SetFocus
                        Else
                            Text1(5).SetFocus
                        End If
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbInformation, "Winkiss Hinweis:"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbInformation, "Winkiss Hinweis:"
                            Text1(1).SetFocus
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
                
            ElseIf Command1(0).Caption = "Drucken" Then
            
                If isEtidruFree And Check5.Value = vbChecked Then
                    isEtidruFree = False
                    
                    Dim sArtnr As String
                    sArtnr = Artnr_is(Text1(1).Text)
                    
                    If sArtnr = "" Then
                        Screen.MousePointer = 0
                        MsgBox "Bitte Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                        
                        isEtidruFree = True
                        Text1(1).SetFocus
                        Exit Sub
                    
                    End If
                    
                    acArtNr(0) = sArtnr
                    acAnzEti(0) = "1"
                
                    Select Case cboStrichEndlos.Text
                        Case "69 x 14 (Var 1)" 'Schmucketikett 69x14 Variante 1
                            DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL311a", gcEtikettenDrucker, False
                                       
                        Case "69 x 14 (Var 2)"  'Schmucketikett 69x14 Variante 2
                            DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL311b", gcEtikettenDrucker, False
                            
                        Case "40 x 18 (Var 1)"  'Etikett 40x18 Variante 1
                            DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312a", gcEtikettenDrucker, False
                        
                        Case "40 x 18 (Var 2)"  'Etikett 40x18 Variante 2
                            DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312b", gcEtikettenDrucker, False
                            
                        Case "40 x 18 (Var 3)"  'Etikett 40x18 Variante 3
                            DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312c", gcEtikettenDrucker, False
                            
                        Case "40 x 18 (Var 4)"  'Etikett 40x18 Variante 4
                            DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312d", gcEtikettenDrucker, False
                            
                        Case "45 x 23 (Var 1)"  'Etikett 45x23 Variante 1
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL313a", gcEtikettenDrucker, False
                            
                        Case "45 x 23 (Var 2)"  'Etikett 45x23 Variante 2
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL313b", gcEtikettenDrucker, False
                            
                        Case "69 x 14 (Var 3)"  'Schmucketikett 69x14 Variante 3
                            DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL311c", gcEtikettenDrucker, False
                            
                        Case "45 x 23 (Var 3)"  'Etikett 45x23 Variante 3
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL313c", gcEtikettenDrucker, False
                            
                        Case "38 x 23 (Var 1)"  'Etikett 38x23 Variante 1
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL314a", gcEtikettenDrucker, False
                            
                        Case "38 x 23 (Var 2)"  'Etikett 38x23 Variante 2
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL314b", gcEtikettenDrucker, False
                            
                        Case "38 x 23 (Var 3)"  'Etikett 38x23 Variante 3
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL314c", gcEtikettenDrucker, False
                            
                        Case "51 x 19 (Var 1)"  'Etikett 51x19 Variante 1
                            DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL315a", gcEtikettenDrucker, False
                            
                        Case "51 x 19 (Var 2)"  'Etikett 51x19 Variante 2
                            DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL315b", gcEtikettenDrucker, False
                            
                        Case "49 x 19 (Var 1)"  'Etikett 49x19 Variante 1
                            DruckeEtikett49x19Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL316a", gcEtikettenDrucker, False
                            
                        Case "44 x 21 (Var 1)"  'Etikett 44x21 Variante 1
                            DruckeEtikett44x21Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL317a", gcEtikettenDrucker, False
                            
                        Case "51 x 19 (Var 3)"  'Etikett 51x19 Variante 3
                            DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL315c", gcEtikettenDrucker, False
                            
                        Case "30 x 15 (Var 1)"  'Etikett 30x15 Variante 1
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL3015a", gcEtikettenDrucker, False
                            
                        Case "30 x 15 (Var 2)"  'Etikett 30x15 Variante 2
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL3015b", gcEtikettenDrucker, False
                            
                        Case "30 x 15 (Var 3)"  'Etikett 30x15 Variante 3
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL3015c", gcEtikettenDrucker, False
                            
                        Case "48 x 18 (Var 1)"  'Etikett 48x18 Variante 1
                            DruckeEtikett48x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL319a", gcEtikettenDrucker, False
                            
                        Case "45 x 23 (Var 4)"  'Etikett 45x23 Variante 4
                            DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL313d", gcEtikettenDrucker, False
                           
                        Case "40 x 18 (Var 5)"  'Etikett 40x18 Variante 5
                            DruckeEtikett40x18Variante5 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312e", gcEtikettenDrucker, False
                            
                        Case "40 x 18 (Var 6)"  'Etikett 40x18 Variante 6
                            DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL312f", gcEtikettenDrucker, False
                            
                        Case "35 x 15 (Var 1)" 'Etikett 35x15 Variante 1
                            DruckeEtikett35x15Variante1 acArtNr(), 0, acAnzEti()
                            reportbildschirmToPrinterETI "aWKL322a", gcEtikettenDrucker, False
                            
                    End Select
                    
'                    anzeige "erfolg", "", Label5
                    isEtidruFree = True
                    
                    Text1(1).Text = ""
                    Text1(1).SetFocus
            
                End If
            
            End If
            
        Case Is = 1
            AktionAustragen "Etiketten w‰hlen"
            frmWKL30.Show 1
        Case Is = 2
            voreinstellungspeichernE31
            Unload frmWKL31
        Case 3
            If Modul6.FindFile(App.Path, "aWOKIBR.rpt") Then
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiforSofort "aWOKIBR"
                        
                        leerStandard
                        If bFocusLibesnr Then
                            Text1(5).SetFocus
                        Else
                            Text1(1).SetFocus
                        End If
                        
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbCritical, "STOP!"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbCritical, "STOP!"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbCritical, "STOP!"
                            Text1(1).SetFocus
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                            
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
            End If
        Case 4
            If Modul6.FindFile(App.Path, "aWOKINE.rpt") Then
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiforSofort "aWOKINE"
                        
                        leerStandard
                        If bFocusLibesnr Then
                            Text1(5).SetFocus
                        Else
                            Text1(1).SetFocus
                        End If
                        
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbCritical, "STOP!"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbCritical, "STOP!"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbCritical, "STOP!"
                            Text1(1).SetFocus
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                            
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
            End If
        Case 5
            If Modul6.FindFile(App.Path, "aWKL30xs.rpt") Then
            
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiforSofort "aWKL30xs"
                        
                        leerStandard
                        If bFocusLibesnr Then
                            Text1(5).SetFocus
                        Else
                            Text1(1).SetFocus
                        End If
                        
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbCritical, "STOP!"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbCritical, "STOP!"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbCritical, "STOP!"
                            Text1(1).SetFocus
                            
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                            
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
            End If
            
            
            If Modul6.FindFile(gcDBPfad, "aWKL30ys.rpt") Then
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
    
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiforSofort "aWKL30ys"
                        
                        leerStandard
                        If bFocusLibesnr Then
                            Text1(5).SetFocus
                        Else
                            Text1(1).SetFocus
                        End If
                        
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbInformation, "Winkiss Hinweis:"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbInformation, "Winkiss Hinweis:"
                            Text1(1).SetFocus
                            
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                            
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
            End If
        Case 6
        
            loeschNEW "ARTERRIN", gdBase
            CreateTable "ARTERRIN", gdBase
        
            MDElesen
            If mdeErr Then
                anzeigeNew "normal", "nicht erkannte Artikel werden angezeigt...", Label2(9)
                reportbildschirm "", "aWKL46e" 'Error artikel mde
            End If
        Case 7

            
            Dim sEANcodi As String
            sEANcodi = ean13(Text1(1).Text)
            MsgBox sEANcodi
            
            sSQL = "Insert into eancode (eancodeWK) values ( '" & sEANcodi & "')"
            gdBase.Execute sSQL, dbFailOnError
            
        Case 8
            
            If Modul6.FindFile(App.Path, "aWKL30zs.rpt") Then
            
                iRet = fnPruefeEingabeDialogWKL31%()
                If iRet = 0 Then
                    If IsAktionZulaessig("Etiketten w‰hlen") Then
                        SchreibeEtiforSofort "aWKL30zs"
                        
                        leerStandard
                        If bFocusLibesnr Then
                            Text1(5).SetFocus
                        Else
                            Text1(1).SetFocus
                        End If
                        
                        AktionAustragen "Etiketten w‰hlen"
                    Else
                        Exit Sub
                    End If
                Else
                    Select Case iRet
                        Case Is = 1
                            MsgBox "Bitte Lieferanten oder Artikelnummer/EAN angeben!", vbCritical, "STOP!"
                            Text1(0).SetFocus
                        Case Is = 2
                            MsgBox "Bitte die zu druckende St¸ckzahl angeben!", vbCritical, "STOP!"
                            Text1(2).SetFocus
                        Case Is = 3
                            MsgBox "Artikelnummer bzw. EAN ung¸ltig!", vbCritical, "STOP!"
                            Text1(1).SetFocus
                            
                        Case Is = 4
                            MsgBox "Bitte LibesNr/NAN oder Artikelnummer/EAN angeben!", vbInformation, "Winkiss Hinweis:"
                            If gbEtiFokEan Then
                                Text1(1).SetFocus
                            Else
                            
                                Text1(5).SetFocus
                            End If
                    End Select
                End If
            End If
        Case 9
            'Etiketten abstellen
            

            If cboPreisAenderungen.Text <> "" Then

                Screen.MousePointer = 11
                
                loeschNEW "LSTEETI", gdBase
                CreateTableT2 "LSTEETI", gdBase
                
                sSQL = "Insert into LSTEETI select Artnr "
                sSQL = sSQL & ", BEZEICH "
                sSQL = sSQL & ", BESTAND "
                sSQL = sSQL & ", ANZAHL "
                sSQL = sSQL & ", vkprneu as VKPR "
                sSQL = sSQL & ", LIBESNR "
                sSQL = sSQL & ", EAN "
                sSQL = sSQL & ", LPZ "
                sSQL = sSQL & ", LINR "
                sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
                sSQL = sSQL & " from etiprots where Bestand > 0 and WEDATE = " & CLng(DateValue(Left(cboPreisAenderungen.Text, 8)))
                
                gdBase.Execute sSQL, dbFailOnError
    
                gsETILS = "aus Lieferschein"
                frmWKL30.Show 1
                
                
                Screen.MousePointer = 0
            Else
                MsgBox "W‰hlen Sie bitte ein Datum aus!", vbOKOnly + vbInformation, "Winkiss Hinweis:"
            End If
            
            
            
            Screen.MousePointer = 11
            
            
        
        
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub fuelle_PreisAenderungen(cbox As ComboBox)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As DAO.Recordset
    Dim ctemp       As String
    
    Screen.MousePointer = 11
    
    Command1(9).Visible = False
    Label2(11).Visible = False
    cbox.Visible = False
    cbox.Clear
    
    If NewTableSuchenDBKombi("ETIPROTS", gdBase) Then
        
        cSQL = "select  distinct(WEDATE) as disdatum , sum(bestand) as mBestand from ETIPROTS "
        cSQL = cSQL & " group by WEDATE order by WEDATE desc"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!disdatum) Then
                    ctemp = Format(rsrs!disdatum, "DD.MM.YY")
                Else
                    ctemp = ""
                End If
                
                If ctemp <> "" Then
                
                    If Not IsNull(rsrs!mBestand) Then
                        
                        If Val(rsrs!mBestand) > 0 Then
                            ctemp = ctemp & " (" & rsrs!mBestand & ")"
                            If cbox.Text = "" Then
                                cbox.Text = ctemp
                            End If
                            
                            cbox.AddItem ctemp
                            cbox.Visible = True
                            Command1(9).Visible = True
                            Label2(11).Visible = True
                        
                        End If
                        
                    End If
                End If
                
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_PreisAenderungen"
    Fehler.gsFehlertext = "Im Programmteil Etiketten selbst w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leerStandard()
    On Error GoTo LOKAL_ERROR
    
    If Check2.Value = vbUnchecked Then
        Text1(0).Text = ""
    End If
    
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    List3.Clear
    List3.Visible = False
'    Option1(0).Value = True
    If Option1(1).Value = True Then
        Text1(2).Text = "1"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leerStandard"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDElesen()
    On Error GoTo LOKAL_ERROR
    
    If MDEeinlesenOhneLinr(Label2(9), txtStatus, picprogress, frmWKL31) = False Then
        anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Ger‰t ausgelesen werden.", Label2(9)
    Else
        anzeigeNew "normal", "", Label2(9)
        MdeVerarbeitung
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDElesen"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsFilBu     As Recordset
    Dim rsArt       As Recordset
    Dim seekEAN     As String
    Dim lMenge      As Long
    Dim lscanfolge  As Long
    
    Screen.MousePointer = 11
    
    sSQL = "Select * from ARTERRIN"

    Set rsFilBu = gdBase.OpenRecordset(sSQL)
    
    mdeErr = False
    lscanfolge = 0
    
    anzeigeNew "normal", "Die Daten aus dem MDE - Ger‰t werden verarbeitet...", Label2(9)
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
        
            lscanfolge = lscanfolge + 1
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
                ElseIf Len(seekEAN) <= 6 Then
                    
                    sSQL = "select * from artikel where artnr = " & seekEAN
                    
                Else
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                End If

                Set rsArt = gdBase.OpenRecordset(sSQL)
                
                If Not rsArt.EOF Then 'hier die bekannten
                
                    'Etiketten f¸llen
                    
                    schreibeWKEtidru rsArt!artnr, CLng(rsMDE!Menge), Val(gcFilNr)
                    
                Else 'hier die unbekannten
                
                    mdeErr = True
                    rsFilBu.AddNew
                    rsFilBu!EAN = seekEAN
                    rsFilBu!Menge = rsMDE!Menge
                    rsFilBu!lfnr = lscanfolge
                    
                    rsFilBu.Update
                    
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsMDE.MoveNext
        Loop
    
    End If
    
    rsMDE.Close: Set rsMDE = Nothing
    rsFilBu.Close: Set rsFilBu = Nothing
    
    anzeigeNew "normal", "Einlesevorgang erfolgreich", Label2(9)
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MdeVerarbeitung"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    positionieren31
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label1
    
    f¸lleCboEtiketten cboStrichEndlos
    
    fuelle_PreisAenderungen cboPreisAenderungen
    
    isEtidruFree = True
    
    bFocusLibesnr = False
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = "1"
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    
    Label2(9).Caption = ""
    
    Frame2.Visible = True
    
    If Modul6.FindFile(App.Path, "aWOKIBR.rpt") Then
        Command1(3).Visible = True
    Else
        Command1(3).Visible = False
    End If
    
    If Modul6.FindFile(App.Path, "aWOKINE.rpt") Then
        Command1(4).Visible = True
    Else
        Command1(4).Visible = False
    End If
    
    If Modul6.FindFile(App.Path, "aWKL30xs.rpt") Then
        Command1(5).Visible = True
    End If
    
    If Modul6.FindFile(gcDBPfad, "aWKL30ys.rpt") Then
        Command1(5).Visible = True
    End If
    
    If Modul6.FindFile(App.Path, "aWKL30zs.rpt") Then
        Command1(8).Visible = True
    End If
    
    If gbEtiFokEan Then
        Text1(1).TabIndex = 0
    Else
        Text1(5).TabIndex = 0
    End If
    
    If gbEtiQuickScanM Then
        Check3.Value = vbChecked
    End If
    
    If NewTableSuchenDBKombi("E31", gdApp) = False Then
        CreateTableT2 "E31", gdApp
    Else
        If SpalteInTabellegefundenNEW("E31", "Eti", gdApp) = False Then
            SpalteAnfuegenNEW "E31", "Eti", "Text(20)", gdApp
        End If
    End If
    
    voreinstellungladenE31
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE31()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim bo4     As Integer
    
    loeschNEW "E31", gdApp
    CreateTableT2 "E31", gdApp
    
    If Check4.Value = vbChecked Then
        bo4 = 0
    Else
        bo4 = -1
    End If
   
    sSQL = "Insert into E31 ( bo4,ETI) "
    sSQL = sSQL & " values (" & bo4 & ",'" & cboStrichEndlos.Text & "')"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE31"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE31()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim sEti As String

    Set rs = gdApp.OpenRecordset("E31")
    If Not rs.EOF Then
        
        If rs!bo4 = True Then
            Check4.Value = vbUnchecked
        Else
            Check4.Value = vbChecked
        End If
        
        sEti = "bitte ausw‰hlen"
        If Not IsNull(rs!Eti) Then
            sEti = Trim(rs!Eti)
        End If

        
        cboStrichEndlos.Text = sEti
    End If
    rs.Close: Set rs = Nothing
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE31"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub positionieren31()
    On Error GoTo LOKAL_ERROR
    
    If gcFilNr = "1" Then
        Frame1.Width = 8775
        Frame4.Visible = True
    Else
        Frame1.Height = 6135
        Frame1.Left = 120
        Frame1.Top = 480
        Frame1.Width = 8775
        Frame4.Visible = False
    End If
    
    Frame2.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionieren31"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    LogtoEnd Me
'    AktionAustragen "Etiketten w‰hlen"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    bBlau = False
    
    Select Case Index
        Case Is = 0
            Frame2.Visible = False
            Text1(2).Text = ""
            Check3.Visible = False
            Check3.Value = vbUnchecked
        Case Is = 1
            Frame2.Visible = True
            If gcFilNr = "1" Then
                Text1(3).Text = "1"
            Else
                Text1(3).Text = ""
            End If
            Text1(2).Text = "1"
            
            Text1(2).SelStart = 0
            Text1(2).SelLength = Len(Text1(2).Text)
            Text1(2).SetFocus
            bBlau = True
            
            Check3.Visible = True
            Check3.Value = vbUnchecked
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

     Select Case Index
        Case 1
            If Len(Text1(1).Text) > 5 Then
                If IsNumeric(Text1(1).Text) Then
                    Label2(8).Caption = bezis(Text1(1).Text, True)
                    
                    Label2(8).Caption = Label2(8).Caption & " " & Format(Preis_is(Text1(1).Text), "#,##0.00")
                    Label2(8).Refresh
                Else
                    Label2(8).Caption = ""
                    Label2(8).Refresh
                End If
            Else
                Label2(8).Caption = ""
                Label2(8).Refresh
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Label0.Caption = Index
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case Index
        Case 0, 2, 3, 4
    
            cValid = "1234567890" & Chr(8)
            cZeichen = Chr$(KeyAscii)
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
        
            Case Is = 1
            
            
            
                If Command1(0).Caption = "Speichern" Then
                    Dim cValid As String
                    Dim cFeld As String
                    Dim cZeichen As String
                    Dim lcount As Long
                    Dim bTextSuche As Boolean
                    
                    Screen.MousePointer = 11
                    
                    cValid = "1234567890"
                    cFeld = Text1(1).Text
                    
                    
                    bTextSuche = False
                    
                    For lcount = 1 To Len(cFeld)
                        cZeichen = Mid(cFeld, lcount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            bTextSuche = True
                            Exit For
                        End If
                    Next lcount
                    
                    If bTextSuche Then
                        gcSuch = Text1(1).Text
                        gsARTNR = ""
                        frmWKL70.Show 1
                        Me.Refresh
                        If gsARTNR <> "" Then
                            Text1(1).Text = gsARTNR
                            gsARTNR = ""
        
                            
                        End If
                    Else
                    Screen.MousePointer = 0
                    
                    
                    End If
'                ElseIf Command1(0).Caption = "Drucken" Then
'                    Command1_Click 0
                End If
            
            Case Is = 0
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
                
            Case 4
                ctmp = Text1(0).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                    Exit Sub
                End If
                
                
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = ctmp
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                    
                List3.Visible = False
                
                List3.Clear
                For lcount = 0 To 100
                    If gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        List3.AddItem gF2Prompt.cArray(lcount)
                    End If
                Next lcount

        End Select
        Text1(Index).SetFocus
        

    End If
    
    If Index = 1 Then
        If KeyCode = vbKeyReturn Then
        
            If Command1(0).Caption = "Speichern" Then
  
                Screen.MousePointer = 11
    
                cValid = "1234567890"
                cFeld = Text1(1).Text
                bTextSuche = False
    
                For lcount = 1 To Len(cFeld)
                    cZeichen = Mid(cFeld, lcount, 1)
                    If InStr(cValid, cZeichen) = 0 Then
                        bTextSuche = True
                        Exit For
                    End If
                Next lcount
    
                Screen.MousePointer = 0
    
                If bTextSuche Then
                    Text1_KeyUp 1, vbKeyF2, 0
                Else
                    If Check3.Value = vbChecked Then
                        Command1_Click 0
                    Else
                        Option1_Click 1
                        Option1(1).Value = True
                    End If
                End If
                
            ElseIf Command1(0).Caption = "Drucken" Then
                Command1_Click 0
            End If

        End If
    
    Else
        If KeyCode = vbKeyReturn Then
            Command1_Click 0
        End If
    End If
    
    
    
    If KeyCode = vbKeyEscape Then
        Command1_Click 2
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Etiketten w‰hlen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
