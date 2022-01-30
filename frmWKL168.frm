VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL168 
   Caption         =   "Penner"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   12
      Left            =   11280
      TabIndex        =   41
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
   Begin VB.PictureBox picprogress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   9315
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   10335
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   2775
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
         Caption         =   "Neuheiten ausschließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Height          =   1335
         Left            =   5400
         TabIndex        =   37
         Top             =   2280
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "nur Ex Artikel"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ohne Ex Artikel"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "alle Artikel"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   11
         Left            =   2760
         TabIndex        =   36
         Top             =   2280
         Width           =   1920
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
         Caption         =   "alle Lieferanten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   1
         Left            =   2760
         TabIndex        =   34
         Top             =   5640
         Width           =   1920
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
         Caption         =   "Liste leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   4575
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   10
         Left            =   1320
         TabIndex        =   32
         Top             =   2280
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
         Height          =   375
         Index           =   2
         Left            =   120
         MaxLength       =   6
         TabIndex        =   30
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "90"
         Top             =   1440
         Width           =   615
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
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
         Caption         =   "kleiner"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   16
         Text            =   "0,1"
         Top             =   480
         Width           =   495
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   6
         Left            =   9600
         TabIndex        =   12
         Top             =   120
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
         Caption         =   "Ermitteln"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
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
         Index           =   16
         Left            =   120
         TabIndex        =   35
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Lieferanten:"
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
         Index           =   15
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "anzeigen"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Neuheitendefinition: alle Artikel bis zu 90 Tage nach dem ersten WE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   3000
         TabIndex        =   21
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Tage"
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
         Index           =   5
         Left            =   840
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Lagerumschlagsgeschwindigkeit"
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   6975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   11775
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   9
         Left            =   9600
         TabIndex        =   46
         Top             =   4080
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   8
         Left            =   9600
         TabIndex        =   15
         Top             =   4680
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   5
         Left            =   11280
         TabIndex        =   13
         ToolTipText     =   "Kalender"
         Top             =   5160
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
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   7
         Top             =   120
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   9600
         TabIndex        =   6
         Top             =   5640
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
         Caption         =   "Markieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   4
         Left            =   9600
         TabIndex        =   5
         Top             =   6240
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11668
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
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
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   9600
         TabIndex        =   28
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "erster WE:"
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
         Index           =   12
         Left            =   9600
         TabIndex        =   27
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Index           =   11
         Left            =   9600
         TabIndex        =   26
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "letzter WE:"
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
         Index           =   10
         Left            =   9600
         TabIndex        =   25
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   9600
         TabIndex        =   24
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "letzter VK:"
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
         Index           =   8
         Left            =   9600
         TabIndex        =   23
         Top             =   1320
         Width           =   2055
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
         Left            =   9600
         TabIndex        =   14
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl der Artikel:"
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
         Index           =   2
         Left            =   9600
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "0"
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
         Index           =   3
         Left            =   9600
         TabIndex        =   8
         Top             =   960
         Width           =   2055
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
      Left            =   6000
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin sevCommand3.Command Command5 
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command11 
      Height          =   360
      Left            =   10800
      TabIndex        =   47
      Top             =   360
      Width           =   405
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
      ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
      ToolTipTitle    =   "Spaltenanordung"
      ButtonStyle     =   2
      Caption         =   ""
      Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
      Picture         =   "frmWKL168.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "F4: Lagerumschlagsgeschwindigkeit"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "F3: Filialspiegel"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   44
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "F2: Artikelmaske"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Penner"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Left            =   120
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL168"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerArtnr          As Byte
Dim SpaltennummerAWM            As Byte
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "Farbnr"
    gstab = "ARTLUG"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loescheausFilz(lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr          As String
    Dim cSQL            As String
   
    cArtNr = MSFlexGrid1.TextMatrix(lrow, SpaltennummerArtnr)
    
    If cArtNr <> "" Then
        If IsNumeric(cArtNr) Then
            cSQL = "Delete from art45 "
            cSQL = cSQL & " where ARTNR = " & cArtNr
            gdBase.Execute cSQL, dbFailOnError
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheausFilz"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FlexGrid_Update(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
  
    Dim lBig As Long
  
    
  With oGrid
    ' aktuelle Selektion merken
    nRow = .Row
    nCol = .Col
    nRowSel = .RowSel
    nColSel = .ColSel
    
    
    If nRow >= nRowSel Then
        lBig = nRow
        nDelRow = nRowSel - 1
    Else
        lBig = nRowSel
        nDelRow = nRow - 1
    End If
    
    
    Do While nDelRow < lBig
    
        nDelRow = nDelRow + 1
        
        If nDelRow >= 1 Then
            loescheausFilz nDelRow
        End If
    Loop
  End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cFarbkenn   As String
    Dim iRet        As Integer
    Dim ctmp        As String
    Dim lcount      As Long
    Dim i           As Integer

    Select Case Index
        Case 0
            Unload frmWKL168
        Case 1
            List3.Clear
            List3.Visible = False
            
            Command5(1).Visible = False
            Label1(16).Visible = False
        Case 2
            If MSFlexGrid1.RowSel >= 1 Then
                FlexGrid_Update MSFlexGrid1
            End If
            ZeigeArtikel58
        Case 3
            ctmp = Trim$(Label4(32).Tag)
            If ctmp <> "" Then
                cFarbkenn = ermFarbkz(ctmp)
            Else
                cFarbkenn = "ohne Kennzeichen"
                ctmp = "0"
            End If
            iRet = MsgBox("Möchten Sie jetzt alle Artikel aus der Tabelle mit dem Farbkennzeichen '" & cFarbkenn & "' versehen?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                Farbanpassung ctmp
                ZeigeArtikel58
            End If
        Case 4
            Frame1.Visible = True
            Frame2.Visible = False
            anzeige "normal", "", Label1(4)
            
            Label2(0).Visible = False
            Label2(1).Visible = False
            Label2(2).Visible = False
        Case 5
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
                Label4(32).Caption = "ohne Kennzeichen"
            End If
        Case 6
            
            ermPenner
            ZeigeArtikel58
        Case 7
            If Command5(7).Caption = "kleiner" Then
                Command5(7).Caption = "größer"
            Else
                Command5(7).Caption = "kleiner"
            End If
            
        Case 8
            If Not NewTableSuchenDBKombi("art45", gdBase) Then
                anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
                Exit Sub
            Else
                If Datendrin("art45", gdBase) = False Then
                    anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
                    Exit Sub
                End If
            End If
        
            LzuFuellen
            DruckeArtikel176
        Case 9
        
            
            LzuFuellen
            ExcelExport "art45", gdBase
            
        Case 10
            Text1_KeyUp 2, vbKeyF2, 0
        Case 11
            alleLiefInListe3
            Command5(1).Visible = True
        Case 12
            gsHelpstring = "Penner"
            frmWKL110.Show 1
        Case 13
            If Command5(13).Caption = "Neuheiten ausschließen" Then
                Command5(13).Caption = "nur Neuheiten"
            Else
                Command5(13).Caption = "Neuheiten ausschließen"
            End If
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeArtikel58()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    
    MSFlexGrid1.Clear
    
    If Not NewTableSuchenDBKombi("art45", gdBase) Then
        anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
        Exit Sub
    Else
        If Datendrin("art45", gdBase) = False Then
            anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
            Exit Sub
        End If
    End If
    
    anzeige "rot1", "wird ermittelt...", Label1(3)
    anzeige "normal", "Artikel werden angezeigt, bitte warten...", Label1(4)
    
    Screen.MousePointer = 11
    
    Tabcheck "ARTLUG"
    FormatGridOverTablay "ARTLUG"

    With MSFlexGrid1
        .Redraw = False
'        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) '* 1.8
        Next j
    End With
    
    ermittlespalten
    GridFuellen "Select * from art45 order by lug"
    
    FaerbenGrid MSFlexGrid1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
    
    Frame2.Visible = True
    Frame1.Visible = False
    
    MSFlexGrid1.Row = 2
    MSFlexGrid1.Col = 1
    MSFlexGrid1.SetFocus
    
    MSFlexGrid1_SelChange
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikel58"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeArtikel176()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim i As Integer
   
    Screen.MousePointer = 11
    
    loeschNEW "ART45PRINT", gdBase
    CreateTableT2 "ART45PRINT", gdBase
    
    cSQL = "Insert into art45Print select "
    cSQL = cSQL & " ARTNR  "
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", EAN  "
    cSQL = cSQL & ", LINR  "
    cSQL = cSQL & ", LPZ  "
    cSQL = cSQL & ", PGN  "
    cSQL = cSQL & ", LIBESNR  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", LEKWERT  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", KVKWERT  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", liefbez  "
    cSQL = cSQL & ", ERSTDAT  "
    cSQL = cSQL & ", AUFDAT  "
    cSQL = cSQL & ", EXDAT  "
    cSQL = cSQL & ", LASTVK  "
    cSQL = cSQL & ", LASTZU  "
'    cSQL = cSQL & ", Filwahl "
'    cSQL = cSQL & ", ARTwahl  "
'    cSQL = cSQL & ", Neuwahl  "
'    cSQL = cSQL & ", NeuHeitwahl  "
'    cSQL = cSQL & ", LUGwahl  "
    cSQL = cSQL & ", FARBTEXT "
    cSQL = cSQL & ", FARBwert  "
    cSQL = cSQL & ", FARBwertS  "
    cSQL = cSQL & ", FARBNR  "
    cSQL = cSQL & ", LUG  "
    cSQL = cSQL & ", AGN  "
    cSQL = cSQL & ", Marke  "
    cSQL = cSQL & ", LINBEZ  "
    cSQL = cSQL & " from art45  "
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 0 To 2
        If Option2(i).Value = True Then
            cSQL = "Update art45Print set Artwahl = '" & Option2(i).Caption & "' "
            gdBase.Execute cSQL, dbFailOnError
        End If
    Next i
    
    cSQL = "Update art45Print set Neuwahl = '" & Label1(6).Caption & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    Dim sLugwahl As String
    
    If Command5(7).Caption = "kleiner" Then
        sLugwahl = "LUG < " & Text1(0).Text
    Else
        sLugwahl = "LUG > " & Text1(0).Text
    End If
    
    cSQL = "Update art45Print set LUGwahl = '" & sLugwahl & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update art45Print set NeuHeitwahl = '" & Command5(13).Caption & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    If Datendrin("art45print", gdBase) = False Then
        anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
        Exit Sub
    End If
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
    
    reportbildschirm "", "aZEN176a"
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeArtikel176"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LzuFuellen()
    On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim siAnzeige As Single
    Dim cART As String
    Dim datLZU As Date
    Dim datLVK As Date
    
    Screen.MousePointer = 11
    
    Set rsrs = gdBase.OpenRecordset("art45")
    If Not rsrs.EOF Then

        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)

            If Not IsNull(rsrs!artnr) Then
                cART = rsrs!artnr
                rsrs.Edit
                datLZU = ErmlzZugang(cART)
                datLVK = ErmlzVK(cART)
                
                If datLVK = "01.01.1980" Then
                    rsrs!lastvk = Null
                Else
                    rsrs!lastvk = datLVK
                End If
                
                If datLZU = "01.01.1980" Then
                    rsrs!lastzu = Null
                Else
                    rsrs!lastzu = datLZU
                End If
                
                
                
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
     
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LzuFuellen"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub LzuUpdate()
'    On Error GoTo LOKAL_ERROR
'
'    Dim rsrs As Recordset
'    Dim lAnz As Long
'    Dim siAnzeige As Single
'    Dim cART As String
'    Dim datLVK As Date
'
'
'    Screen.MousePointer = 11
'
'    Set rsrs = gdBase.OpenRecordset("art45")
'    If Not rsrs.EOF Then
'
'        rsrs.MoveLast
'        lAnz = rsrs.RecordCount
'        rsrs.MoveFirst
'        Do While Not rsrs.EOF
'
'            siAnzeige = siAnzeige + 1
'            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
'
'            If Not IsNull(rsrs!artnr) Then
'                cART = rsrs!artnr
'                rsrs.Edit
'                datLVK = ErmlzVKproFil(cART, iFil)
'                rsrs!LASTVK = datLVK
'                rsrs.Update
'            End If
'        rsrs.MoveNext
'        Loop
'
'    End If
'    rsrs.Close: Set rsrs = Nothing
'
'    Screen.MousePointer = 0
'
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "LzuUpdate"
'    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "FARBNR"
                SpaltennummerAWM = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

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
    Dim lAnz        As Long
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    picprogress.Visible = True
    With MSFlexGrid1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        lAnz = lMax
        

'        Anzeige "normal", "Es werden " & lMax & " Artikel angezeigt...", Label1(4)
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            .Rows = lrow + 1
            .Col = 0
            
            txtStatus.Text = (lrow * 100) / lMax
            
            
            
            Select Case lMax
                Case Is > 5000
                
                    j = lAnz Mod 500
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
                
                Case Is > 1000
                
                    j = lAnz Mod 100
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
                
                Case Is <= 500
                
                    j = lAnz Mod 50
                    If j = 0 Then
                        anzeige "normal", "Es werden noch " & lAnz & " zur Anzeige vorbereitet...", Label1(4)
                    End If
        
            End Select
    
            lAnz = lAnz - 1
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case UCase(sSpaltenname(i))
                        Case Is = "LEK", "KVK", "LUG", "LEK-WERT", "KVK-WERT"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            
                        Case Is = "RKZ"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "N"
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                        Case Is = "FARBE"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert
                            FaerbenFlex sWert, MSFlexGrid1, 0, CInt(lrow)
                        
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
        
        Frame2.Visible = True
        
        anzeige "normal", CStr(lMax), Label1(3)
        anzeige "normal", lMax & " Artikel", Label1(4)
        
        Label2(0).Visible = True
        If Val(gcFilNr) > 0 Then
            Label2(1).Visible = True
        End If
        Label2(2).Visible = True
    Else
        Frame2.Visible = False
        anzeige "normal", "", Label1(3)
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label1(4)
        
        Label2(0).Visible = False
        Label2(1).Visible = False
        Label2(2).Visible = False
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
    
    picprogress.Visible = False
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    
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
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenWKL58
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Me.Refresh
    
    Frame1.Visible = True
    Frame2.Visible = False
    
    Text1(2).Text = "alle"
    
    If NewTableSuchenDBKombi("List3", gdBase) Then
        LadeList3
        Command5(1).Visible = True
        Label1(16).Visible = True
        Label1(16).Caption = List3.ListCount & " Lieferanten"
        Label1(16).Refresh
    End If
    
    VorBereitLagerumschlag
        
    anzeige "normal", "", Label1(3)
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
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
Private Sub Farbanpassung(cFabm As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    Screen.MousePointer = 11
    
    sSQL = "update art45 set farbnr = " & Val(cFabm) & " "
    gdBase.Execute sSQL, dbFailOnError
    
    BringFarbeInsSpiel "Art45", gdBase
    
    sSQL = "update artikel inner join art45 on artikel.artnr = art45.artnr"
    sSQL = sSQL & " set AWM = '" & cFabm & "'"
    sSQL = sSQL & " , LASTDATE = '" & DateValue(Now) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Farbanpassung"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL58()
On Error GoTo LOKAL_ERROR

    With Frame1
        .Top = 960
        .Height = 6735
        .Width = 11775
        .Left = 0

    End With

    With Frame2
        .Top = 960
        .Height = 6735
        .Width = 11775
        .Left = 0

    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL58"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    
    If List3.ListCount > 0 Then
        iRet = MsgBox("Möchten Sie die ausgewählten Lieferanten für die nächste Ermittlung abspeichern", vbYesNo + vbQuestion, "Winkiss Frage:")
        If iRet = vbYes Then
            SpeicherList3
        End If
    End If

    loeschNEW "art45", gdBase
    loeschNEW "ART45PRINT", gdBase
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
Private Sub SpeicherList3()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim i           As Integer
    
    loeschNEW "LIST3", gdBase
    
    sSQL = "Create Table List3 (AuswahlTEXT Text(100)) "
    gdBase.Execute sSQL, dbFailOnError

    cLBSatz = ""
    For i = 0 To List3.ListCount - 1
        cLBSatz = List3.list(i)
        sSQL = "Insert into List3 (AuswahlTEXT) "
        sSQL = sSQL & " Values (  '" & cLBSatz & "')"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherList3"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LadeList3()
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    List3.Clear
    List3.Visible = True
    
    Text1(2).Text = ""
    
    cSQL = "Select * from LIST3 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!AuswahlTEXT) Then
                List3.AddItem rsrs!AuswahlTEXT
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeList3"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermPenner()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim lHeute As Long
    Dim lAnz As Long
    Dim siAnzeige As Single
    Dim bAnd As Boolean
    Dim sLUG As String
    Dim i As Integer
    
    bAnd = False
    
    lHeute = CLng(DateValue(Now))
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "ART45", gdBase
    CreateTableT2 "ART45", gdBase
    
    anzeige "normal", "die Artikel werden ermittelt...", Label1(4)

    sSQL = " Insert into art45 select  distinct(a.ARTNR)"
    sSQL = sSQL & " , a.Bezeich "
    sSQL = sSQL & " , a.EAN "
    sSQL = sSQL & " , b.RKZ "
    sSQL = sSQL & " , b.LEKPR "
    sSQL = sSQL & " , a.KVKPR1 "
    sSQL = sSQL & " , a.LINR "
    sSQL = sSQL & " , a.LPZ "
    sSQL = sSQL & " , a.PGN "
    sSQL = sSQL & " , a.AGN "
    sSQL = sSQL & " , a.BESTAND "
    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", a.AUFDAT  "
    sSQL = sSQL & ", b.EXDAT  "
    sSQL = sSQL & ", val(a.AWM) as FARBNR "
    sSQL = sSQL & ", l.LUG "
    
    
    sSQL = sSQL & ", null as Last_ZU"
    sSQL = sSQL & ", null as Last_VK"
    sSQL = sSQL & ", a.LIBESNR from Artikel a , ALLARTLU l,artlief b "
    sSQL = sSQL & " where a.artnr = l.artnummer "
    sSQL = sSQL & " and a.artnr = b.artnr "
    
    bAnd = True
    If Text1(0).Text <> "" Then
    
        sLUG = SwapStr(Text1(0).Text, ",", ".")
        
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        If Command5(7).Caption = "kleiner" Then
            sSQL = sSQL & " l.LUG < " & sLUG & " "
        Else
            sSQL = sSQL & " l.LUG > " & sLUG & " "
        End If
        bAnd = True
    End If
    
    'nur Ex Artikel
    If Option2(0).Value = True Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        'MussRKZ
        sSQL = sSQL & " b.RKZ = 'J' "
        bAnd = True
    End If
    
    'ohne Ex Artikel
    If Option2(1).Value = True Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & " b.RKZ <> 'J' "
        bAnd = True
    End If
    
    
    'Lieferant
    If Text1(2).Text <> "alle" Or IsNumeric(Text1(2).Text) Then
        If Text1(2).Text = "" Then
            If List3.ListCount = 0 Then

            Else
            
                If bAnd Then
                    sSQL = sSQL & " and "
                End If
                
                sSQL = sSQL & " (b.LINR= " & Val(Left(List3.list(0), 6))
                For i = 1 To List3.ListCount - 1
                    sSQL = sSQL & " or b.LINR= " & Val(Left(List3.list(i), 6))
                Next i
                sSQL = sSQL & ")"
            End If
        Else
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & " b.LINR = " & Trim(Text1(2).Text)
        End If
   
    End If

    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    anzeige "normal", "Bestände werden aktualisiert...", Label1(4)

    txtStatus.Text = 42
    
    anzeige "normal", "nicht relevante Artikel werden gelöscht...", Label1(4)

    sSQL = "Delete from art45 where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 56

    sSQL = "Delete from art45 where bestand is null "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 63
    
    anzeige "normal", "Neuheiten werden selektiert...", Label1(4)
    
    Set rsrs = gdBase.OpenRecordset("art45")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                rsrs.Edit
                rsrs!ERSTDAT = ErmFirstZugang(rsrs!artnr)
                rsrs.Update
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    txtStatus.Text = 71
    
    If Command5(13).Caption = "Neuheiten ausschließen" Then
    
        anzeige "normal", "Neuheiten werden entfernt...", Label1(4)
        
        sSQL = " Delete from art45  where ERSTDAT > datevalue(now) -  '" & Text1(1).Text & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = " Delete from art45  where ERSTDAT = DateValue('01.01.1980')"
        gdBase.Execute sSQL, dbFailOnError
        
    Else
        anzeige "normal", "nur Neuheiten beibehalten...", Label1(4)
        
        sSQL = " Delete from art45  where ERSTDAT < datevalue(now) -  '" & Text1(1).Text & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        'aber die letzten 14 Tage ausblenden
        sSQL = " Delete from art45  where ERSTDAT > datevalue(now) - 14 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    txtStatus.Text = 78
    
    anzeige "normal", "Pennerwerte werden ermittelt...", Label1(4)
    
    sSQL = "Update art45 Set LEKWERT =  Bestand * LEKPR "
    sSQL = sSQL & " , KVKWERT = BESTAND * KVKPR1 "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 85

    sSQL = "Update art45 inner join lisrt on art45.linr = lisrt.linr "
    sSQL = sSQL & " Set art45.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 95
    
    anzeige "normal", "Markeninformationen werden abgeglichen...", Label1(4)
    
    Markenabgleich "Art45", gdBase
    
    txtStatus.Text = 98
    
    BringFarbeInsSpiel "Art45", gdBase
    
    
    
    'Duplikate löschen
    
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lcount          As Long
    Dim cArtNr          As String
    
    loeschNEW "alit" & srechnertab, gdBase
    sSQL = "select count(Artnr) as count ,Artnr into alit" & srechnertab & " from Art45 group by Artnr having count(Artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artdupli", gdBase
    sSQL = "Select * into artDupli from Art45 where artnr = -1 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Ermittlung der Duplikate...", Label1(4)
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli", dbOpenTable)
    
    Set rsrs = gdBase.OpenRecordset("alit" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            sSQL = "Select * from Art45 where artnr = " & cArtNr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close
    
    loeschNEW "alit" & srechnertab, gdBase
    
    anzeige "normal", "Ermittlung letzte Wareneingänge...", Label1(4)
    
    loeschNEW "Last_ZU" & srechnertab, gdBase
    
    sSQL = "Select Zugang.Artnr, Max(adate) as LASTZU into Last_ZU" & srechnertab
    sSQL = sSQL & " from Zugang inner join Art45 on Zugang.ARTNR = Art45.Artnr  "
    sSQL = sSQL & " group by  Zugang.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Art45 inner join Last_ZU" & srechnertab & " on Art45.Artnr = Last_ZU" & srechnertab & ".Artnr "
    sSQL = sSQL & " Set Art45.Last_ZU = Last_ZU" & srechnertab & ".LASTZU  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Last_ZU" & srechnertab, gdBase
    
    anzeige "normal", "Ermittlung letzte Verkäufe...", Label1(4)
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    sSQL = "Select Kassjour.Artnr, Max(adate) as LASTVK into Last_VK" & srechnertab
    sSQL = sSQL & " from KASSJOUR inner join Art45 on KASSJOUR.ARTNR = Art45.Artnr  "
    sSQL = sSQL & " group by  KASSJOUR.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Art45 inner join Last_VK" & srechnertab & " on Art45.Artnr = Last_VK" & srechnertab & ".Artnr "
    sSQL = sSQL & " Set Art45.Last_VK = Last_VK" & srechnertab & ".LASTVK  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    
    
    
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0

        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermPenner"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Label4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 32 Then
    Label4(Index).Caption = "ohne Kennzeichen"
    Label4(Index).Tag = ""
    Label4(Index).BackColor = Label1(2).BackColor
    Label4(Index).ForeColor = Label1(2).ForeColor
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = 46    'Del
            If List3.ListCount = 0 Then
                List3.Clear
                List3.Visible = False
               
                Command5(1).Visible = False
                Label1(16).Visible = False
            Else
                If Not List3.ListIndex = -1 Then
                    List3.RemoveItem (List3.ListIndex)
                    If List3.ListCount = 0 Then
                        List3.Clear
                        List3.Visible = False
                        
                        Command5(1).Visible = False
                        Label1(16).Visible = False
                    Else
                        Label1(16).Visible = True
                        Label1(16).Caption = List3.ListCount & " Lieferanten"
                        Label1(16).Refresh
                    End If
                End If
            End If
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    If KeyCode = vbKeyF2 Then
        lrow = MSFlexGrid1.Row
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
        If gsARTNR <> "" Then
            
            frmWKL10.Show 1
'                frmZENcb.Show 1
            Me.Refresh
            Screen.MousePointer = 11

            MSFlexGrid1.TopRow = lrow
            MSFlexGrid1.Col = SpaltennummerArtnr
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    ElseIf KeyCode = vbKeyF3 Then
        gcArtNrFiliale = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
        If IsNumeric(gcArtNrFiliale) Then
            frmWKLae.Show 1
        Else
            gcArtNrFiliale = ""
        End If
    ElseIf KeyCode = vbKeyF4 Then
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
        If IsNumeric(gsARTNR) Then
            frmWKL62.Show 1
        End If
        gsARTNR = ""
    End If
    
    MSFlexGrid1.Redraw = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error GoTo LOKAL_ERROR

Dim cART As String

cART = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)

If cART <> "" Then
    If IsNumeric(cART) Then
    
    Label1(9).Caption = ErmlzVK(cART)
    Label1(11).Caption = ErmlzZugang(cART)
    Label1(13).Caption = ErmFirstZugang(cART)
    
    
    If Right(Label1(11).Caption, 2) = "80" Then Label1(11).Caption = ""
    If Right(Label1(13).Caption, 2) = "80" Then Label1(13).Caption = ""
    
    End If
End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 1 Then
        Label1(6).Caption = "Neuheitendefinition: alle Artikel bis zu " & Text1(1).Text & " Tage nach dem ersten WE"
        Label1(6).Refresh
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

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
        Case 0
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 1, 2
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
'''            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
'''            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
'''            cValid = cValid & "+äÄÜüÖöß#"
'''
'''            If InStr(cValid, cZeichen) = 0 Then
'''                KeyAscii = 0
'''            End If
'''            'alle Zeichen erlaubt
    End Select
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub alleLiefInListe3()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cLBSatz As String
    Dim ctmp    As String
    Dim rsrs    As Recordset
    Dim sLinr   As String
    
    List3.Clear
    List3.Visible = True
    
    Text1(2).Text = ""
    
    Screen.MousePointer = 11
    
    cSQL = "Select * from LISRT "
    cSQL = cSQL & " where HL <> True and not kuerzel is null and kuerzel <> '' "
    cSQL = cSQL & " order by Liefbez "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLBSatz = ""
            If Not IsNull(rsrs!linr) Then
                ctmp = rsrs!linr
                sLinr = rsrs!linr
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(6 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & " "
            
            If Not IsNull(rsrs!Kuerzel) Then
                ctmp = rsrs!Kuerzel
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(6 - Len(ctmp))
            cLBSatz = cLBSatz & ctmp & " "
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                ctmp = rsrs!LIEFBEZ
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            cLBSatz = cLBSatz & ctmp & " "
            
            If sLinr <> "" Then
                If IsNumeric(sLinr) Then
                    If LAGERBestand(sLinr) > 0 Then
    
                        If ohneArtikel(sLinr) > 0 Then
                
                            List3.AddItem cLBSatz
                            
                            Label1(16).Visible = True
                            Label1(16).Caption = List3.ListCount & " Lieferanten"
                            Label1(16).Refresh
                            
                            List3.Refresh
                            
                            
                        End If
                    End If
                End If
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alleLiefInListe3"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim sAuswahlfeld As String
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyReturn Then
        Command5_Click 6
    End If
    
    If KeyCode = vbKeyEscape Then
        Command5_Click 0
    End If

    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = True
        
        Select Case Index
            Case 2
                gF2Prompt.cFeld = "LINR"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                        Text1(Index).Text = Trim(Text1(Index).Text)
                    End If
                End If
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                    
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 6)
                            Text1(Index).Text = Trim(Text1(Index).Text)
                        End If
                        
                    End If
                Next lcount
                
                If List3.Visible = True Then
                    Label1(16).Visible = True
                    Label1(16).Caption = List3.ListCount & " Lieferanten"
                    Label1(16).Refresh
                    Command5(1).Visible = True
                Else
                    Label1(16).Visible = False
                    Command5(1).Visible = False
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
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

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

