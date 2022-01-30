VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL169 
   BackColor       =   &H80000001&
   Caption         =   "Kundenbeteiligung"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   615
      Left            =   120
      TabIndex        =   77
      Top             =   6360
      Visible         =   0   'False
      Width           =   11775
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   9480
         TabIndex        =   83
         Top             =   2760
         Width           =   2055
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
         Caption         =   "Kundendaten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   9
         Left            =   9480
         TabIndex        =   82
         Top             =   2280
         Width           =   2055
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
         Caption         =   "Verkaufsdetails"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   10
         Left            =   9480
         TabIndex        =   81
         Top             =   1800
         Width           =   2055
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
         Caption         =   "alle zurücksetzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   11
         Left            =   11160
         TabIndex        =   80
         ToolTipText     =   "Kalender"
         Top             =   120
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
         Height          =   495
         Index           =   12
         Left            =   9480
         TabIndex        =   79
         Top             =   600
         Width           =   2055
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
         Caption         =   "zurücksetzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
         Height          =   5535
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Visible         =   0   'False
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9763
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Index           =   0
         Left            =   9480
         TabIndex        =   84
         Top             =   240
         Width           =   1575
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   75
      Top             =   8040
      Visible         =   0   'False
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
      Caption         =   "Excel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   7
      Left            =   8040
      TabIndex        =   71
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
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
      Caption         =   "Euro pro Stück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   7320
      MaxLength       =   5
      TabIndex        =   70
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   69
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
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   6
      Left            =   9600
      TabIndex        =   68
      Top             =   7560
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   63
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
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
      Caption         =   "vom Rohertrag"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   62
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
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
      Caption         =   "vom Umsatz"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   5760
      MaxLength       =   5
      TabIndex        =   61
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   41
      Top             =   7080
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   10815
      Begin VB.CheckBox Check5 
         Caption         =   "rabattierte Artikel ausschließen"
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
         Left            =   6120
         TabIndex        =   74
         Top             =   3600
         Width           =   3735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "keine Artikel unter Listenpreis"
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
         Left            =   6120
         TabIndex        =   73
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "nur bonusfähige Artikel"
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
         Left            =   6120
         TabIndex        =   72
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   5520
         MaxLength       =   13
         TabIndex        =   59
         Top             =   120
         Width           =   1575
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   58
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   56
         Top             =   960
         Width           =   855
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   1
         Left            =   4320
         TabIndex        =   55
         Top             =   600
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
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   0
         Left            =   9480
         TabIndex        =   52
         Top             =   600
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   9000
         MaxLength       =   6
         TabIndex        =   51
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "nur umsatzrelevante Artikelverkäufe"
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
         Left            =   6120
         TabIndex        =   50
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   2055
         Left            =   3120
         TabIndex        =   43
         Top             =   3840
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "Vorjahr Zeitraum"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "aktuelles Jahr"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vorjahr"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "aktueller Monat"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vormonat"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Datum Voreinstellung"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ohne Gutscheine"
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
         Left            =   6120
         TabIndex        =   29
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   3120
         TabIndex        =   24
         Top             =   2160
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "nur Warengruppen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Tag             =   "Menge"
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ausschließen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Tag             =   "Preis"
            Top             =   720
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "einschließen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Tag             =   "Ertrag"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Warengruppen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Kundenname"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   66
            Tag             =   "Menge"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Kundennummer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Tag             =   "Menge"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rohertrag"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Tag             =   "Ertrag"
            Top             =   360
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Umsatz"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Tag             =   "Preis"
            Top             =   720
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Stückzahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Tag             =   "Menge"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "sortiert nach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.ComboBox cboFil 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         Style           =   2  'Dropdown-Liste
         TabIndex        =   18
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   4680
         MaxLength       =   6
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   6720
         MaxLength       =   13
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   5
         Left            =   5400
         TabIndex        =   14
         Top             =   600
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
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   4
         Left            =   8640
         TabIndex        =   13
         Top             =   600
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   2
         Left            =   6360
         TabIndex        =   11
         Top             =   600
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
      Begin sevCommand3.Command Command0 
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   10
         Top             =   240
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
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   9960
         MultiSelect     =   2  'Erweitert
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   7920
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   6
         Left            =   9480
         TabIndex        =   7
         Top             =   120
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
      Begin MSComCtl2.DTPicker Text2 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65273857
         UpDown          =   -1  'True
         CurrentDate     =   38453
      End
      Begin MSComCtl2.DTPicker Text2 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   32
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   65273857
         UpDown          =   -1  'True
         CurrentDate     =   38453
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   7
         Left            =   1440
         TabIndex        =   86
         ToolTipText     =   "Kalender"
         Top             =   960
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
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   8
         Left            =   3360
         TabIndex        =   87
         ToolTipText     =   "Kalender"
         Top             =   960
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
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000001&
         Caption         =   "ArtikelNr / EAN"
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
         Index           =   10
         Left            =   3960
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Bed"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   57
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PGN"
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
         Index           =   0
         Left            =   9120
         TabIndex        =   53
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Filialauswahl"
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
         Index           =   12
         Left            =   5400
         TabIndex        =   40
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Nr"
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
         Index           =   2
         Left            =   4680
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Bestell-Nr."
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
         Index           =   3
         Left            =   6840
         TabIndex        =   38
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "AGN"
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
         Index           =   4
         Left            =   8280
         TabIndex        =   37
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Linie"
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
         Index           =   8
         Left            =   5880
         TabIndex        =   36
         Top             =   720
         Width           =   495
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
         Left            =   1080
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Verkaufszeitraum"
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
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000001&
         Caption         =   "Marke"
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
         Index           =   9
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   615
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
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   5235
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   5295
   End
   Begin sevCommand3.Command cmdEnd 
      Height          =   375
      Left            =   9600
      TabIndex        =   1
      Top             =   8040
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
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
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
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   4
      Left            =   10800
      TabIndex        =   85
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
      Picture         =   "frmWKL169.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label20 
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   76
      Top             =   7800
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Prozent"
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
      Index           =   14
      Left            =   7440
      TabIndex        =   67
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "vom Umsatz"
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
      Index           =   13
      Left            =   7440
      TabIndex        =   65
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Prozent"
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
      Index           =   11
      Left            =   6360
      TabIndex        =   64
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblAnzeige 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   8160
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenbeteiligung"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmWKL169"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerKUNDNR As Integer
Dim SpaltennummerAUSG As Integer

Dim lAusgewählt As Long

Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL169
        
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Text1_KeyUp 0, vbKeyF2, 0
        Case Is = 1
            Text1_KeyUp 1, vbKeyF2, 0
        Case Is = 4
            Text1_KeyUp 4, vbKeyF2, 0
        Case Is = 5
            Text1_KeyUp 2, vbKeyF2, 0
        Case Is = 2
            Text1_KeyUp 5, vbKeyF2, 0
        Case 3
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
                Label4(32).Caption = "alle Farben"
            End If
        Case Is = 7
            Text2(0).Value = Format(Datumschreiben11a(5600, Text2(0).Left), "DD.MM.YY")
            Text2(1).Value = Text2(0).Value
            
        Case Is = 8
            Text2(1).Value = Format(Datumschreiben11a(5600, Text2(1).Left), "DD.MM.YY")
            'fertig
        Case 6
            Text1_KeyUp 7, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub WKLatPositionieren()
    On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Top = 960
        .Left = 120
        .Height = 5775
        .Width = 11895
        .BorderStyle = 0
    End With
    
    With Frame5
        .Top = 960
        .Left = 120
        .Height = 5775
        .Width = 11895
        .BorderStyle = 0
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKLatPositionieren"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub KUNTOPIupdate()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    Dim cKdnr As String
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    MSHFLEX1.Redraw = False
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = SpaltennummerAUSG
    
    For lrow = 2 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If MSHFLEX1.Text = "" Then
            MSHFLEX1.Col = SpaltennummerKUNDNR
            cKdnr = MSHFLEX1.Text
            If IsNumeric(cKdnr) Then
                sSQL = "Update KUNTOPI set ausg = False where KUNDNR = " & cKdnr
                gdBase.Execute sSQL, dbFailOnError
            End If
        ElseIf MSHFLEX1.Text = "X" Then
            MSHFLEX1.Col = SpaltennummerKUNDNR
            cKdnr = MSHFLEX1.Text
            If IsNumeric(cKdnr) Then
                sSQL = "Update KUNTOPI set ausg = true where KUNDNR = " & cKdnr
                gdBase.Execute sSQL, dbFailOnError
            End If
        
        End If
    Next lrow
    
    MSHFLEX1.Redraw = True
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KUNTOPIupdate"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cFarbkenn As String
    Dim iRet As Integer
    Dim sOrder As String
    Dim sSQL As String
    Dim i As Integer
    Dim lrow As Long
    Dim ctmp As String
    
    Select Case Index
        Case Is = 0     '** ermitteln *
        
            lAusgewählt = 0
            
            Tabcheck "KB"
            FormatGridOverTablay "KB"

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
            
            If Option1(0).Value = True Then
                sOrder = " Order by ertrag desc"
            ElseIf Option1(1).Value = True Then
                sOrder = " Order by preis desc" 'Umsatz
            ElseIf Option1(2).Value = True Then
                sOrder = " Order by menge desc" 'Menge
            ElseIf Option1(3).Value = True Then
                sOrder = " Order by KUNDNR asc" 'Bedienernummer
            ElseIf Option1(9).Value = True Then
                sOrder = " Order by name" 'Bedienername
            End If
            
            GridFuellen "Select * from KUNTOPI " & sOrder
            
            ermittlespalten
            Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
            
        Case 1 'Drucken
            KUNTOPIupdate
            Drucke154
        Case 2 'kundendaten
            lrow = Val(MSHFLEX1.Row)
            If lrow > 0 Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = SpaltennummerKUNDNR
                gcKundenNr = MSHFLEX1.Text
                iKasse = 2
                frmWKL13.Show 1
            End If
        Case 3
        
            KUNTOPIupdate
            Text1(9).Text = ""
            Label1(13).Caption = "Umsatz"
            If Text1(8).Text <> "" Then
                If IsNumeric(Text1(8).Text) Then
                    sSQL = "Update KUNTOPI SET PROV = Preis * '" & CDbl(Text1(8).Text) & "' / 100 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from KUNTOPI " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        Case 4
            gsZSpalte = "Kundnr"
            gstab = "KB"
            frmWKL36.Show 1
            'fertig
        Case 5
            KUNTOPIupdate
            Text1(9).Text = ""
            Label1(13).Caption = "Rohertrag"
        
            If Text1(8).Text <> "" Then
                If IsNumeric(Text1(8).Text) Then
                    sSQL = "Update KUNTOPI SET PROV = Ertrag * '" & CDbl(Text1(8).Text) & "' / 100 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from KUNTOPI  " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        
        Case 6
            If Command1(6).Caption = "Zurück" Then
            
                Frame1.Visible = True
                Frame5.Visible = False
                MSHFLEX1.Visible = False
                
                Text1(9).Visible = False
                Text1(8).Visible = False
                Label1(11).Visible = False
                Command1(3).Visible = False
                Command1(5).Visible = False
                Command1(1).Visible = False
                Command1(7).Visible = False
                Command1(8).Visible = False
                
                Command1(6).Caption = "Leeren"
            
            Else
                List1.Clear
                List2.Clear
                List3.Clear
                List4.Clear
                List1.Visible = False
                List2.Visible = False
                List3.Visible = False
                List4.Visible = False
                
                For i = 0 To 8
                    Text1(i).Text = ""
                Next i
                füllefil cboFil
                
            End If
        Case 7
            KUNTOPIupdate
            Text1(8).Text = ""
            Label1(13).Caption = "Euro pro Stück"
            If Text1(9).Text <> "" Then
                If IsNumeric(Text1(9).Text) Then
                    sSQL = "Update KUNTOPI SET PROV = Menge * '" & CDbl(Text1(9).Text) & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    GridFuellen "Select * from KUNTOPI  " & sOrder
                    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
                End If
            End If
        Case 8
            KUNTOPIupdate
            Excel154
            ExcelExport "KUNTOPIPRINT", gdBase
        Case 9 'historie
            lrow = Val(MSHFLEX1.Row)
            If lrow > 0 Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = SpaltennummerKUNDNR
                gckundnr = MSHFLEX1.Text
                
                gckundnr = Trim$(gckundnr)
                gsARTNR = ""
                
                If gckundnr <> "" Then
                    frmWKL74.Show 1
                End If
            End If
            
            
        Case 10
            If Command1(10).Caption = "alle zurücksetzen" Then
            
                SchalteKunden (2)
                Command1(10).Caption = "alle auswählen"
            ElseIf Command1(10).Caption = "alle auswählen" Then
                SchalteKunden (3)
                Command1(10).Caption = "alle zurücksetzen"
            End If
        Case 11
            Screen.MousePointer = 0
            
            gsBackcolor = Label4(0).BackColor
            gsForecolor = Label4(0).ForeColor
            gsKundenFarbe = Label4(0).Tag
            
            frmWKL65.Show 1
            
            Label4(0).BackColor = gsBackcolor
            Label4(0).ForeColor = gsForecolor
            Label4(0).Tag = gsKundenFarbe
            If gsKundenFarbe <> "" Then
                Label4(0).Caption = "Farbauswahl"
            Else
                Label4(0).Caption = "alle Farben"
            End If
        Case 12
            ctmp = Trim$(Label4(0).Tag)
            If ctmp <> "" Then
                cFarbkenn = ermFarbeKU(ctmp)
            Else
                cFarbkenn = "alle Farben"
                SchalteKunden (2)
                Exit Sub
                ctmp = "0"
            End If
            
            If cFarbkenn = "" Then cFarbkenn = "ohne Kennzeichen"
            
            iRet = MsgBox("Möchten Sie jetzt alle Kunden aus der Tabelle mit dem Farbkennzeichen '" & cFarbkenn & "' zurücksetzen?", vbYesNo + vbQuestion + vbDefaultButton2, "Zentrale Frage:")
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
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
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
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = SpaltennummerAUSG
    
    If iSchaltung = 3 Then
        lAusgewählt = 0
    End If
    
    If iSchaltung = 2 Then
        lAusgewählt = 0
    End If
    
    MSHFLEX1.Redraw = False
    For lrow = 2 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If iSchaltung = 2 Then
            MSHFLEX1.Text = ""
            
        End If
        If iSchaltung = 4 Then
        
            'ja aber hat der kunden bestimmte farbe
'            anzeige "normal", lrow - 1 & "...", lblAnzeige
                
            ctmp = Trim$(Label4(0).Tag)
            If ctmp = "" Then ctmp = "0"
            
            MSHFLEX1.Col = SpaltennummerKUNDNR
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
        
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label20
    ElseIf lAusgewählt = 1 Then
        
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label20
    Else
        
        anzeige "normal", "", Label20
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
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "KUNDNR"
                SpaltennummerKUNDNR = i
            Case Is = "AUSG"
                SpaltennummerAUSG = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Drucke154()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim cProv       As String
    Dim cproz       As String
    Dim sOrder      As String
    Dim sSorti      As String
    Dim dGumsatz    As Double
    
    cDatVon = Text2(0).Value
    cDatBis = Text2(1).Value
    dGumsatz = CDbl(Label1(14).Caption)
    
    If Text1(8).Text <> "" Then
        cProv = Label1(13).Caption
        cproz = Text1(8).Text & " % vom"
    End If
    
    If Text1(9).Text <> "" Then
        cproz = Text1(9).Text & " € pro Stück"
    End If
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblAnzeige
    
    If Option1(0).Value = True Then
        sOrder = " Order by ertrag desc"
        sSorti = "sortiert nach Rohertrag"
    ElseIf Option1(1).Value = True Then
        sOrder = " Order by preis desc" 'Umsatz
        sSorti = "sortiert nach Umsatz"
    ElseIf Option1(2).Value = True Then
        sOrder = " Order by menge desc" 'Menge
        sSorti = "sortiert nach Stückzahlen"
    ElseIf Option1(3).Value = True Then
        sOrder = " Order by Kundnr asc" 'Bediener
        sSorti = "sortiert nach Kundennummer"
    ElseIf Option1(9).Value = True Then
        sOrder = " Order by name " 'Bedienername
        sSorti = "sortiert nach Kundenname"
    End If
    
    loeschNEW "KUNTOPIPRINT", gdBase
    CreateTableT2 "KUNTOPIPRINT", gdBase
            
    cSQL = "Insert into KUNTOPIPRINT Select * from KUNTOPI where ausg = true " & sOrder
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Kopf154", gdBase
    CreateTableT2 "KOPF154", gdBase
    
    cSQL = "Insert into KOPF154 (DATVON,DATBIS,Prov,Proz,Sortierung,Gumsatz) values ("
    cSQL = cSQL & " '" & cDatVon & "'  "
    cSQL = cSQL & ", '" & cDatBis & "'  "
    cSQL = cSQL & ", '" & cProv & "'  "
    cSQL = cSQL & ", '" & cproz & "'  "
    cSQL = cSQL & ", '" & sSorti & "'  "
    cSQL = cSQL & ", '" & dGumsatz & "'  "
    cSQL = cSQL & "  ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Kunden set awm = 0 where awm is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KUNTOPIPRINT inner join Kunden on KUNTOPIPRINT.Kundnr = Kunden.Kundnr "
    cSQL = cSQL & " SET KUNTOPIPRINT.farbnr = val(Kunden.awm) "
    gdBase.Execute cSQL, dbFailOnError
    
    BringFarbeInsSpielforKunden "KUNTOPIPRINT", gdBase
    
    reportbildschirm "", "aZEN154a"
    
    anzeige "normal", "Fertig", lblAnzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke154"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Excel154()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim sOrder      As String
    Dim sSorti      As String
    
    
    
    Screen.MousePointer = 11
    
    
    If Option1(0).Value = True Then
        sOrder = " Order by ertrag desc"
        sSorti = "sortiert nach Rohertrag"
    ElseIf Option1(1).Value = True Then
        sOrder = " Order by preis desc" 'Umsatz
        sSorti = "sortiert nach Umsatz"
    ElseIf Option1(2).Value = True Then
        sOrder = " Order by menge desc" 'Menge
        sSorti = "sortiert nach Stückzahlen"
    ElseIf Option1(3).Value = True Then
        sOrder = " Order by Kundnr asc" 'Bediener
        sSorti = "sortiert nach Kundennummer"
    ElseIf Option1(9).Value = True Then
        sOrder = " Order by name " 'Bedienername
        sSorti = "sortiert nach Kundenname"
    End If
    
    cSQL = "Update KUNTOPI inner join Kunden on KUNTOPI.kundnr = Kunden.kundnr"
    cSQL = cSQL & " set KUNTOPI.KUVORNAME = Kunden.vorname "
    cSQL = cSQL & " , KUNTOPI.KUTITEL = Kunden.titel "
    cSQL = cSQL & " , KUNTOPI.KUPLZ = Kunden.plz "
    cSQL = cSQL & " , KUNTOPI.KUSTADT = Kunden.stadt "
    cSQL = cSQL & " , KUNTOPI.KUANREDE = Kunden.anrede "
    cSQL = cSQL & " , KUNTOPI.KUGESCHLECHT = Kunden.geschlecht "
    cSQL = cSQL & " , KUNTOPI.KUFIRMA = Kunden.firma "
    cSQL = cSQL & " , KUNTOPI.KUSTRASSE = Kunden.strasse "
    gdBase.Execute cSQL, dbFailOnError
            
    loeschNEW "KUNTOPIPRINT", gdBase
    CreateTableT2 "KUNTOPIPRINT", gdBase
            
    cSQL = "Insert into KUNTOPIPRINT Select * from KUNTOPI where ausg = true " & sOrder
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Excel154"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
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
    Dim cAWM        As String
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    Screen.MousePointer = 11
    
    With MSHFLEX1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
    
        anzeige "normal", "Es werden " & lMax & " Kunden angezeigt...", lblAnzeige
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
                        Case Is = "Umsatz", "NS", "Rohertrag", "K. Schnitt", "Provision"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                        Case Is = "Kund Nr."
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
            
                            cAWM = ""
                            If sWert <> "" Then
                                cAWM = WhatIsAwmKU(sWert)
                            Else
                                
                            End If
                            
                            If cAWM = "" Then cAWM = "0"
                            FaerbenFlexHKunde cAWM, MSHFLEX1, i, lrow
                        Case Is = "ausg"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                If rsrs(sSpaltenbez(i)) = True Then
                                    sWert = "X"
                                Else
                                    sWert = ""
                                End If
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                        
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
        
        Frame1.Visible = False
        Frame5.Visible = True
        
        Text1(9).Visible = True
        Text1(8).Visible = True
        Label1(11).Visible = True
        Command1(3).Visible = True
        Command1(5).Visible = True
        Command1(1).Visible = True
        Command1(7).Visible = True
        Command1(8).Visible = True
        
        Command1(6).Caption = "Zurück"
    Else
        Frame1.Visible = True
        Frame5.Visible = False
        Command1(6).Caption = "Leeren"
        
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblAnzeige
        anzeige "normal", "", Label20
        
        Exit Sub

    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.8
    Next i
        
    rsrs.Close
    
    lAusgewählt = lMax
    
    If lMax > 1 Then
        anzeige "normal", lMax & " Kunden wurden ermittelt.", lblAnzeige
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label20
    ElseIf lMax = 1 Then
        anzeige "normal", lMax & " Kunde wurden ermittelt.", lblAnzeige
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label20
    Else
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblAnzeige
        anzeige "normal", "", Label20
    End If
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    .Redraw = True
    .Visible = True
    End With
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
   Resume Next
End Sub
Private Sub ermitteln()
On Error GoTo LOKAL_ERROR

    Dim cVon            As String
    Dim cBis            As String
    Dim lVon            As Long
    Dim lBis            As Long
'    Dim iFil            As Integer
    Dim dProz           As Double
    Dim corder          As String
    Dim i               As Integer
    Dim bAnd            As Boolean
    Dim ctmp            As String
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsb             As Recordset
    Dim iAnzahlKunden   As Integer

    'vorbereitung
    If Text2(0).Value <> "" Then
        cVon = Text2(0).Value
    Else
        cVon = DateValue(Now) - 30
        Text2(0).Value = DateValue(Now) - 30
    End If
    
    If Text2(1).Value <> "" Then
        cBis = Text2(1).Value
    Else
        cBis = DateValue(Now)
        Text2(1).Value = DateValue(Now)
    End If
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))

'    If cboFil.Text = "alle Filialen" Then
'        iFil = 0
'    Else
'        iFil = CInt(Left$(cboFil.Text, 3))
'    End If
'
    'Vorbereitung ende
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", lblAnzeige
    
    sSQL = "Select "
    sSQL = sSQL & "  Sum(preis) as Maxi "
    sSQL = sSQL & " from Kassjour A "
    sSQL = sSQL & " where A.adate between  " & cVon & " And " & cBis
    sSQL = sSQL & " and UMS_OK = 'J' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!maxi) Then
            Label1(14).Caption = Format(rsrs!maxi, "########0.00")
        End If
        
    End If
    rsrs.Close
        
    loeschNEW "KUNB", gdBase
    CreateTableT2 "KUNB", gdBase
    
    sSQL = "Insert into KUNB Select "
    sSQL = sSQL & "  A.preis "
    sSQL = sSQL & " ,A.menge "
    sSQL = sSQL & " , A.artnr "
    sSQL = sSQL & " , A.ekpr "
    sSQL = sSQL & " , A.VKpr "
    sSQL = sSQL & " , A.mwst "
    sSQL = sSQL & " , A.KUNDNR  "
    sSQL = sSQL & " , A.BELEGNR  "
    sSQL = sSQL & " , A.adate  "
    sSQL = sSQL & " from Kassjour A "
    sSQL = sSQL & " where A.adate between  " & cVon & " And " & cBis
    


'    If iFil = 0 Then
'
'    Else
'       sSQL = sSQL & " and A.filiale = " & iFil
'    End If

     bAnd = True
     
    If Check2.Value = vbChecked Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & "  A.UMS_OK = 'J' "
    End If
    
    If Check5.Value = vbChecked Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & "  (A.VKPR*a.menge) <=  A.Preis"
    End If

    'LiefNr
    ctmp = Trim$(Text1(2).Text)
    If ctmp <> "" Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & " A.LINR = " & ctmp & " "
        bAnd = True
    End If
    
    'ArtNr oder EAN
    ctmp = Trim$(Text1(6).Text)
    If ctmp <> "" Then
    
        If bAnd Then
            sSQL = sSQL & " and "
        End If
       
        If Len(ctmp) <= 6 Then
            'KISS-ArtNr
            sSQL = sSQL & " A.ARTNR = " & ctmp & " "
            bAnd = True
            
        ElseIf Len(ctmp) = 8 Then
            'KISS-ArtNr als Barcode oder echter EAN-8
            If Left$(ctmp, 1) = "2" Or Left$(ctmp, 1) = "0" Then
                ctmp = Mid$(ctmp, 2, 6)
                sSQL = sSQL & " A.ARTNR = " & ctmp & " "
                bAnd = True
            Else
                sSQL = sSQL & " ( A.EAN = '" & ctmp & "' "
                sSQL = sSQL & " or A.EAN2 = '" & ctmp & "' "
                sSQL = sSQL & " or A.EAN3 = '" & ctmp & "' ) "
                bAnd = True
            End If
        Else
            'Irgendwas anderes für die EAN-Felder
            sSQL = sSQL & " ( A.EAN = '" & ctmp & "' "
            sSQL = sSQL & " or A.EAN2 = '" & ctmp & "' "
            sSQL = sSQL & " or A.EAN3 = '" & ctmp & "' ) "
            bAnd = True
        End If
    End If
    
    

    'Linie
    If List3.Visible = True And List3.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( lpz=" & Mid$(List3.list(0), 1, InStr(1, List3.list(0), " "))
        For i = 1 To List3.ListCount - 1
            sSQL = sSQL & " or lpz=" & Mid$(List3.list(i), 1, InStr(1, List3.list(i), " "))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'Linie
        ctmp = Trim$(Text1(5).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.LPZ = " & ctmp & " "
            bAnd = True
        End If
    End If

    'Marke
    ctmp = Trim$(Text1(7).Text)
    If ctmp <> "" Then
        If LoeseMarkenInArtnr(ctmp) Then
            If bAnd Then
                sSQL = sSQL & "and "
            End If
            sSQL = sSQL & " A.artnr in(Select artnr from MY" & srechnertab & ")"
            bAnd = True
        End If
    End If

    'LiefBestNr
    ctmp = Trim$(Text1(3).Text)
    If ctmp <> "" Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If
        sSQL = sSQL & "A.LIBESNR like '" & ctmp & "*' "
        bAnd = True
    End If

    'AGN
    If List1.Visible = True And List1.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( agn=" & Mid$(List1.list(0), 1, InStr(1, List1.list(0), " "))
        For i = 1 To List1.ListCount - 1
            sSQL = sSQL & " or agn=" & Mid$(List1.list(i), 1, InStr(1, List1.list(i), " "))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'agn
        ctmp = Trim$(Text1(4).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.AGN = " & ctmp & " "
            bAnd = True
        End If
    End If
    
    'bediener
    If List4.Visible = True And List4.ListCount > 0 Then
        If bAnd Then
            sSQL = sSQL & " and "
        End If

        sSQL = sSQL & "( bediener=" & Mid$(List4.list(0), 1, InStr(1, List4.list(0), " "))
        For i = 1 To List4.ListCount - 1
            sSQL = sSQL & " or bediener=" & Mid$(List4.list(i), 1, InStr(1, List4.list(i), " "))
        Next i
        sSQL = sSQL & " ) "
        bAnd = True
    Else
        'Kunde
        ctmp = Trim$(Text1(1).Text)
        If ctmp <> "" Then
            If bAnd Then
                sSQL = sSQL & " and "
            End If
            sSQL = sSQL & "A.bediener= " & ctmp & " "
            bAnd = True
        End If
    End If
    
'    MsgBox sSQL
    gdBase.Execute sSQL, dbFailOnError
    
    'Farben
    ctmp = Trim$(Label4(32).Tag)
    If ctmp <> "" Then
        sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
        sSQL = sSQL & " where artikel.awm <> '" & ctmp & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check3.Value = vbChecked Then
        sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
        sSQL = sSQL & " where artikel.BONUS_OK = 'N'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check4.Value = vbChecked Then
        sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
        sSQL = sSQL & " where artikel.VKPR * Kunb.menge > kunb.preis "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
'    If Check5.Value = vbChecked Then
'        sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
'        sSQL = sSQL & " where artikel.KVKPR1 * Kunb.menge > kunb.preis "
'        gdbase.Execute sSQL, dbFailOnError
'    End If
    
'    If Check6.Value = vbChecked Then
'
'        If Text1(10).Text = "" Then Text1(10).Text = "0"
'        dProz = Text1(10).Text
'        If dProz > 0 Then
'            sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
'            sSQL = sSQL & " where 100 -( kunb.preis * (artikel.KVKPR1 * Kunb.menge)/100) >= '" & dProz & "'"
'            gdbase.Execute sSQL, dbFailOnError
'        End If
'    End If
    
    'PGN
    If List2.Visible = True And List2.ListCount > 0 Then
        For i = 1 To List2.ListCount - 1
            sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
            sSQL = sSQL & " where artikel.pgn <> " & Trim$(Left$(List2.list(i), 2))
            gdBase.Execute sSQL, dbFailOnError
        Next i
    Else
        'Pgn
        ctmp = Trim$(Text1(0).Text)
        If ctmp <> "" Then
            sSQL = "delete KUNB.* from KUNB inner join ARTIKEL on ARTIKEL.artnr = KUNB.artnr "
            sSQL = sSQL & " where artikel.pgn <> " & ctmp
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    'Warengruppen
    If Option2(0).Value = True Then
        'einschließen
    
        
    ElseIf Option2(1).Value = True Then
        'ausschließen
        
        If checkwarengru Then

            
            sSQL = "Delete from KUNB where artnr in(select artnr from warengru) "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
    ElseIf Option2(2).Value = True Then
        'nur Warengruppen
        
        If checkwarengru Then
            sSQL = "Delete from KUNB where artnr not in(select artnr from warengru) "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    'End Warengruppen
    
    
    'Gutscheine
    If Check1.Value = vbChecked Then
        sSQL = "delete from KUNB where artnr = 666666 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    'End Gutscheine
    sSQL = " Create index  MWST on KUNB(MWST) "
    gdBase.Execute sSQL, dbFailOnError

    
    sSQL = "Update KUNB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStV & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStV & "))* 100)"
    sSQL = sSQL & " where MWST = 'V' "
    sSQL = sSQL & " and PREIS <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStE & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStE & "))* 100)"
    sSQL = sSQL & " where MWST = 'E' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNB "
    sSQL = sSQL & " set "
    sSQL = sSQL & " ENS1 = ((((Preis/(100 + " & gdMWStO & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStO & "))* 100)"
    sSQL = sSQL & " where MWST = 'O' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update KUNB set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNB set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNB set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "KUNUMSATZ", gdBase
    
    sSQL = " Create index  kundnr on KUNB(kundnr) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select KUNDNR, sum(rertrag) as mrertrag,sum(preis) as mpreis,sum(menge) as mmenge ,avg(ens1) as ens into KunUMSATZ "
    sSQL = sSQL & " from KUNB group by KUNDNR "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "KUNTOPI", gdBase
    CreateTableT2 "KUNTOPI", gdBase
    
    sSQL = "Insert into KUNTOPI SELECT  KUNDNR, mrertrag as ertrag, mmenge as menge "
    sSQL = sSQL & " , mpreis as preis , ens  "
    sSQL = sSQL & " from KUNUMSATZ "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Create index  BELEGNR on KUNB(BELEGNR) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Create index  adate on KUNB(adate) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update KUNTOPI SET PROV = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNTOPI SET ausg = true "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNTOPI inner join Kunden on KUNTOPI.Kundnr = Kunden.kundnr "
    sSQL = sSQL & " SET KUNTOPI.name = Kunden.name "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "", lblAnzeige
    
    Screen.MousePointer = 0
   
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermitteln"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub

Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
    
    Case 11
        gsHelpstring = "Kundenbeteiligung"
        frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
       
    WKLatPositionieren
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Screen.MousePointer = 11
    
    Text2(1).Value = Format(Date, "DD.MM.YY")
    Text2(0).Value = Format(Date - 7, "DD.MM.YY")
    
    If NewTableSuchenDBKombi("C145E", gdApp) Then
        voreinstellungladen
    End If
    
    füllefil cboFil
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Set rsrs = gdApp.OpenRecordset("C145E")
    
    If Not rsrs.EOF Then
        
        Text2(0).Value = rsrs!Von
        Text2(1).Value = rsrs!Bis
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    Dim lVon As Long
    Dim lBis As Long
    
    loeschNEW "C145E", gdApp
    CreateTableT2 "C145E", gdApp
    
    lVon = Text2(0).Value
    lBis = Text2(1).Value
    
    sSQL = "Insert into C145E (von,bis) "
    sSQL = sSQL & " values (" & lVon & " ," & lBis & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label4_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 32 Then
    Label4(Index).Caption = "alle Farben"
    Label4(Index).Tag = ""
    Label4(Index).BackColor = Label1(1).BackColor
    Label4(Index).ForeColor = Label1(1).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR

    
    
    If MSHFLEX1.Row = 1 Then
        sortierenHGrid MSHFLEX1
    Else
        MSHFLEX1.Col = SpaltennummerAUSG
        If MSHFLEX1.Text = "X" Then
            MSHFLEX1.Text = ""
            lAusgewählt = lAusgewählt - 1
        Else
            MSHFLEX1.Text = "X"
            lAusgewählt = lAusgewählt + 1
        End If
        
        If lAusgewählt > 1 Then
            anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label20
        ElseIf lAusgewählt = 1 Then
            anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label20
        Else
            anzeige "normal", "", Label20
        End If
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 4     'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text2(0).Value = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text2(1).Value = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text2(0).Value = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text2(1).Value = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case 2
                        Text2(1).Value = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    Case Else
                        Text2(1).Value = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
        Case Is = 5     'ak monat
            Text2(0).Value = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM.YY")
        Case Is = 8     'vorjahrzr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM") & "." & Year(Now) - 1
        Case Is = 6     'vorjahr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Text2(1).Value = "31.12." & Year(Now) - 1
        Case Is = 7     'ak jahr
            Text2(0).Value = Format("01.01." & Year(DateValue(Now)), "DD.MM.YY")
            Text2(1).Value = Format(DateValue(Now), "DD.MM.YY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

Dim sAuswahlfeld As String
Dim ctmp As String
Dim lcount As Long

If KeyCode = vbKeyF2 Then
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    Select Case Index
        Case Is = 0
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "PGN"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List2.Visible = False
                List2.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List2.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List2.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List2.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), 2)
                        End If
                        
                    End If
                Next lcount
            End If
        Case Is = 1
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "BED"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List4.Visible = False
                List4.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List4.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List4.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List4.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                        End If
                        
                    End If
                Next lcount
                
            End If
        Case Is = 2
            gF2Prompt.bMultiple = False
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
            End If
        Case Is = 4
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "AGN"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                
                
                List1.Visible = False
                List1.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List1.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List1.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List1.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                        End If
                        
                    End If
                Next lcount

            End If


        Case 5
            ctmp = Text1(7).Text
            ctmp = Trim$(ctmp)
            If ctmp = "" Then
                ctmp = Text1(2).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    anzeige "Rot", "Bitte einen Lieferanten oder eine Marke angeben!", lblAnzeige
                    Text1(7).SetFocus
                    Exit Sub
                Else
                    sAuswahlfeld = "LINR"
                End If
            Else
                sAuswahlfeld = "MARKE"
            End If
            
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "LPZ"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cEsFeld = sAuswahlfeld
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount) & Space(50) & Right(gF2Prompt.cArray(lcount), 6)
                        End If
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 3)
                        End If
                    End If
                Next lcount
            End If

        Case Is = 7
            gF2Prompt.cFeld = "MARKE"
            
            ctmp = Text1(2).Text 'Linr eventuell
            gF2Prompt.cEsFeld = ctmp
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
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
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
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
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

loeschNEW "KUNTOPI", gdBase
loeschNEW "KUNTOPIPRINT", gdBase
voreinstellungspeichern
LogtoEnd Me

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
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
        Case 6, 1, 2, 5, 4, 0 'ARTNR, EAN, LIEFNR, ARTGRU,linie
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 8, 10 'Proz
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 3, 7       'BEZEICH, LIBESNR
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß#"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            'alle Zeichen erlaubt
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kundenbeteiligung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



