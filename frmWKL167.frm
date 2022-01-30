VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmWKL167 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Preiskalkulation"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4095
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   14
         Left            =   9600
         TabIndex        =   22
         Top             =   6000
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
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   13
         Left            =   9600
         TabIndex        =   21
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
         Caption         =   "Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   12
         Left            =   9600
         TabIndex        =   20
         Top             =   720
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   4815
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   18
         FixedCols       =   2
         ForeColorSel    =   8454143
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.Label Label2 
         Caption         =   "Artikel"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   5535
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   17
         Left            =   9600
         TabIndex        =   24
         Top             =   2520
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
         Caption         =   "Kalkulieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   10
         Left            =   9600
         TabIndex        =   17
         Top             =   720
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   9
         Left            =   9600
         TabIndex        =   16
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
         Caption         =   "Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   8
         Left            =   9600
         TabIndex        =   15
         Top             =   6000
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4815
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   18
         FixedCols       =   2
         ForeColorSel    =   8454143
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.Label Label2 
         Caption         =   "Lieferant"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   12
      Top             =   3960
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
      Caption         =   "AGN "
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   11
      Top             =   4800
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
      Caption         =   "Kalkulieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   10
      Top             =   1920
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
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   1320
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
      Caption         =   "Artikel"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin sevCommand3.Command Command4 
      Height          =   355
      Index           =   16
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   375
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Index           =   6
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   1575
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   3
      Top             =   240
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
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   3
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   18
      FixedCols       =   2
      ForeColorSel    =   8454143
      FocusRect       =   0
      HighLight       =   2
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
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   18
      Left            =   9600
      TabIndex        =   26
      Top             =   7200
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
      Caption         =   "Klassische Ansicht"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   0
      Left            =   10800
      TabIndex        =   27
      Top             =   240
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
      Picture         =   "frmWKL167.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   7
      Left            =   10320
      TabIndex        =   28
      Top             =   240
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
      Picture         =   "frmWKL167.frx":0692
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   360
      Index           =   15
      Left            =   9840
      TabIndex        =   29
      Top             =   240
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
      Picture         =   "frmWKL167.frx":0D24
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Lieferant"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1095
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Preiskalkulation"
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
      TabIndex        =   1
      Top             =   0
      Width           =   6615
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
End
Attribute VB_Name = "frmWKL167"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerLINR As Byte
Dim SpaltennummerNS As Byte
Dim SpaltennummerAGN As Byte
Dim SpaltennummerNSAGN As Byte
Dim SpaltennummerNewKVK As Byte
Dim SpaltennummerArtnr As Byte
Private Sub speicherDieWahl()
On Error GoTo LOKAL_ERROR

    '1. Nettospanne speichern
    '2. Preise neukalkulieren
    '3. Etiketten bereitstellen
    
    Dim sSQL        As String
    Dim lLinr       As Long
    Dim rsrs        As Recordset
    Dim dNS         As Double
    Dim lcount      As Long
    
    Screen.MousePointer = 11
    
'    anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt...", lblanzeige
    

    anzeige "normal", "", lblAnzeige
    Screen.MousePointer = 11

    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1

        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = SpaltennummerLINR
        lLinr = Val(MSFlexGrid1.Text)
        MSFlexGrid1.Col = SpaltennummerNS
        
        dNS = 0
        If MSFlexGrid1.Text <> "" Then
            If IsNumeric(MSFlexGrid1.Text) Then
                dNS = CDbl(MSFlexGrid1.Text)
                MSFlexGrid1.Text = Format(dNS, "####0.00")
            End If
        End If

        If dNS > 0 Then
            sSQL = "Update  PREISKALKLINR set Ns = '" & dNS & "'"
            sSQL = sSQL & " where Linr = " & lLinr
            gdBase.Execute sSQL, dbFailOnError
        End If

    Next lcount

    MSFlexGrid1.Redraw = True

    Screen.MousePointer = 0


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDieWahl"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicherDieWahlAgn(cLinr As String)
On Error GoTo LOKAL_ERROR

    '1. Nettospanne speichern
    '2. Preise neukalkulieren
    '3. Etiketten bereitstellen
    
    '4. alles für rückgängig vorbereiten
    
    Dim sSQL        As String
    Dim lagn        As Long
    Dim rsrs        As Recordset
    Dim dNS         As Double
    Dim lcount      As Long
    
    Screen.MousePointer = 11
    
'    anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt...", lblanzeige

    anzeige "normal", "", lblAnzeige
    Screen.MousePointer = 11

    MSFlexGrid2.Redraw = False
    For lcount = 1 To MSFlexGrid2.Rows - 1

        MSFlexGrid2.Row = lcount
        MSFlexGrid2.Col = SpaltennummerAGN
        lagn = Val(MSFlexGrid2.Text)
        MSFlexGrid2.Col = SpaltennummerNSAGN
        
        dNS = 0
        If MSFlexGrid2.Text <> "" Then
            If IsNumeric(MSFlexGrid2.Text) Then
                dNS = CDbl(MSFlexGrid2.Text)
                MSFlexGrid2.Text = Format(dNS, "####0.00")
            End If
        End If

        If dNS > 0 Then
            sSQL = "Update  PREISKALKAGN set Ns = '" & dNS & "'"
            sSQL = sSQL & " where Linr = " & cLinr
            sSQL = sSQL & " and AGN = " & lagn
            gdBase.Execute sSQL, dbFailOnError
        End If

    Next lcount

    MSFlexGrid2.Redraw = True

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDieWahlAgn"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Preise_Speichern()
On Error GoTo LOKAL_ERROR

    Dim i As Long
    Dim cPreis As String
    Dim cartnrzuSpeichern As String
    
    anzeige "normal", "", lblAnzeige
    
    For i = 1 To MSFlexGrid3.Rows - 1
    
        cPreis = MSFlexGrid3.TextMatrix(i, SpaltennummerNewKVK)
        cartnrzuSpeichern = MSFlexGrid3.TextMatrix(i, SpaltennummerArtnr)
        If cartnrzuSpeichern <> "" Then
            If Val(cartnrzuSpeichern) > 0 Then
                If IsNumeric(cPreis) Then
                
                    anzeige "normal", "neuer Preis " & cPreis & " wird gespeichert " & cartnrzuSpeichern, lblAnzeige
                    Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel Kalk neu", "KVKPR1"
        
                    schreibeWKEtidru cartnrzuSpeichern, ermBESTAND(cartnrzuSpeichern), CLng(gcFilNr)
                End If
            End If
        End If
    
    Next i
    
    anzeige "normal", "Fertig! " & MSFlexGrid3.Rows - 2 & " Artikelpreise gespeichert", lblAnzeige
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Preise_Speichern"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        Case 0
            gsZSpalte = "LINR"
            gstab = "KALKLINR"
            frmWKL36.Show 1
            'fertig
        Case 1
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gcSuch = "LINR" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL70.Show 1
                Me.Refresh
                gcSuch = ""
            End If
        Case 2
            speicherDieWahl
        Case 3
            Unload frmWKL167
        Case 4 'Kalkulieren
        
            Command4_Click 2
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
            
                If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerNS)) > 0 Then
                    Frame2.Visible = True
                    ZeigeKALKARTIKEL MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR), "", MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerNS)
                    Me.Refresh
                Else
                    anzeige "rot", "Nettospannenangabe fehlt...", lblAnzeige
                End If
                
            End If
        Case 5
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gcSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                Label2(0).Caption = gcSuch
                Frame1.Visible = True
                ZeigeKALK_AGN gcSuch
                
                Me.Refresh
                gcSuch = ""
            End If
            
        Case 6
            ZeigeKALKLinr
        Case 7
            gsZSpalte = "AGN"
            gstab = "KALKAGN"
            frmWKL36.Show 1
            'fertig
        Case 8
            Frame1.Visible = False
        Case 10
            speicherDieWahlAgn Label2(0).Caption
        Case 11
            gsHelpstring = "Preiskalkulation"
            frmWKL110.Show 1
        Case 12
            'neue Preise speichern
            Preise_Speichern
        Case 14
            Frame2.Visible = False
        Case 15
            gsZSpalte = "ARTNR"
            gstab = "KALKARTIKEL"
            frmWKL36.Show 1
            'fertig
        Case 16
            Text1_KeyUp 4, vbKeyF2, 0
            
        Case 17 'Kalkulieren Agn
            If Val(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, SpaltennummerAGN)) > 0 Then
                If Val(MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, SpaltennummerNSAGN)) > 0 Then
                    Frame2.Visible = True
                    
                    
                    ZeigeKALKARTIKEL Label2(0).Caption, MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, SpaltennummerAGN), MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, SpaltennummerNSAGN)
                    Me.Refresh
                Else
                    anzeige "rot", "Nettospannenangabe fehlt...", lblAnzeige
                End If
            End If
        Case 18
            frmWKL35.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    WKL167Positionieren
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    anzeige "normal", "", lblAnzeige
    
    Me.Refresh
    
    If NewTableSuchenDBKombi("PREISKALKLINR", gdBase) = False Then
        CreateTableT2 "PREISKALKLINR", gdBase
        Fuelle_PreisKalkLinr
    Else
        If Datendrin("PREISKALKLINR", gdBase) = False Then
            Fuelle_PreisKalkLinr
        End If
    End If
    
    If NewTableSuchenDBKombi("PREISKALKAGN", gdBase) = False Then
        CreateTableT2 "PREISKALKAGN", gdBase
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL167Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 1200
    Frame1.Left = 0
    Frame1.Height = 6615
    Frame1.Width = 11775
    Frame1.BorderStyle = 0
    
    Frame2.Top = 1200
    Frame2.Left = 0
    Frame2.Height = 6615
    Frame2.Width = 11775
    Frame2.BorderStyle = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL167Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
         Case Is = SpaltennummerNS
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                    
                    
                End If
            End If
     End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSFlexGrid2.Col
    lrow = MSFlexGrid2.Row
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
         Case Is = SpaltennummerNSAGN
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid2.Row = lrow
                MSFlexGrid2.Col = lcol
                cValid = MSFlexGrid2.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid2.Text = cValid
                    
                    
                End If
            End If
     End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_LeaveCell()
    On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_LeaveCell()
    On Error GoTo LOKAL_ERROR

    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    
    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSFlexGrid1.Rows Then
        lrow = lrow - 1
    End If
    
    If KeyCode = &H28 Or KeyCode = &H27 Or KeyCode = &H26 Or KeyCode = &H25 Or KeyCode = vbKeyF2 Then
        Exit Sub
    End If
    
    If iKeypress = 0 And KeyCode <> vbKeyBack Then
        
        If KeyCode <> 46 Then
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Text = ""
        End If
    End If
    iKeypress = iKeypress + 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    
    lrow = MSFlexGrid2.Row
    lcol = MSFlexGrid2.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSFlexGrid2.Rows Then
        lrow = lrow - 1
    End If
    
    If KeyCode = &H28 Or KeyCode = &H27 Or KeyCode = &H26 Or KeyCode = &H25 Or KeyCode = vbKeyF2 Then
        Exit Sub
    End If
    
    If iKeypress = 0 And KeyCode <> vbKeyBack Then
        
        If KeyCode <> 46 Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            MSFlexGrid2.Text = ""
        End If
    End If
    iKeypress = iKeypress + 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "LINR"
                SpaltennummerLINR = i
            Case Is = "NS"
                SpaltennummerNS = i
        End Select
    Next i
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermittlespalten2()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "AGN"
                SpaltennummerAGN = i
            Case Is = "NS"
                SpaltennummerNSAGN = i
        End Select
    Next i
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten2"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermittlespalten3()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "PREISNACH"
                SpaltennummerNewKVK = i
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
        End Select
    Next i
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten3"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex167()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim cSQL        As String
    
    anzeige "normal", "Lieferanten werden angezeigt...", lblAnzeige
   
    cSQL = "Select * from PREISKALKLINRZ order by LINR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
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
                            Case Is = "NS z.Zeit", "errechn. Umsatz", "hinterlegte NS", "Umsatz Hochrechnung"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                            

                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
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
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    anzeige "normal", "", lblAnzeige
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex167"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex167a()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim cSQL        As String
    
    anzeige "normal", "Artikelgruppen werden angezeigt...", lblAnzeige
   
    cSQL = "Select * from PREISKALKAGNZ order by AGN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid2
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
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
                            Case Is = "NS z.Zeit", "errechn. Umsatz", "hinterlegte NS", "Umsatz Hochrechnung"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                                
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
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
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    anzeige "normal", "", lblAnzeige
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex167a"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
        
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub FuellenMSFlex167b()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim cSQL        As String
    
    anzeige "normal", "Artikel werden angezeigt...", lblAnzeige
   
    cSQL = "Select * from PREISKALKART "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid3
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
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
                            Case Is = "KVK vor", "KVK neu"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                                
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
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
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    anzeige "normal", "", lblAnzeige
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex167b"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
        
    Fehlermeldung1
'    Resume Next
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
Private Sub Fuelle_PreisKalkLinr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lLinr       As Long
    Dim lAnz        As Long
   
    sSQL = "Insert into PreisKalkLinr select LINR from LISRT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKalkLinr Set NS = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Fuelle_PreisKalkLinr"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Fuelle_PreisKalkAGN(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lLinr       As Long
    Dim lAnz        As Long
   
    sSQL = "Insert into PreisKalkAGN select AGN , " & cLinr & " as linr from AGNDBF "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKalkAGN Set NS = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Fuelle_PreisKalkAGN"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermdurchNS(lLinr As Long, lagn As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim dNSPv       As Double
    Dim dNSPe       As Double
    Dim dNSPo       As Double
    
    ermdurchNS = 0
    
    loeschNEW "DNSPART", gdBase
    CreateTableT2 "DNSPART", gdBase
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) - (ARTIKEL.EKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) as NSP "
    sSQL = sSQL & "  from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and Artikel.EKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'V' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) - (ARTIKEL.EKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) as NSP "
    sSQL = sSQL & " from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and Artikel.EKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'E' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) - (ARTIKEL.EKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) as NSP "
    sSQL = sSQL & " from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and Artikel.EKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'O' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Select AVG(NSP) as Mittel from DNSPArt "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Mittel) Then
            ermdurchNS = rsrs!Mittel
        Else
            ermdurchNS = 0
        End If
        
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermdurchNS"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Function ermdurchNS_LEKPR(lLinr As Long, lagn As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim dNSPv       As Double
    Dim dNSPe       As Double
    Dim dNSPo       As Double
    
    ermdurchNS_LEKPR = 0
    
    loeschNEW "DNSPART", gdBase
    CreateTableT2 "DNSPART", gdBase
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) - (ARTLIEF.LEKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) as NSP "
    sSQL = sSQL & "  from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and ARTLIEF.LEKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'V' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStV & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) - (ARTLIEF.LEKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) as NSP "
    sSQL = sSQL & " from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and ARTLIEF.LEKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'E' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStE & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DNSPArt select ARTIKEL.Artnr, ((((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) - (ARTLIEF.LEKPR ))* 100) / ((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) as NSP "
    sSQL = sSQL & " from ARTIKEL inner join ARTLIEF "
    sSQL = sSQL & " on ARTIKEL.ARTNR = ARTLIEF.ARTNR where ARTLIEF.linr = " & lLinr
    sSQL = sSQL & " and ARTLIEF.LEKPR > 0 "
    sSQL = sSQL & " and Artikel.MWST = 'O' "
    sSQL = sSQL & " and ((ARTIKEL.KVKPR1 /(100 + " & gdMWStO & "))* 100) <> 0 "
    
    If lagn > 0 Then
        sSQL = sSQL & " and Artikel.AGN = " & lagn & " "
    End If
    
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Select AVG(NSP) as Mittel from DNSPArt "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Mittel) Then
            ermdurchNS_LEKPR = rsrs!Mittel
        Else
            ermdurchNS_LEKPR = 0
        End If
        
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermdurchNS_LEKPR"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub ZeigeKALKLinr()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim lLinr       As Long
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    
    anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt...", lblAnzeige
    
    MSFlexGrid1.Clear
    
    loeschNEW "PREISKALKLINRZ", gdBase
    CreateTableT2 "PREISKALKLINRZ", gdBase
   
    sSQL = " Insert into PREISKALKLINRZ select LINR, NS from PREISKALKLINR "
    
    If Text1(4).Text <> "" Then
        If IsNumeric(Text1(4).Text) Then
            sSQL = sSQL & " where Linr = " & Text1(4).Text
        End If
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PREISKALKLINRZ inner join LISRT on PREISKALKLINRZ.LINR = LISRT.LINR "
    sSQL = sSQL & " Set  PREISKALKLINRZ.LIEFBEZ = LISRT.LIEFBEZ "
    sSQL = sSQL & " ,  PREISKALKLINRZ.KUERZEL = LISRT.KUERZEL "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("PREISKALKLINRZ")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lLinr = 0
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
            End If

            anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt... " & lLinr, lblAnzeige
            
            rsrs.Edit
            
            If gsSpanne = "LEK" Then
                rsrs!dNS = ermdurchNS_LEKPR(lLinr, 0)
            ElseIf gsSpanne = "SEK" Then
                rsrs!dNS = ermdurchNS(lLinr, 0)
            End If
            
            rsrs!AnzArtikel = ermAnzArt(" where Linr = " & lLinr)
            rsrs.Update

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
    
    ZeigeLINRTAB
    
    If MSFlexGrid1.Visible = True Then
        MSFlexGrid1.Col = SpaltennummerNS
        MSFlexGrid1.Row = 2
        MSFlexGrid1.SetFocus
    End If
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKALKLinr"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKALK_AGN(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim lagn        As Long
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    
    anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt...", lblAnzeige
    
    MSFlexGrid2.Clear
    
    loeschNEW "PREISKALKAGNZ", gdBase
    CreateTableT2 "PREISKALKAGNZ", gdBase
    
    If DatendrinSQL("select * from PREISKALKAGN where Linr = " & cLinr, gdBase) = False Then
        Fuelle_PreisKalkAGN cLinr
    End If
   
    sSQL = " Insert into PREISKALKAGNZ select AGN, NS from PREISKALKAGN "
    sSQL = sSQL & " where Linr = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PREISKALKAGNZ inner join AGNDBF on PREISKALKAGNZ.AGN = AGNDBF.AGN "
    sSQL = sSQL & " Set  PREISKALKAGNZ.AGNBEZ = AGNDBF.AGTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("PREISKALKAGNZ")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lagn = 0
            If Not IsNull(rsrs!AGN) Then
                lagn = rsrs!AGN
            End If

            anzeige "normal", "durchschnittlich kalkulierte Nettospannen werden ermittelt... " & lagn, lblAnzeige
            
            rsrs.Edit
            If gsSpanne = "LEK" Then
                rsrs!dNS = ermdurchNS_LEKPR(CLng(cLinr), 0)
            ElseIf gsSpanne = "SEK" Then
                rsrs!dNS = ermdurchNS(CLng(cLinr), lagn)
            End If
            rsrs!AnzArtikel = ermAnzArt(" where Linr = " & cLinr & " and agn = " & lagn)
            rsrs.Update

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "Delete from PREISKALKAGNZ where dNS = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    ZeigeAGNTAB

    If MSFlexGrid2.Visible = True Then
        MSFlexGrid2.Col = SpaltennummerNSAGN
        MSFlexGrid2.Row = 2
        MSFlexGrid2.SetFocus
    End If
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKALK_AGN"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub ZeigeKALKARTIKEL(cLinr As String, cAgn As String, dNS As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim lagn            As Long
    Dim rsrs            As Recordset
    Dim dLEKPR          As Double
    Dim cMW             As String
    Dim cNewKassenPr    As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Artikel werden neu kalkuliert...", lblAnzeige
    
    Label2(2).Caption = "Ihre Auswahl: "
    Label2(2).Caption = Label2(2).Caption & " Lieferant: " & cLinr & " " & ermLiefBez(CLng(cLinr)) & " "
    
    If cAgn <> "" Then
        Label2(2).Caption = Label2(2).Caption & " Artikelgruppe: " & cAgn & " " & ermAGNbez(cAgn, gdBase)
    End If
    
    MSFlexGrid3.Clear
    
    loeschNEW "PREISKALKART", gdBase
    CreateTableT2 "PREISKALKART", gdBase
    
    sSQL = "Insert into PREISKALKART select artnr, LEKPR "
    sSQL = sSQL & " from Artlief where linr = " & cLinr & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PREISKALKART inner join Artikel on PREISKALKART.artnr = Artikel.artnr "
    sSQL = sSQL & " set PREISKALKART.BEZEICH = Artikel.BEZEICH"
    sSQL = sSQL & " , PREISKALKART.MWST = Artikel.MWST"
    sSQL = sSQL & " , PREISKALKART.Preisvor = Artikel.KVKPR1"
    sSQL = sSQL & " , PREISKALKART.AGN = Artikel.AGN"
    gdBase.Execute sSQL, dbFailOnError
    
    If gsSpanne = "SEK" Then
    
        sSQL = "Update PREISKALKART set LEKPR = 0 "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Update PREISKALKART inner join Artikel on PREISKALKART.artnr = Artikel.artnr "
        sSQL = sSQL & " set PREISKALKART.LEKPR = Artikel.EKPR"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    sSQL = "Delete from PREISKALKART where BEZEICH = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from PREISKALKART where BEZEICH is null "
    gdBase.Execute sSQL, dbFailOnError
    
    If cAgn <> "" Then
        sSQL = "Delete from PREISKALKART where AGN <> " & cAgn
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Set rsrs = gdBase.OpenRecordset("PREISKALKART")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!lekpr) Then
                dLEKPR = rsrs!lekpr
            Else
                dLEKPR = 0
            End If
            
            If Not IsNull(rsrs!MWST) Then
                cMW = rsrs!MWST
            Else
                cMW = "V"
            End If
            cNewKassenPr = Runden(CDbl(fnVKneuNS(dLEKPR, cMW, dNS)))

            rsrs.Edit
            rsrs!Preisnach = cNewKassenPr
            rsrs.Update

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    gdBase.Execute sSQL, dbFailOnError
    
    ZeigeARTIKELTAB

    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKALKArtikel"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub ZeigeLINRTAB()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    Dim recAnz      As Recordset
    
    Set recAnz = gdBase.OpenRecordset("PREISKALKLINRZ")
    If recAnz.EOF Then
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        
        anzeige "rot", "Keine Lieferanten gefunden!", lblAnzeige
        recAnz.Close
        Exit Sub
    End If
    recAnz.Close
    
    Screen.MousePointer = 11

    Tabcheck "KALKLINR"
    
    FormatGridOverTablay "KALKLINR"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex167
        ermittlespalten
        
        .Redraw = False
    
        Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeLINRTAB"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeAGNTAB()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    Dim recAnz      As Recordset
    
    Set recAnz = gdBase.OpenRecordset("PREISKALKAGNZ")
    If recAnz.EOF Then
        MSFlexGrid2.Visible = False
        MSFlexGrid2.Clear
        
        anzeige "rot", "Keine Artikelgruppen gefunden!", lblAnzeige
        recAnz.Close
        Exit Sub
    End If
    recAnz.Close
    
    Screen.MousePointer = 11

    Tabcheck "KALKAGN"
    
    FormatGridOverTablay "KALKAGN"

    With MSFlexGrid2
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex167a
        ermittlespalten2
        
        .Redraw = False
    
        Tabellenbreiteanpassen MSFlexGrid2, 1.25 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeAGNTAB"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeARTIKELTAB()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    Dim recAnz      As Recordset
    
    Set recAnz = gdBase.OpenRecordset("PREISKALKART")
    If recAnz.EOF Then
        MSFlexGrid3.Visible = False
        MSFlexGrid3.Clear
        
        anzeige "rot", "Keine Artikel gefunden!", lblAnzeige
        recAnz.Close
        Exit Sub
    End If
    recAnz.Close
    
    Screen.MousePointer = 11

    Tabcheck "KALKARTIKEL"
    
    FormatGridOverTablay "KALKARTIKEL"

    With MSFlexGrid3
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex167b
        ermittlespalten3
        
        .Redraw = False
    
        Tabellenbreiteanpassen MSFlexGrid3, 1.25 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeARTIKELTAB"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
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
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR

    If MSFlexGrid1.Row > 1 Then
    
    Else
        sortierenGrid MSFlexGrid1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_DblClick()
On Error GoTo LOKAL_ERROR

    If MSFlexGrid2.Row > 1 Then
    
    Else
        sortierenGrid MSFlexGrid2
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
            
        Case Is = vbKeyF3
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gcSuch = "LINR" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL70.Show 1
                Me.Refresh
                gcSuch = ""
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Rechneneu(lLinr As Long, dNS As Double) As Double
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sSQL As String
    Dim lartnr As Long
    Dim dEK As Double
    Dim cMWST As String
    Dim cKVKN As String
    Dim lBestand As Long
    Dim dKVKPR  As Double
    Dim dSumme As Double
    
    sSQL = "Select Artikel.ARTNR"
    sSQL = sSQL & " , Artikel.KVKPR1 "
    sSQL = sSQL & " , Artikel.EKPR "
    sSQL = sSQL & " , Artikel.MWST "
    sSQL = sSQL & " , Artikel.BESTAND "
    sSQL = sSQL & " from Artikel inner join Artlief on Artikel.artnr = Artlief.Artnr "
    sSQL = sSQL & " Where artlief.Linr = " & lLinr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
            
            End If
            
            If Not IsNull(rsrs!ekpr) Then
                dEK = rsrs!ekpr
            Else
            
            End If
        
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            Else
            
            End If
            
            cKVKN = fnVKneuNS(dEK, cMWST, dNS)
            
            If Not IsNull(rsrs!BESTAND) Then
                lBestand = rsrs!BESTAND
            Else
            
            End If
            
            If lBestand > 0 Then
                dSumme = dSumme + lBestand * CLng(cKVKN)
            End If
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close
    
    Rechneneu = dSumme
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rechneneu"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            
            Case Is = 4  'Linr
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Command4_Click 6
                End If
        End Select
    ElseIf KeyCode = vbKeyReturn Then
        Select Case Index
            
            Case Is = 4  'Linr
                Command4_Click 6
        End Select
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Preiskalkulation ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

