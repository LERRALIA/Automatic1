VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL130 
   Caption         =   "Rechnungsübersicht"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL130.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   14055
      Left            =   1080
      TabIndex        =   27
      Top             =   1560
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "bezahlte"
         Height          =   195
         Index           =   2
         Left            =   7800
         TabIndex        =   46
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "nicht bezahlte"
         Height          =   195
         Index           =   1
         Left            =   7800
         TabIndex        =   45
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "alle"
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   44
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   120
         MaxLength       =   6
         TabIndex        =   42
         Top             =   360
         Width           =   2175
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   0
         Left            =   2400
         TabIndex        =   41
         ToolTipText     =   "Auswahlhilfe"
         Top             =   340
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   38
         Tag             =   "2"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   6
         Left            =   9600
         TabIndex        =   36
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   5
         Left            =   9600
         TabIndex        =   35
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   31
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
         Caption         =   "Eingabe"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   30
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
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   9600
         TabIndex        =   28
         Top             =   6840
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
         Height          =   6015
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   29
         FixedCols       =   2
         ForeColorSel    =   8454143
         AllowBigSelection=   0   'False
         FocusRect       =   0
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
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   0
         Left            =   7200
         TabIndex        =   51
         ToolTipText     =   "Kalender"
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         Height          =   405
         Index           =   1
         Left            =   4800
         TabIndex        =   52
         ToolTipText     =   "Kalender"
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   10
         Left            =   4440
         TabIndex        =   53
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   9
         Left            =   4440
         TabIndex        =   54
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   8
         Left            =   6840
         TabIndex        =   55
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   7
         Left            =   6840
         TabIndex        =   56
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "kein Lieferant"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum bis:"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum von:"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         TabIndex        =   33
         Top             =   6840
         Width           =   7215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "schon bezahlt"
      Height          =   255
      Left            =   6600
      TabIndex        =   26
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3360
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   6600
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin sevCommand3.Command Command8 
      Height          =   165
      Left            =   4920
      TabIndex        =   22
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
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
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Zurück"
      ToolTipTitle    =   "Zurück"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command7 
      Height          =   165
      Left            =   4920
      TabIndex        =   21
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
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
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Vor"
      ToolTipTitle    =   "Vor"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   20
      ToolTipText     =   "Hilfe"
      Top             =   360
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
   Begin sevCommand3.Command Command1 
      Height          =   355
      Index           =   6
      Left            =   2400
      TabIndex        =   19
      ToolTipText     =   "Auswahlhilfe"
      Top             =   1200
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
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   1080
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
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   5
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   4
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   3
      Left            =   120
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   10
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   32
      Top             =   1680
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
      Caption         =   "Übersicht"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check2 
      Caption         =   "wird vom Konto abgebucht"
      Height          =   255
      Left            =   6600
      TabIndex        =   47
      Top             =   1920
      Width           =   2415
   End
   Begin sevCommand3.Command Command98 
      Height          =   360
      Left            =   10800
      TabIndex        =   48
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
      Picture         =   "frmWKL130.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command0 
      Height          =   405
      Index           =   20
      Left            =   5280
      TabIndex        =   49
      ToolTipText     =   "Kalender"
      Top             =   1080
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
      Height          =   405
      Index           =   21
      Left            =   8520
      TabIndex        =   50
      ToolTipText     =   "Kalender"
      Top             =   1080
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
   Begin sevCommand3.Command Command2 
      Height          =   165
      Index           =   0
      Left            =   8160
      TabIndex        =   57
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
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
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Vor"
      ToolTipTitle    =   "Vor"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   165
      Index           =   3
      Left            =   8160
      TabIndex        =   58
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   291
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
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
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
      ToolTip         =   "Zurück"
      ToolTipTitle    =   "Zurück"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
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
      Left            =   7440
      TabIndex        =   34
      Top             =   360
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
   Begin VB.Label Label2 
      Caption         =   "dazugehörige Lieferscheinnummer"
      Height          =   495
      Index           =   8
      Left            =   3360
      TabIndex        =   25
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Rechnungsnummer"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Bemerkung"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   23
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Rechnungsbetrag ohne MwSt."
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   18
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Rechnungsbetrag erm. MwSt."
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   17
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Rechnungsbetrag volle MwSt."
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "fällig am:"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Rechnungsdatum:"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   14
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "kein Lieferant"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      TabIndex        =   12
      Top             =   7920
      Width           =   9255
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnungsübersicht"
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
      TabIndex        =   11
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmWKL130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpaltennummerPOS As Byte
Dim SpaltennummerLINR As Byte
Dim SpaltennummerStatus As Byte
Dim SpaltennummerFaellig As Byte
Dim SpaltennummerRechdat As Byte
Dim cwhere As String
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 1        ' Kalender
            Text1(9).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(10).SetFocus
        Case Is = 0        ' Kalender
            Text1(10).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 20        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(2).SetFocus
        Case Is = 21        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 6     'F2 Lieferant
            Text1_KeyUp 0, vbKeyF2, 0
        Case Is = 0     'F2 Lieferant
            Text1_KeyUp 11, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 11
            gsHelpstring = "Rechnungsübersicht"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL130
        Case 1
            
            Rechnungspeichern
            leereFelder
        Case 2
            Frame1.Visible = True
            
            Dim iIndex As Integer
            If Option1(0).Value = True Then
                cwhere = ""
                iIndex = 0
            ElseIf Option1(1).Value = True Then
                cwhere = " where bezahlt = -1"
                iIndex = 1
            ElseIf Option1(2).Value = True Then
                cwhere = " where bezahlt = 0"
                iIndex = 2
            End If
            zeige_Grid iIndex, cwhere
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Rechnungspeichern()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim dateFaellig As Long
    Dim dateRechDat As Long
    Dim sRechnr     As String
    Dim sLiefsNr    As String
    Dim dBetragv    As Double
    Dim dBetrage    As Double
    Dim dBetrago    As Double
    Dim dAnteilv    As Double
    Dim dAnteile    As Double
    Dim dBetragNv   As Double
    Dim dBetragNe   As Double
    Dim sBemerk     As String
    Dim lLinr       As Long
    Dim sLiefBez    As String
    Dim ibezahlt    As Integer
    Dim sSTATUS     As String
    Dim sEinzug     As String
    
    lLinr = 0
    sLiefBez = ""
    If Text1(0).Text <> "" Then
        If IsNumeric(Text1(0).Text) Then
            lLinr = Val(Text1(0).Text)
            sLiefBez = ermLiefBez(lLinr)
        Else
            anzeige "rot", "Bitte geben Sie einen Lieferanten ein!", Label1(4)
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte geben Sie einen Lieferanten ein!", Label1(4)
        Exit Sub
    End If
    
    If Text1(1).Text <> "" Then
        If IsDate(Text1(1).Text) Then
            dateRechDat = CLng(DateValue(Text1(1).Text))
        Else
            anzeige "rot", "Bitte geben Sie ein Datum ein!", Label1(4)
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte geben Sie ein Datum ein!", Label1(4)
        Text1(1).SetFocus
        Exit Sub
    End If
    
    If Text1(2).Text <> "" Then
        If IsDate(Text1(2).Text) Then
            dateFaellig = CLng(DateValue(Text1(2).Text))
        Else
            anzeige "rot", "Bitte geben Sie ein Datum ein!", Label1(4)
            Exit Sub
        End If
    Else
        anzeige "rot", "Bitte geben Sie ein Datum ein!", Label1(4)
        Text1(2).SetFocus
        Exit Sub
    End If
    
    sBemerk = Text1(6).Text
    sLiefsNr = Text1(8).Text
    sRechnr = Text1(7).Text
    
    dBetragv = 0
    If Text1(3).Text <> "" Then
        If IsNumeric(Text1(3).Text) Then
            dBetragv = Text1(3).Text
        End If
    End If
    
    dBetrage = 0
    If Text1(4).Text <> "" Then
        If IsNumeric(Text1(4).Text) Then
            dBetrage = Text1(4).Text
        End If
    End If
    
    dBetrago = 0
    If Text1(5).Text <> "" Then
        If IsNumeric(Text1(5).Text) Then
            dBetrago = Text1(5).Text
        End If
    End If
    
    dBetragNv = dBetragv * 100 / (100 + gdMWStV)
    
    dBetragNe = dBetrage * 100 / (100 + gdMWStE)
    
    dAnteilv = dBetragv * gdMWStV / (100 + gdMWStV)
    
    dAnteile = dBetrage * gdMWStE / (100 + gdMWStE)
    
    If Check1.Value = vbChecked Then
        ibezahlt = 0
        sSTATUS = "bezahlt"
    Else
        ibezahlt = -1
        sSTATUS = "nicht bezahlt"
    End If
    
    If Check2.Value = vbChecked Then
        
        sEinzug = "J"
    Else
       
        ssEinzug = ""
    End If
    
    
    cSQL = "Insert into RECHUE"
    cSQL = cSQL & " ( "
    cSQL = cSQL & " RechNr  "
    cSQL = cSQL & ", LiefSNr  "
    cSQL = cSQL & ", RechBTv  "
    cSQL = cSQL & ", RechBTe  "
    cSQL = cSQL & ", RechBTo  "
    cSQL = cSQL & ", Anteilv  "
    cSQL = cSQL & ", Anteile  "
    cSQL = cSQL & ", RechBTnV  "
    cSQL = cSQL & ", RechBTnE  "
    cSQL = cSQL & ", Bemerk  "
    cSQL = cSQL & ", Rechdat  "
    cSQL = cSQL & ", faellig  "
    cSQL = cSQL & ", LiNr  "
    cSQL = cSQL & ", LiefBez  "
    cSQL = cSQL & ", BEZAHLT  "
    cSQL = cSQL & ", STATUS  "
    cSQL = cSQL & ", EINZUG  "
    cSQL = cSQL & " ) values ( "
    cSQL = cSQL & " '" & sRechnr & "'  "
    cSQL = cSQL & ", '" & sLiefsNr & "'  "
    
    cSQL = cSQL & ", '" & dBetragv & "'  "
    cSQL = cSQL & ", '" & dBetrage & "'  "
    cSQL = cSQL & ", '" & dBetrago & "'  "
    
    cSQL = cSQL & ", '" & dAnteilv & "'  "
    cSQL = cSQL & ", '" & dAnteile & "'  "
    cSQL = cSQL & ", '" & dBetragNv & "'  "
    cSQL = cSQL & ", '" & dBetragNe & "'  "
    cSQL = cSQL & ", '" & sBemerk & "'  "
    cSQL = cSQL & ", " & dateRechDat & "  "
    cSQL = cSQL & ", " & dateFaellig & "  "
    cSQL = cSQL & "," & lLinr & " "
    cSQL = cSQL & ",'" & sLiefBez & "' "
    cSQL = cSQL & "," & ibezahlt & " "
    cSQL = cSQL & ",'" & sSTATUS & "' "
    cSQL = cSQL & ",'" & sEinzug & "' "
    cSQL = cSQL & " )  "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rechnungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leereFelder()
On Error GoTo LOKAL_ERROR

    Dim i  As Integer
    
    For i = 0 To 8
        Text1(i).Text = ""
    Next i
    
    Check1.Value = vbUnchecked
    
    Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    Text1(0).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leereFelder"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        End If
        lDat = lDat + 1
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        End If
        lDat = lDat - 1
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lDat As Long
    Dim sSQL As String
    Dim iRet As Integer
    Dim cSuch As String
    Dim ctmp As String
    Dim dSummeNetto As Double
    Dim rsrs As Recordset
    
    Select Case Index
        Case 0
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat + 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 1
            'suche
            Rechnungssuche
        Case 10
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat + 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 8
            If IsDate(Text1(10).Text) = False Then
                Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(10).Text) = True Then
                    lDat = CLng(DateValue(Text1(10).Text))
                End If
                lDat = lDat + 1
                Text1(10).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 1
            'suche
        Case 2
            Frame1.Visible = False
        Case 3
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat - 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 9
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat - 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 7
            If IsDate(Text1(10).Text) = False Then
                Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(10).Text) = True Then
                    lDat = CLng(DateValue(Text1(10).Text))
                End If
                lDat = lDat - 1
                Text1(10).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 4
            Unload frmWKL130
        Case 5
            'Löschen
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                anzeige "rot", "Bitte einen Satz in der Tabelle markieren!", Label1(0)
                Exit Sub
            End If
            
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, CLng(SpaltennummerPOS))
            cSuch = Trim$(cSuch)
                
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                anzeige "rot", "Bitte einen Satz in der Tabelle markieren!", Label1(0)
                Exit Sub
            End If
            
            ctmp = "Möchten Sie diese Rechnung wirklich löschen?" & vbCrLf & vbCrLf
            ctmp = ctmp & "Diese Informationen stehen dann für eine eventuelle Jahresauswertung nicht mehr zur Verfügung."
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
                LoescheRechnung CLng(cSuch)
                
                Dim iIndex As Integer
                If Option1(0).Value = True Then
                    cwhere = ""
                    iIndex = 0
                ElseIf Option1(1).Value = True Then
                    cwhere = " where bezahlt = -1"
                    iIndex = 1
                ElseIf Option1(2).Value = True Then
                    cwhere = " where bezahlt = 0"
                    iIndex = 2
                End If
                zeige_Grid iIndex, cwhere
            End If
        
        Case 6
            'drucken
            
            If MSFlexGrid1.Visible = False Then
                Screen.MousePointer = 0
                
                Exit Sub
            End If
            
            Dim dNettovoll As Double
            Dim dNettoerm As Double
            
            loeschNEW "PRIRECH2", gdBase
            CreateTable "PRIRECH2", gdBase
            
            sSQL = "Insert Into PRIRECH2 select * from Prirech order by faellig "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update PRIRECH2 set DRUCKMARK = true where faellig <= clng(datevalue(now)) and Status = 'nicht bezahlt' "
            gdBase.Execute sSQL, dbFailOnError
            
            dSummeNetto = 0
            Set rsrs = gdBase.OpenRecordset("select * from prirech2 order by pos")
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                
                    dNettovoll = 0
                    If Not IsNull(rsrs!Rechbtnv) Then
                        dNettovoll = CDbl(rsrs!Rechbtnv)
                    End If
                    
                    dNettoerm = 0
                    If Not IsNull(rsrs!Rechbtne) Then
                        dNettoerm = CDbl(rsrs!Rechbtne)
                    End If
                    
                    dSummeNetto = dSummeNetto + dNettoerm + dNettovoll
                    
                    rsrs.Edit
                    rsrs!summenetto = dSummeNetto
                    rsrs.Update
        
                rsrs.MoveNext
                Loop
            End If
            rsrs.Close: Set rsrs = Nothing
            
            anzeige "normal", "Druckvorschau wird erstellt...", Label1(0)
            
            reportbildschirm "", "aWKL130a"
            anzeige "normal", "", Label1(0)
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Rechnungssuche()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bAnd As Boolean
    Dim sLinr As String
    Dim iIndex As Integer
    Dim lVon As Long
    Dim lBis As Long
    Dim cVon As String
    Dim cBis As String
    
    sLinr = ""
    bAnd = False
    
    If Option1(0).Value = True Then
        cwhere = ""
        iIndex = 0
    ElseIf Option1(1).Value = True Then
        cwhere = " where bezahlt = -1"
        bAnd = True
        iIndex = 1
    ElseIf Option1(2).Value = True Then
        cwhere = " where bezahlt = 0"
        bAnd = True
        iIndex = 2
    End If
    
    If Text1(11).Text <> "" Then
        If IsNumeric(Text1(11).Text) Then
            sLinr = Trim(Text1(11).Text)
            If bAnd = True Then
                cwhere = cwhere & " and "
            Else
                cwhere = cwhere & " where "
            End If
            
            cwhere = cwhere & " linr = " & sLinr
            bAnd = True
        End If
    End If
    
    'von
    If Text1(9).Text <> "" Then
        If IsDate(Text1(9).Text) Then
            cVon = Trim(Text1(9).Text)
            lVon = DateValue(cVon)
            
    
            

            
            If bAnd = True Then
                cwhere = cwhere & " and "
            Else
                cwhere = cwhere & " where "
            End If
            
            cwhere = cwhere & " rechdat >= " & lVon
            bAnd = True
        End If
    End If
    
    'bis
    If Text1(10).Text <> "" Then
        If IsDate(Text1(10).Text) Then
            cBis = Trim(Text1(10).Text)
            lBis = DateValue(cBis)
            
            If bAnd = True Then
                cwhere = cwhere & " and "
            Else
                cwhere = cwhere & " where "
            End If
            
            cwhere = cwhere & " rechdat <= " & lBis
            bAnd = True
        End If
    End If
    
    zeige_Grid iIndex, cwhere
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rechnungssuche"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub LoescheRechnung(lPos As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from RECHUE where Autopos = " & lPos
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheRechnung"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "AUTOPOS"
                SpaltennummerPOS = i
            Case Is = "STATUS"
                SpaltennummerStatus = i
            Case Is = "LINR"
                SpaltennummerLINR = i
            Case Is = "RECHDAT"
                SpaltennummerRechdat = i
            Case Is = "FAELLIG"
                SpaltennummerFaellig = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "LINR"
    gsZSpalte1 = "AUTOPOS"
    gstab = "RECHNUNG"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    

    WKL130Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    
    anzeige "normal", "", Label1(4)
    
    Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    Text1(9).Text = Format("01.01." & Year(Now), "DD.MM.YYYY")
    Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    Label2(3).Caption = "Rechnungsbetrag " & gdMWStV & " % MwSt."
    Label2(3).Refresh
    Label2(4).Caption = "Rechnungsbetrag " & gdMWStE & " % MwSt."
    Label2(4).Refresh
    
    Option1(1).Value = True
    cwhere = " where bezahlt = -1"
    zeige_Grid 1, cwhere
    
    Label3(6).Caption = "Heute: " & Format(DateValue(Now), "DD.MM.YY")
    Label3(6).Refresh
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
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
Private Sub FuellenMSFlex130(cwhere As String)
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
    
    loeschNEW "PRIRECH", gdBase
    CreateTable "PRIRECH", gdBase
    
    cSQL = "Insert into PRIRECH Select * from RECHUE "
    
    cSQL = cSQL & cwhere

    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Select * from PRIRECH "
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
                            Case Is = "Betrag v MwSt", "Betrag e MwSt", "Betrag o MwSt", "Netto v Mwst", "Netto e Mwst"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                            Case Is = "MwSt Anteil v", "MwSt Anteil e"
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
            
        
        rsrs.Close: Set rsrs = Nothing
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex130"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
        
    Fehlermeldung1
   
End Sub
Private Sub zeige_Grid(ibezahlt As Integer, cwhere As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    If Not NewTableSuchenDBKombi("RECHUE", gdBase) Then
        anzeige "rot2", "Keine Daten gefunden!", Label1(0)
        
        Exit Sub
    End If
    
    Set recAnz = gdBase.OpenRecordset("RECHUE")
    
    If recAnz.EOF Then
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        anzeige "rot2", "Keine Daten gefunden!", Label1(0)
        Exit Sub
    Else
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    
    Screen.MousePointer = 11

    Tabcheck "RECHNUNG"
    
    FormatGridOverTablay "RECHNUNG"

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
    
        FuellenMSFlex130 cwhere
        
        ermittlespalten
        
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
        
        If ibezahlt = 0 Then
            lblUeberschrift.Caption = "alle Rechnungen"
            lblUeberschrift.Refresh
        ElseIf ibezahlt = 1 Then
            lblUeberschrift.Caption = "noch nicht bezahlte"
            lblUeberschrift.Refresh
        ElseIf ibezahlt = 2 Then
            lblUeberschrift.Caption = "bezahlte"
            lblUeberschrift.Refresh
        End If
        
'        .SetFocus
    
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Grid"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub WKL130Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 0
    Frame1.Width = 11775
    Frame1.Height = 7455
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL130Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
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

Private Sub MSFlexGrid1_Click()
On Error GoTo LOKAL_ERROR
    
    Text1(11).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, CLng(SpaltennummerLINR))
    
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
    
    Dim sSQL As String
    Dim cStat As String
    Dim cSuch As String
    Dim ibezahlt As String
    Dim iIndex As Integer
    
    If MSFlexGrid1.Row > 1 Then
    
        cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, CLng(SpaltennummerPOS))
        cSuch = Trim$(cSuch)
            
        If IsNumeric(cSuch) Then

        Else
            Exit Sub
        End If

        'wenn spalte Status dann Status wechseln
        If MSFlexGrid1.Col = SpaltennummerStatus Then
            cStat = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, CLng(SpaltennummerStatus)))
            If cStat = "nicht bezahlt" Then
                ibezahlt = 0
                cStat = "bezahlt"
            Else
                ibezahlt = -1
                cStat = "nicht bezahlt"
            End If
            
            sSQL = "Update RECHUE set Status = '" & cStat & "' "
            sSQL = sSQL & " , bezahlt = " & ibezahlt & " where Autopos = " & CLng(cSuch)
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
        If Option1(0).Value = True Then
            cwhere = ""
            iIndex = 0
        ElseIf Option1(1).Value = True Then
            cwhere = " where bezahlt = -1"
            iIndex = 1
        ElseIf Option1(2).Value = True Then
            cwhere = " where bezahlt = 0"
            iIndex = 2
        End If
        zeige_Grid iIndex, cwhere
    Else
        If Option1(0).Value = True Then
            cwhere = ""
            iIndex = 0
        ElseIf Option1(1).Value = True Then
            cwhere = " where bezahlt = -1"
            iIndex = 1
        ElseIf Option1(2).Value = True Then
            cwhere = " where bezahlt = 0"
            iIndex = 2
        End If
        
    
        If MSFlexGrid1.Col = SpaltennummerFaellig Then
        
            If byteSortReihen = 1 Then
                byteSortReihen = 2
                zeige_Grid iIndex, cwhere & " order by faellig desc"
            Else
                byteSortReihen = 1
                zeige_Grid iIndex, cwhere & " order by faellig asc"
            End If
            
        
        ElseIf MSFlexGrid1.Col = SpaltennummerRechdat Then
            If byteSortReihen = 1 Then
                byteSortReihen = 2
                zeige_Grid iIndex, cwhere & " order by Rechdat desc"
            Else
                byteSortReihen = 1
                zeige_Grid iIndex, cwhere & " order by Rechdat asc"
            End If
        
        Else
            sortierenGrid MSFlexGrid1
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Command2_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Index = 0 Then
        LiefKuerzelAufloesung Label2(0), Text1(0)
    ElseIf Index = 11 Then
        LiefKuerzelAufloesung Label2(9), Text1(11)
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 0, 6, 7, 8, 11 'Bemerkung
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%"
        Case 1, 2 ' Datum
            cValid = "1234567890." & Chr$(8)
        Case 3, 4, 5 'Rechnungsbeträge
            cValid = "1234567890," & Chr$(8)
    End Select

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Text1(Index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
    
    End If
    
    If KeyCode = vbKeyEscape Then
        Command5_Click 0
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 0
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label2(0).Caption = gF2Prompt.cWert
                End If
            Case Is = 11
                gF2Prompt.cFeld = "LINR"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label2(9).Caption = gF2Prompt.cWert
                End If
        End Select
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Rechnungsübersicht ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
