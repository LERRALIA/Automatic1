VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL192 
   Caption         =   "MHD"
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   8160
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   12615
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   9600
         TabIndex        =   27
         Top             =   1080
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "Heute"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Morgen"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "nächsten 7 Tage"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   8
         Left            =   9600
         TabIndex        =   9
         Top             =   6120
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
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   5
         Top             =   3000
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11668
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
         SelectionMode   =   1
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
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   11160
         TabIndex        =   22
         ToolTipText     =   "Kalender"
         Top             =   120
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
         Index           =   1
         Left            =   11160
         TabIndex        =   23
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   24
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
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
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   25
         Top             =   120
         Width           =   495
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
         TabIndex        =   15
         Top             =   5760
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
         TabIndex        =   14
         Top             =   5520
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
         TabIndex        =   13
         Top             =   5160
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
         TabIndex        =   12
         Top             =   4920
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
         TabIndex        =   11
         Top             =   4560
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
         TabIndex        =   10
         Top             =   4320
         Width           =   2055
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
         TabIndex        =   7
         Top             =   3720
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
         TabIndex        =   6
         Top             =   3960
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
      Left            =   3240
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
      TabIndex        =   19
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
      Picture         =   "frmWKL193.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   31
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.Label Label2 
      Caption         =   "Artikelbestand 0 = aus Liste entfernen"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   32
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "F2: Artikelmaske"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "MHD"
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
      Width           =   2175
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
Attribute VB_Name = "frmWKL192"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerArtnr          As Byte
Dim SpaltennummerAWM            As Byte
Dim SpaltennummerMDHDAT         As Byte
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "Farbnr"
    gstab = "MDH"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loescheausFilz(lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr          As String
    Dim lDat            As Long
    Dim cSQL            As String
   
    cArtNr = MSFlexGrid1.TextMatrix(lrow, SpaltennummerArtnr)
    lDat = DateValue(MSFlexGrid1.TextMatrix(lrow, SpaltennummerMDHDAT))
    If cArtNr <> "" Then
        If IsNumeric(cArtNr) Then
            cSQL = "Delete from art192 "
            cSQL = cSQL & " where ARTNR = " & cArtNr
            cSQL = cSQL & " and MDHDAT = " & lDat
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Delete from artmdh "
            cSQL = cSQL & " where ARTNR = " & cArtNr
            cSQL = cSQL & " and MDHDAT = " & lDat
            gdBase.Execute cSQL, dbFailOnError
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheausFilz"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
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
            Unload frmWKL192
        Case 1
            'suchen
            ermMDH_Artikel Text1(0).Text, Text1(1).Text
            ZeigeArtikel192
        Case 2
            If MSFlexGrid1.RowSel >= 1 Then
                FlexGrid_Update MSFlexGrid1
            End If
            ZeigeArtikel192
        Case 3
        
            'war mal
            DoppelteBer
            
            'Artikelbestand = 0 = entfernen
            ArtikelBestandNullEntfernen
            
            'suchen
            ermMDH_Artikel Text1(0).Text, Text1(1).Text
            ZeigeArtikel192
        Case 8
            Drucke_MDH

        Case 12
            gsHelpstring = "MDH"
            frmWKL110.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1(0).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            Text1(1).SetFocus
        Case Is = 1
            Text1(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ZeigeArtikel192()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    
    MSFlexGrid1.Clear
    
    If Not NewTableSuchenDBKombi("art192", gdBase) Then
        anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
        Exit Sub
    Else
        If Datendrin("art192", gdBase) = False Then
            anzeige "rot", "Es sind keine Artikel ermittelt worden.", Label1(4)
            Exit Sub
        End If
    End If
    
    anzeige "normal", "Artikel werden angezeigt, bitte warten...", Label1(4)
    
    Screen.MousePointer = 11
    
    Tabcheck "MDH"
    FormatGridOverTablay "MDH"

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
    
    GridFuellen "Select * from art192 order by Bezeich, mdhdat "
    
    FaerbenGrid MSFlexGrid1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikel192"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Drucke_MDH()
    On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim i As Integer

    Screen.MousePointer = 11

    loeschNEW "ART192PRINT", gdBase
    CreateTableT2 "ART192PRINT", gdBase

    cSQL = "Insert into ART192PRINT select * from art192  "
    gdBase.Execute cSQL, dbFailOnError
    
'    cSQL = "Update ART192PRINT inner join Artikel on ART192PRINT.ARTNR = Artikel.ARTNR "
'    cSQL = cSQL & " set ART192PRINT.ean = Artikel.ean "
'    gdBase.Execute cSQL, dbFailOnError

    anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)

    reportbildschirm "", "aZEN192a"

    anzeige "normal", "", Label1(4)

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_MDH"
    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "FARBNR"
                SpaltennummerAWM = i
            Case Is = "MDHDAT"
                SpaltennummerMDHDAT = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

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
        
        
    Else
        Frame2.Visible = False
        anzeige "normal", "", Label1(3)
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label1(4)
        
        Label2(0).Visible = False
        
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
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim cVon As String
    Dim cBis As String
    
    Screen.MousePointer = 11
    
    PositionierenWKL192
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    
    
    
    
    
    Me.Refresh
    
    cVon = ""
    cBis = ""
    
    ermMDH_Artikel cVon, cBis
    ZeigeArtikel192
    
    Frame2.Visible = True
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index

        Case Is = 5     'nächsten 7 Tage
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now) + 7, "DD.MM.YY")
        Case Is = 6     'morgen
            Text1(0).Text = Format(DateValue(Now) + 1, "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now) + 1, "DD.MM.YY")
        Case Is = 7     'heute
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

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
Private Sub PositionierenWKL192()
On Error GoTo LOKAL_ERROR

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
    Fehler.gsFunktion = "PositionierenWKL192"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR


    loeschNEW "ART192", gdBase
    loeschNEW "ART192PRINT", gdBase
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
Private Sub ermMDH_Artikel(cVon As String, cBis As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim siAnzeige As Single
    Dim lVon As Long
    Dim lBis As Long
    
    If cVon <> "" And cBis <> "" Then
        lVon = DateValue(cVon)
        lBis = DateValue(cBis)
        
        cVon = Trim$(Str$(lVon))
        cBis = Trim$(Str$(lBis))
    End If
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "ART192", gdBase
    CreateTableT2 "ART192", gdBase
    
    anzeige "normal", "die Artikel werden ermittelt...", Label1(4)
    
    sSQL = " Insert into art192 select  a.ARTNR"
    sSQL = sSQL & ", a.Bezeich "
    sSQL = sSQL & ", a.KVKPR1 "
    sSQL = sSQL & ", a.BESTAND "
    sSQL = sSQL & ", m.MDHDAT  "
    sSQL = sSQL & ", a.ean  "
    sSQL = sSQL & ", val(a.AWM) as FARBNR "
    sSQL = sSQL & " from Artikel a , artmdh m "
    sSQL = sSQL & " where a.artnr = m.artnr "
    
    If cVon <> "" And cBis <> "" Then
        sSQL = sSQL & " and m.MDHDAT between  " & cVon & " And " & cBis
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35

    anzeige "normal", "nicht relevante Artikel werden gelöscht...", Label1(4)
    
    txtStatus.Text = 56
    
    BringFarbeInsSpiel "Art192", gdBase
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
 
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermMDH_Artikel"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub DoppelteBer()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim sArtnr As String
    Dim lMDHDAT As Long
    Dim sAzeit As String


    Screen.MousePointer = 11
    
    
    
    
    
    loeschNEW "ARTMHDDOPP", gdBase
    
    sSQL = "Select * into ARTMHDDOPP from artmdh"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from ARTMHDDOPP "
    gdBase.Execute sSQL, dbFailOnError



    sSQL = "Select * from artmdh order by ARTNR, MDHDAT "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!MDHDAT) Then
                lMDHDAT = rsrs!MDHDAT
            End If
            
            
            
            If Not IsNull(rsrs!AZEIT) Then
                sAzeit = rsrs!AZEIT
            End If
            
            
            
            
            
            sSQL = "Select * from ARTMHDDOPP where artnr = " & sArtnr & " and MDHDAT = " & Trim$(Str$(lMDHDAT)) & ""
    
            Set rsRs2 = gdBase.OpenRecordset(sSQL)
            If rsRs2.EOF Then
            
                'dann insert
                
                sSQL = "Insert into ARTMHDDOPP Select * from ArtMDH where artnr = " & sArtnr & " and MDHDAT = " & Trim$(Str$(lMDHDAT)) & ""
                sSQL = sSQL & " and azeit = '" & sAzeit & "'"
                gdBase.Execute sSQL, dbFailOnError
            
                
                
            End If
        
            rsRs2.Close: Set rsRs2 = Nothing
            
        
            
            
            
            
        rsrs.MoveNext
        Loop
    End If
        
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "artmdh", gdBase
    
    sSQL = "Select * into artmdh from ARTMHDDOPP "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "ARTMHDDOPP", gdBase

    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DoppelteBer"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ArtikelBestandNullEntfernen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String


    Screen.MousePointer = 11
    
    If SpalteInTabellegefundenNEW("artmdh", "Bestand", gdBase) = False Then
        sSQL = " Alter table artmdh add Bestand long  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "update artmdh inner join artikel on artmdh.artnr = artikel.artnr   "
    sSQL = sSQL & " set artmdh.bestand = artikel.bestand "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from artmdh where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
            
    If SpalteInTabellegefundenNEW("artmdh", "Bestand", gdBase) = True Then
        sSQL = " Alter table artmdh drop Bestand   "
        gdBase.Execute sSQL, dbFailOnError
    End If
   
    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelBestandNullEntfernen"
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."

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
            Me.Refresh
            Screen.MousePointer = 11

            MSFlexGrid1.TopRow = lrow
            MSFlexGrid1.Col = SpaltennummerArtnr
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
            Screen.MousePointer = 0
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
    Fehler.gsFehlertext = "Im Programmteil MDH Bearbeitung ist ein Fehler aufgetreten."
    
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

'Private Sub Text1_GotFocus(Index As Integer)
'On Error GoTo LOKAL_ERROR
'    Text1(Index).BackColor = glSelBack1
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text1_GotFocus"
'    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    On Error GoTo LOKAL_ERROR
'
'    Dim cZeichen As String
'    Dim cValid As String
'    cZeichen = Chr$(KeyAscii)
'    cZeichen = UCase$(cZeichen)
'    KeyAscii = Asc(cZeichen)
'
'    Select Case Index
'        Case 0
'            cValid = "1234567890," & Chr$(8)
'            If InStr(cValid, cZeichen) = 0 Then
'                KeyAscii = 0
'            End If
'        Case 1, 2
'            cValid = "1234567890" & Chr$(8)
'            If InStr(cValid, cZeichen) = 0 Then
'                KeyAscii = 0
'            End If
''''            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
''''            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
''''            cValid = cValid & "+äÄÜüÖöß#"
''''
''''            If InStr(cValid, cZeichen) = 0 Then
''''                KeyAscii = 0
''''            End If
''''            'alle Zeichen erlaubt
'    End Select
'
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text1_KeyPress"
'    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub

'Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo LOKAL_ERROR
'
'    Dim sAuswahlfeld As String
'    Dim ctmp As String
'    Dim lcount As Long
'
'    If KeyCode = vbKeyReturn Then
'        Command5_Click 6
'    End If
'
'    If KeyCode = vbKeyEscape Then
'        Command5_Click 0
'    End If
'
'
'    If KeyCode = vbKeyF2 Then
'        gF2Prompt.cFeld = ""
'        gF2Prompt.cWert = ""
'        gF2Prompt.cWert2 = ""
'        gF2Prompt.cWahl = ""
'        gF2Prompt.bMultiple = True
'
'        Select Case Index
'            Case 2
'                gF2Prompt.cFeld = "LINR"
'                If gF2Prompt.cFeld <> "" Then
'                    frmWK00a.Show 1
'                    If gF2Prompt.cWahl <> "" Then
'                        Text1(Index).Text = gF2Prompt.cWahl
'                        Text1(Index).Text = Trim(Text1(Index).Text)
'                    End If
'                End If
'
'                List3.Visible = False
'                List3.Clear
'                For lcount = 0 To 100
'                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
'                        List3.Visible = True
'                        Text1(Index).Text = ""
'
'                        If gF2Prompt.cArray(lcount) <> "" Then
'                            List3.AddItem gF2Prompt.cArray(lcount)
'                        End If
'
'                    Else
'
'                        If gF2Prompt.cArray(lcount) <> "" Then
'
'                            List3.AddItem gF2Prompt.cArray(lcount)
'                            Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 6)
'                            Text1(Index).Text = Trim(Text1(Index).Text)
'                        End If
'
'                    End If
'                Next lcount
'
'                If List3.Visible = True Then
'                    Label1(16).Visible = True
'                    Label1(16).Caption = List3.ListCount & " Lieferanten"
'                    Label1(16).Refresh
'                    Command5(1).Visible = True
'                Else
'                    Label1(16).Visible = False
'                    Command5(1).Visible = False
'                End If
'
'        End Select
'        Text1(Index).SetFocus
'    End If
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text1_KeyUp"
'    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub Text1_LostFocus(Index As Integer)
'On Error GoTo LOKAL_ERROR
'    Text1(Index).BackColor = vbWhite
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text1_LostFocus"
'    Fehler.gsFehlertext = "Im Programmteil Pennerbearbeitung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
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


