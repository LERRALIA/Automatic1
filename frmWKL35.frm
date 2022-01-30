VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL35 
   BackColor       =   &H00C0C000&
   Caption         =   "Kalkulation der Preise"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL35.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Tag             =   "KALKULA"
   WhatsThisHelp   =   -1  'True
   Begin sevCommand3.Command cmdHelp 
      Height          =   555
      Left            =   10440
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
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
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C000&
      Caption         =   "schon kalkulierte Artikel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   42
      Top             =   120
      Width           =   975
   End
   Begin sevCommand3.Command Command6 
      Height          =   375
      Left            =   8880
      TabIndex        =   37
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
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
      Caption         =   "Vorschau drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   360
      TabIndex        =   26
      Top             =   7200
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C000&
         Caption         =   "Automatische Kalkulation für diese Artikel aktivieren"
         Height          =   615
         Left            =   3480
         TabIndex        =   44
         Top             =   600
         Value           =   1  'Aktiviert
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin sevCommand3.Command Command3 
         Height          =   360
         Left            =   5880
         TabIndex        =   34
         Top             =   720
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Artikel kalkulieren"
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   120
         Width           =   2295
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Left            =   5880
         TabIndex        =   32
         Top             =   360
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
         Caption         =   "Runden"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   360
         Left            =   5880
         TabIndex        =   31
         Top             =   0
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
         Caption         =   "Berechnung"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Preis ersetzen in Euro"
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Aufschlag in Euro"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   39
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Aufschlag in %"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nettospanne in %"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   1335
      End
   End
   Begin sevCommand3.Command cmdGo 
      Height          =   310
      Left            =   10800
      TabIndex        =   8
      Top             =   960
      Width           =   450
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
      MaskColor       =   16777215
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Go"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   8760
      MaxLength       =   13
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   8160
      MaxLength       =   3
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   7560
      MaxLength       =   3
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   960
      MaxLength       =   13
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   7920
      Width           =   2415
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
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   1095
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      _Version        =   393216
      FocusRect       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin sevCommand3.Command Command1 
      Height          =   315
      Left            =   480
      TabIndex        =   45
      Top             =   960
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
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
      Picture         =   "frmWKL35.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Berechnungsgrundlage:"
      Height          =   255
      Index           =   9
      Left            =   6480
      TabIndex        =   36
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EK"
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   35
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   3
      X1              =   9240
      X2              =   9240
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   2
      X1              =   5040
      X2              =   9240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   1
      X1              =   480
      X2              =   11280
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RK"
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   25
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ab"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "auf"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Rundungskriterium:"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "aufrunden auf:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "abrunden auf:"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   20
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label0 
      Caption         =   "Label2"
      Height          =   135
      Index           =   1
      Left            =   4680
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label0 
      Caption         =   "Label2"
      Height          =   135
      Index           =   0
      Left            =   3600
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   16
      Top             =   6720
      Width           =   10815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   11280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      Caption         =   "Artikel-Bezeichnung"
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
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      Caption         =   "EAN-Code / Artnr"
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
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      Caption         =   "Lieferanten-Nr"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      Caption         =   "Lief.Best.Nr"
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
      Index           =   6
      Left            =   8760
      TabIndex        =   10
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Index           =   8
      Left            =   7560
      TabIndex        =   9
      Top             =   720
      Width           =   615
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
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmWKL35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim berstesZeichen As Boolean
Private Sub Urzustandherstellen()
    On Error GoTo LOKAL_ERROR

    Check1.Value = vbUnchecked
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Urzustandherstellen"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "artueb", gdBase
    loeschNEW "KALKHEAD", gdBase
    
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
Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check1.Value = vbChecked Then
        sSQL = "Update ARTUEB SET ETIMERK = 'J'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update ARTUEB SET ETIMERK = 'N'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    FormatMShFlex1WKLad
    FuellenMShFlex1WKLad
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
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
Private Sub cmdGo_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    
    Command2.Enabled = False
    Command4.Enabled = False
    Text2.Text = ""
    Text3.Text = ""
    
    iRet = fnPruefeEingabeWKLad()
    If iRet <> 0 Then
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Bitte mindestens ein Suchkriterium angeben!"
        Text1(2).SetFocus
        Exit Sub
    End If

    lblAnzeige.ForeColor = glS1
    lblAnzeige.Caption = "Daten werden ermittelt, bitte warten..."
    lblAnzeige.Refresh
    
    Screen.MousePointer = 11
    
    'Grid formatieren
    Tabcheck "Artueb"
    FormatGridOverTablay "Artueb"
    
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
            aBreite(j) = Len(.TextMatrix(0, j)) * 80
        Next j
    End With
    
    SucheArtikelWKLad
    
    If ErmittlungArtikelDuplis("ARTUEB", gdBase) <> "0" Then
        Screen.MousePointer = 0
        
        
        'doppelte anzeigen
        sSQL = " Update ARTUEB inner join  ALIT on"
        sSQL = sSQL & " ARTUEB.ARTNR = ALIT.ARTNR  Set AWM = '96' "
        gdBase.Execute sSQL, dbFailOnError
        
        FuellenMShFlex1WKLad
        
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "doppelte Artikel!! Bitte schränken Sie ihre Suche weiter ein!"
        lblAnzeige.Refresh

        Exit Sub
    Else
    
        FuellenMShFlex1WKLad
    
    End If
    
    Urzustandherstellen
    
    If MSHFLEX1.Visible = True Then
        Frame1.Visible = True
        Command6.Visible = True
    Else
        Frame1.Visible = False
        Command6.Visible = False
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeWKLad()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnPruefeEingabeWKLad = 1
    
    For lcount = 2 To 7
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeWKLad = 0
            Exit Function
        End If
    Next lcount
               
    If Check3.Value = vbChecked Then fnPruefeEingabeWKLad = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKLad"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub SucheArtikelWKLad()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cFeld       As String
    Dim cwhere      As String
    Dim lcol        As Long
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    Dim cVon        As String
    Dim cBis        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim cJoin       As String
    Dim sSQL        As String
    
    cSQL = "Select A.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", B.LEKPR as LEK"
    cSQL = cSQL & ", B.LINR"
    cSQL = cSQL & ", B.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.ETIMERK"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR as SEK"
    cSQL = cSQL & ", A.KVKPR1 as KVKA"
    cSQL = cSQL & ", B.Spanne as NSA"
    cSQL = cSQL & ", A.MWST "
    
    cwhere = ""
    
    cSQL = cSQL & " from ARTIKEL A, Artlief B  "
    
    If cwhere = "" Then
        cwhere = "where "
    Else
        cwhere = cwhere & "and "
    End If
    cwhere = cwhere & "A.ARTNR = B.ARTNR "
    
    
    cFeld = Text1(2).Text       'Bezeich
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.BEZEICH like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(3).Text   'EAN oder ARTNR
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cEAN = cFeld
        If Len(cFeld) <= 6 Then
            cArtNr = cFeld
        Else
            cArtNr = ""
        End If
        If Left(cFeld, 1) = "2" Or Left(cFeld, 1) = "0" And Len(cFeld) = 8 Then
            cEigNr = Mid(cFeld, 2, 6)
        Else
            cEigNr = ""
        End If
        
        cwhere = cwhere & "("
        If cEAN <> "" Then
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "A.EAN like '" & cEAN & "' "
            Else
                cwhere = cwhere & "A.EAN = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN2 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN2 = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN3 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN3 = '" & cEAN & "' "
            End If
        End If
        If cArtNr <> "" Then
            If InStr(cArtNr, "*") > 0 Then
                cwhere = cwhere & " or A.ARTNR like '" & cArtNr & "' "
            Else
                cwhere = cwhere & " or A.ARTNR = " & cArtNr & " "
            End If
        End If
        If cEigNr <> "" Then
            cwhere = cwhere & " or A.ARTNR = " & cEigNr & " "
        End If
        cwhere = cwhere & ") "
        
    End If
    
    cFeld = Text1(4).Text       'Linr
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "B.LINR = " & cFeld & " "
    End If
    
    cFeld = Text1(7).Text       'LPZ
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.LPZ = " & cFeld & " "
    End If
    
    cFeld = Text1(5).Text       'AGN
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.AGN = " & cFeld & " "
    End If
    
    cFeld = Text1(6).Text       'Liebesnr
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "B.LIBESNR like '" & cFeld & "' "
    End If
    
    If Check3.Value = vbChecked Then    'schon Kalkuliert
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " B.Spanne > 0 "
    End If

    cSQL = cSQL & cwhere
    cSQL = cSQL & " and GEFUEHRT = 'J' "
    cSQL = cSQL & "order by A.LINR, A.LPZ, A.BEZEICH "
    
    loeschNEW "artueb", gdBase
    
    sSQL = "Create Table ARTUEB ( "
    sSQL = sSQL & " ARTNR double"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", AGN double"
    sSQL = sSQL & ", LINR double"
    sSQL = sSQL & ", LPZ double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", EAN Text(13)"
    sSQL = sSQL & ", RKZ Text(1)"
    sSQL = sSQL & ", NOTIZEN Text(25)"
    sSQL = sSQL & ", BESTAND double"
    sSQL = sSQL & ", GEFUEHRT Text(1)"
    sSQL = sSQL & ", MWST Text(1)"
    sSQL = sSQL & ", KVKA double"
    sSQL = sSQL & ", KVKN double"
    sSQL = sSQL & ", DiffKVKE double"
    sSQL = sSQL & ", DiffKVKP double"
    sSQL = sSQL & ", ETIMERK Text(1)"
    sSQL = sSQL & ", LEK double"
    sSQL = sSQL & ", SEK double"
    sSQL = sSQL & ", NSA double"
    sSQL = sSQL & ", NSN double"
    sSQL = sSQL & ", AWM Text(2)"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTUEB " & cSQL
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTUEB set etimerk = 'N' where NSA = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTUEB set etimerk = 'N' where NSA is null "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKLad"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKLad()
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
    
    Set rsrs = gdBase.OpenRecordset("ARTUEB", dbOpenTable)
    
    MSHFLEX1.Redraw = False
    
    pbrZeit.Visible = True
    pbrZeit.Max = 100
    counter = 0
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If counter = 100 Then
                counter = 0
            End If
            counter = counter + 1
            pbrZeit.Value = counter
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Row = 0
                MSHFLEX1.Col = i
                
                If sSpaltenname(i) = MSHFLEX1.Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "KVK - Preis neu", "KVK - Preis alt", "Nettospanne alt", "Nettospanne neu"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow

                                MSHFLEX1.Text = Format$(sWert, "####0.00")

                        Case Is = "Diff KVK in Euro", "Diff KVK in %"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow

                                MSHFLEX1.Text = Format$(sWert, "###0.00")

                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = sWert
                            
                         If Not IsNull(rsrs!AWM) Then
                            sWert = rsrs!AWM
                            If Trim(sWert) = "96" Then
                                For j = 0 To byAnzahlSpalten - 1
                                    MSHFLEX1.Col = j
                                    MSHFLEX1.CellBackColor = vbWhite
                                    MSHFLEX1.CellForeColor = &HFF&
                                Next j
                            End If
                        End If
                            
                    End Select
                    
            
                    If Len(MSHFLEX1.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                        aBreite(i) = Len(MSHFLEX1.TextMatrix(lrow, i)) * 80
                    End If
                    
                End If
            Next i
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        MSHFLEX1.Col = i
        MSHFLEX1.ColWidth(i) = aBreite(i) * 1.8
    Next i
        
    
    rsrs.Close: Set rsrs = Nothing
    pbrZeit.Visible = False
    If byAnzahlSpalten < 2 Then
    
    Else
        MSHFLEX1.FixedCols = 1
    End If
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    If lrow > 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurden ermittelt."
        lblAnzeige.Refresh
    ElseIf lrow = 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurde ermittelt."
        lblAnzeige.Refresh
    Else
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Es wurden keine Artikel ermittelt."
        lblAnzeige.Refresh
        Exit Sub
    End If
    
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKLad"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub FormatMShFlex1WKLad()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsRL As Recordset
    
    Dim i As Byte
    Dim j As Byte
    
    sSQL = "Select * from TABLay" & srechnertab & " where ANZEIGE = 'J' and Tabname = 'ARTUEB' order by Reihenf"
    Set rsRL = gdBase.OpenRecordset(sSQL)
    
    If rsRL.EOF Then
    
    Else
        byAnzahlSpalten = rsRL.RecordCount
        ReDim sSpaltenname(byAnzahlSpalten)
        ReDim sSpaltenbez(byAnzahlSpalten)
        ReDim aBreite(byAnzahlSpalten)
        rsRL.MoveFirst
        i = 0
        Do While Not rsRL.EOF
            sSpaltenname(i) = rsRL!Spaltenna
            sSpaltenbez(i) = rsRL!Spaltenbez
            i = i + 1
            rsRL.MoveNext
        Loop
    End If
    rsRL.Close
    
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
            aBreite(j) = Len(.TextMatrix(0, j)) * 80
'            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex1WKLad"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub cmdHelp_Click()
On Error GoTo LOKAL_ERROR

    zeigeHilfe "KISSHELP", Me.Tag & ".doc", gcPfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Artnr"
    gstab = "ARTUEB"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim dAufschlag      As Double
    Dim dNettospanne    As Double
    Dim dAufschlagEuro  As Double
    Dim dErsetzungEuro  As Double
    Dim sSQL            As String
    Dim rsNs            As Recordset
    Dim dEK             As Double
    Dim cMWST           As String
    Dim dKVKN           As Double
    Dim cKVKN           As String
    Dim sKalkkrit       As String
    Dim lrow            As Long

    dNettospanne = 0
    If Text3.Text <> "" Then
        dNettospanne = CDbl(Text3.Text)
        If dNettospanne = 0 Then
            Text3.SetFocus
            Exit Sub
        End If
    End If
    
    dAufschlag = 0
    If Text2.Text <> "" Then
        dAufschlag = CDbl(Text2.Text)
        If dAufschlag = 0 Then
            Text2.SetFocus
            Exit Sub
        End If
    End If
    
    dAufschlagEuro = 0
    If Text4.Text <> "" Then
        dAufschlagEuro = CDbl(Text4.Text)
        If dAufschlagEuro = 0 Then
            Text4.SetFocus
            Exit Sub
        End If
    End If
    
    dErsetzungEuro = 0
    If Text5.Text <> "" Then
        dErsetzungEuro = CDbl(Text5.Text)
        If dErsetzungEuro = 0 Then
            Text5.SetFocus
            Exit Sub
        End If
    End If
    
    If Label3(8).Caption <> "" Then
        sKalkkrit = Label3(8).Caption
    Else
        sKalkkrit = "nicht festgelegt"
    End If

    loeschNEW "KALKHEAD", gdBase
    CreateTable "KALKHEAD", gdBase
    
    sSQL = "Insert into Kalkhead (NettoSProz,AufschlagProz,AufschlagEuro,ErsetzungEuro,Kalkkrit) "
    sSQL = sSQL & " values ('" & CStr(dNettospanne) & "','" & CStr(dAufschlag) & "','" & CStr(dAufschlagEuro) & "','" & CStr(dErsetzungEuro) & "','" & sKalkkrit & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from ARTUEB where Etimerk = 'J'"
    Set rsNs = gdBase.OpenRecordset(sSQL)
    
    If Not rsNs.EOF Then
        rsNs.MoveLast
    End If
    lrow = rsNs.RecordCount
    
    If lrow = 0 Then
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Keine Artikel zur Kalkulation ermittelt. Bitte in der Spalte 'kalkulieren' ein 'J' eintragen!"
        lblAnzeige.Refresh
        Exit Sub
    Else
        
        rsNs.MoveFirst
        Do While Not rsNs.EOF

            If Not IsNull(rsNs(gsSpanne)) Then
                dEK = Trim(CDbl(rsNs(gsSpanne)))
            Else
                dEK = "0"
            End If
            
            If Not IsNull(rsNs!MWST) Then
                cMWST = Trim(rsNs!MWST)
            Else
                cMWST = "V"
            End If
            
            If dAufschlag > 0 Then
                cKVKN = fnVKneu(dEK, cMWST, dAufschlag)             'über Aufschlag
            ElseIf dNettospanne > 0 Then
                cKVKN = fnVKneuNS(dEK, cMWST, dNettospanne)         'über Nettospanne
            ElseIf dAufschlagEuro > 0 Then
                cKVKN = CStr(CDbl(rsNs!KVKA) + dAufschlagEuro)      'über AufschlagEuro
            ElseIf dErsetzungEuro > 0 Then
                cKVKN = CStr(dErsetzungEuro)                        'über ErsetzungEuro
            End If
            
            rsNs.Edit
            rsNs!KVKN = cKVKN
            rsNs!DiffKVKE = rsNs!KVKN - rsNs!KVKA
            rsNs!DiffKVKP = rsNs!DiffKVKE * 100 / IIf(rsNs!KVKN = "0", "1", rsNs!KVKN)
            rsNs!nsn = NettospanneInProzent(cKVKN, CStr(dEK), cMWST)
            rsNs.Update
        
            rsNs.MoveNext
        Loop
    End If
    rsNs.Close
    
    FuellenMShFlex1WKLad
    
    If lrow > 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurden kalkuliert (berechnet)."
        lblAnzeige.Refresh
    ElseIf lrow = 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurde kalkuliert (berechnet)."
        lblAnzeige.Refresh
    Else
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Es wurden keine Artikel kalkuliert (berechnet)."
        lblAnzeige.Refresh
        Exit Sub
    End If
    
    Command4.Enabled = True

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnVKneu(dEK As Double, cMW As String, dAufschlag As Double) As String
    On Error GoTo LOKAL_ERROR
    
    Dim dVKNEU1    As Double
    Dim dVKNEU2    As Double
    Dim dVKNEU3    As Double
    Dim dVKNEU4    As Double
    Dim dVKNEU5    As Double

    fnVKneu = "0"
     
    dVKNEU1 = (dEK * dAufschlag) / 100
    dVKNEU2 = dVKNEU1 + dEK
    
    If cMW = "V" Then
        dVKNEU3 = dVKNEU2 * 1.16
    ElseIf cMW = "E" Then
        dVKNEU3 = dVKNEU2 * 1.07
    Else
        dVKNEU3 = dVKNEU2
    End If
    
    fnVKneu = Format$(dVKNEU3, "#####0.00")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnVKneu"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim sArtnr  As String
    Dim rsDu    As Recordset
    Dim lrow    As Long
    
    Screen.MousePointer = 11
    
    frmWKL35.SetFocus
    
    lblAnzeige.ForeColor = glS1
    lblAnzeige.Caption = "Die neuen Kassenverkaufspreise werden jetzt übernommen..."
    lblAnzeige.Refresh

    MSHFLEX1.Redraw = False
    
    updateARTUEB
    FuellenMShFlex1WKLad
    Nochmalrechnen

    sSQL = " Update Artikel inner join ARTUEB ON "
    sSQL = sSQL & " Artikel.ARTNR = ARTUEB.ARTNR  "
    sSQL = sSQL & " Set Artikel.LASTDATE = DateValue(now),Artikel.KVKPR1 = ARTUEB.KVKN, Artikel.Spanne = ARTUEB.NSN, Artikel.Etimerk = ARTUEB.Etimerk,Artikel.SYNSTATUS = 'E' "
    sSQL = sSQL & " where ARTUEB.ETIMERK = 'J' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    If Check2.Value = vbChecked Then
    
        sSQL = " Update ARTLIEF inner join ARTUEB ON "
        sSQL = sSQL & " ARTLIEF.ARTNR = ARTUEB.ARTNR and ARTLIEF.LINR = ARTUEB.LINR "
        sSQL = sSQL & " Set ARTLIEF.Spanne = ARTUEB.NSN"
        sSQL = sSQL & " where ARTUEB.ETIMERK = 'J' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    Else
    
        sSQL = " Update ARTLIEF inner join ARTUEB ON "
        sSQL = sSQL & " ARTLIEF.ARTNR = ARTUEB.ARTNR and ARTLIEF.LINR = ARTUEB.LINR "
        sSQL = sSQL & " Set ARTLIEF.Spanne = 0"
        sSQL = sSQL & " where ARTUEB.ETIMERK = 'J' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    sSQL = " Update ARTLIEF inner join ARTUEB ON "
    sSQL = sSQL & " ARTLIEF.ARTNR = ARTUEB.ARTNR and ARTLIEF.LINR = ARTUEB.LINR "
    sSQL = sSQL & " Set ARTLIEF.Spanne = 0 "
    sSQL = sSQL & " where ARTUEB.ETIMERK = 'N' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete Etidru.* from Etidru inner join ARTUEB ON "
    sSQL = sSQL & " Etidru.ARTNR = ARTUEB.ARTNR "
    sSQL = sSQL & " where Etidru.filnr = " & gcFilNr
    sSQL = sSQL & " and ARTUEB.ETIMERK = 'J' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Etidru inner join ARTUEB ON "
    sSQL = sSQL & " Etidru.ARTNR = ARTUEB.ARTNR "
    sSQL = sSQL & " Set Etidru.VKPR = ARTUEB.KVKN"
    sSQL = sSQL & " where ARTUEB.ETIMERK = 'J' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Etidru Select ARTNR,Bezeich,BESTAND,Bestand as anzahl, KVKN as VKPR,Libesnr,ean,lpz,linr, "
    sSQL = sSQL & gcFilNr & " as filnr, "
    sSQL = sSQL & " '" & srechnertab & "' as pcname  from artueb "
    sSQL = sSQL & " where ARTUEB.ETIMERK = 'J' "
    sSQL = sSQL & " and ARTUEB.KVKN <> ARTUEB.KVKA "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update Etidru inner join ARTIKEL ON "
    sSQL = sSQL & " Etidru.ARTNR = ARTIKEL.ARTNR "
    sSQL = sSQL & " Set Etidru.Bestand = ARTIKEL.Bestand"
    sSQL = sSQL & " where Etidru.filnr = " & gcFilNr
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTUEB "
    sSQL = sSQL & " where "
    sSQL = sSQL & " ETIMERK = 'N' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    MSHFLEX1.Redraw = True
    
    If Not NewTableSuchenDBKombi("KALKHEAD", gdBase) Then
        sSQL = "Update Kalkhead set Rundkrit = 'nicht gerundet'"
        sSQL = sSQL & " where Rundkrit is null "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set abru = 'nicht gerundet'"
        sSQL = sSQL & " where abru is null "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set aufru = 'nicht gerundet'"
        sSQL = sSQL & " where aufru is null "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set AufschlagEuro = null"
        sSQL = sSQL & " where AufschlagEuro = '0' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set ErsetzungEuro = null"
        sSQL = sSQL & " where ErsetzungEuro = '0' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set AufschlagProz = null"
        sSQL = sSQL & " where AufschlagProz = '0' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update Kalkhead set NettoSProz = null"
        sSQL = sSQL & " where NettoSProz = '0' "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    Set rsDu = gdBase.OpenRecordset("ARTUEB")
    If Not rsDu.EOF Then
        rsDu.MoveLast
    End If
    lrow = rsDu.RecordCount
    If lrow = 0 Then
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Es wurde keine Artikel kalkuliert."
        lblAnzeige.Refresh
    Else
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = "Druckvorschau wird erstellt..."
        lblAnzeige.Refresh
        reportbildschirm "dWKL35", "aWKL35"  'nur Access
        
        If lrow > 1 Then
            lblAnzeige.ForeColor = glS1
            lblAnzeige.Caption = lrow & " Artikel wurden kalkuliert übernommen."
            lblAnzeige.Refresh
        ElseIf lrow = 1 Then
            lblAnzeige.ForeColor = glS1
            lblAnzeige.Caption = lrow & " Artikel wurde kalkuliert übernommen."
            lblAnzeige.Refresh
        
        End If

    End If
    
    rsDu.Close
    
   
    
    Screen.MousePointer = 0
    
    MSHFLEX1.Visible = False
    Frame1.Visible = False
    Command6.Visible = False
    
    Urzustandherstellen
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsNs            As Recordset
    Dim dEK             As Double
    Dim cMWST           As String
    Dim dKVKN           As Double
    Dim cKVKN           As String
    Dim lrow            As Long
    
    
    sSQL = "Select * from ARTUEB where Etimerk = 'J'"
    Set rsNs = gdBase.OpenRecordset(sSQL)
    
    If Not rsNs.EOF Then
        rsNs.MoveLast
    End If
    
    lrow = rsNs.RecordCount
    
    If lrow = 0 Then
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Keine Artikel zur Kalkulation ermittelt. Bitte in der Spalte 'kalkulieren' ein 'J' eintragen!"
        lblAnzeige.Refresh
        Exit Sub
    Else
        rsNs.MoveFirst
        Do While Not rsNs.EOF

            If Not IsNull(rsNs!KVKN) Then
                dKVKN = CDbl((rsNs!KVKN))
            Else
                dKVKN = 0
            End If
            
            If Not IsNull(rsNs(gsSpanne)) Then
                dEK = Trim(CDbl(rsNs(gsSpanne)))
            Else
                dEK = "0"
            End If
            
            If Not IsNull(rsNs!MWST) Then
                cMWST = Trim(rsNs!MWST)
            Else
                cMWST = "V"
            End If
            
            rsNs.Edit
            rsNs!KVKN = Runden(dKVKN)
            
            rsNs!DiffKVKE = rsNs!KVKN - rsNs!KVKA
            rsNs!DiffKVKP = rsNs!DiffKVKE * 100 / IIf(rsNs!KVKN = "0", "1", rsNs!KVKN)
            rsNs!nsn = NettospanneInProzent(rsNs!KVKN, CStr(dEK), cMWST)
            rsNs.Update
            
            rsNs.MoveNext
        Loop
    End If
    rsNs.Close
    
    FuellenMShFlex1WKLad
    
    If lrow > 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurden kalkuliert (gerundet)."
        lblAnzeige.Refresh
    ElseIf lrow = 1 Then
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = lrow & " Artikel wurde kalkuliert (gerundet)."
        lblAnzeige.Refresh
    Else
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Es wurden keine Artikel kalkuliert (gerundet)."
        lblAnzeige.Refresh
'        Exit Sub
    End If
    
    sSQL = "Update Kalkhead set Rundkrit = '" & Label3(7).Caption & "'"
    sSQL = sSQL & ", abru = '" & Label3(6).Caption & "'"
    sSQL = sSQL & ", aufru = '" & Label3(5).Caption & "'"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub Nochmalrechnen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsNs            As Recordset
    Dim dEK             As Double
    Dim cMWST           As String
    Dim dKVKN           As Double
    Dim cKVKN           As String
    Dim lrow            As Long
    
    sSQL = "Select * from ARTUEB where Etimerk = 'J'"
    Set rsNs = gdBase.OpenRecordset(sSQL)
    
    If Not rsNs.EOF Then

        rsNs.MoveFirst
        Do While Not rsNs.EOF

            If Not IsNull(rsNs!KVKN) Then
                dKVKN = CDbl((rsNs!KVKN))
            Else
                dKVKN = 0
            End If
            
            If Not IsNull(rsNs(gsSpanne)) Then
                dEK = Trim(CDbl(rsNs(gsSpanne)))
            Else
                dEK = "0"
            End If
            
            If Not IsNull(rsNs!MWST) Then
                cMWST = Trim(rsNs!MWST)
            Else
                cMWST = "V"
            End If
            
            rsNs.Edit
            
            rsNs!DiffKVKE = rsNs!KVKN - rsNs!KVKA
            rsNs!DiffKVKP = rsNs!DiffKVKE * 100 / IIf(rsNs!KVKN = "0", "1", rsNs!KVKN)
            rsNs!nsn = NettospanneInProzent(rsNs!KVKN, CStr(dEK), cMWST)
            rsNs.Update
            
            rsNs.MoveNext
        Loop
    End If
    rsNs.Close
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Nochmalrechnen"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    iRet = MsgBox("Möchten Sie wirklich die Preiskalkulation beenden?", vbQuestion + vbYesNo, "ENDE")
    If iRet = vbYes Then
        Unload frmWKL35
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    lblAnzeige.ForeColor = glS1
    lblAnzeige.Caption = "Druckvorschau wird erstellt..."
    lblAnzeige.Refresh
    reportbildschirm "dWKL35", "aWKL35a"
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL35
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    cmdHelp.Visible = istHilfeda(Me.Tag)
    
    Label3(5).Caption = giAufrunden
    Label3(6).Caption = giAbrunden
    Label3(7).Caption = giRundkrit
    Label3(8).Caption = IIf(gsSpanne = "LEK", "List - EK", "Schnitt - EK")
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL35()
    On Error GoTo LOKAL_ERROR
    
    MSHFLEX1.Top = 1500
    MSHFLEX1.Left = 480
    MSHFLEX1.Width = 10815
    MSHFLEX1.Height = 5015
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL35"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Index = 7 Then
        If Len(Text1(4).Text) = 0 Then
            Text1(4).SetFocus
            lblAnzeige.ForeColor = vbRed
            lblAnzeige.Caption = "Wenn Sie eine Linie wählen möchten, dann müssen Sie erst einen Lieferanten eingeben!"
            lblAnzeige.Refresh
        End If
    End If
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    MSHFLEX1.Visible = False
    Frame1.Visible = False
    Command6.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim bSpringen As Boolean
    
    bSpringen = False
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    If Len(Text1(Index).Text) = Text1(Index).MaxLength - 1 Then
        If KeyAscii <> 8 And KeyAscii <> 0 And KeyAscii <> 13 And KeyAscii <> 27 Then
            bSpringen = True
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_Click()
    On Error GoTo LOKAL_ERROR
    
    posimerk
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    sortierenHGrid MSHFLEX1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_DblClick"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_GotFocus()
    On Error GoTo LOKAL_ERROR

    posimerk
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    
    lrow = MSHFLEX1.Row
    lcol = MSHFLEX1.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSHFLEX1.Rows Then
        lrow = lrow - 1
    End If
    
    If KeyCode = &H28 Or KeyCode = &H27 Or KeyCode = &H26 Or KeyCode = &H25 Then
        Exit Sub
    End If
    
    If iKeypress = 0 And KeyCode <> vbKeyBack Then
        
        If KeyCode <> 46 Then
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = lcol
            MSHFLEX1.Text = ""
        End If
    End If
    iKeypress = iKeypress + 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyDown"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen    As String
    Dim cValid      As String
    Dim cArtNr      As String
    Dim lcol        As Long
    Dim lrow        As Long
    Dim bKalk       As Boolean
    Dim sSQL        As String
    Dim i           As Integer
    
    bKalk = False
    posimerk
    
    cZeichen = Chr$(KeyAscii)
    
    MSHFLEX1.Row = 0
    Select Case MSHFLEX1.Text
        Case Is = "kalkulieren"
            bKalk = True
            If cZeichen = "J" Then
                KeyAscii = 74
            ElseIf cZeichen = "N" Then
                KeyAscii = 78
            ElseIf cZeichen = "n" Then
                KeyAscii = 110
            ElseIf cZeichen = "j" Then
                KeyAscii = 106
            Else
                KeyAscii = 0
            End If
            
        Case Is = "KVK - Preis neu"
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = "Bestand"
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
    
    lcol = Val(Label0(1).Caption)
    lrow = Val(Label0(0).Caption)
    MSHFLEX1.Row = lrow
    MSHFLEX1.Col = lcol
    
    If KeyAscii <> 0 Then
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        cValid = MSHFLEX1.Text
        If InStr(cValid, ",") > 0 And cZeichen = "," Then
            KeyAscii = 0
        End If
        
        If KeyAscii <> 0 Then
            If KeyAscii <> 8 Then
                If bKalk Then
                    
'                    If Len(cValid) > 0 Then
                        cValid = Chr$(KeyAscii)
                        cValid = UCase(cValid)
                        For i = 0 To byAnzahlSpalten - 1
                            MSHFLEX1.Row = 0
                            MSHFLEX1.Col = i
                            If MSHFLEX1.Text = "Artnr" Then
                                MSHFLEX1.Row = lrow
                                cArtNr = MSHFLEX1.Text
                                Exit For
                            End If
                        Next i
                        sSQL = "Update artueb set Etimerk = '" & cValid & "'"
                        sSQL = sSQL & " where artnr = " & cArtNr
                        gdBase.Execute sSQL, dbFailOnError
                        
                        MSHFLEX1.Row = lrow
                        MSHFLEX1.Col = lcol
'                    End If
                Else
                    cValid = cValid & Chr$(KeyAscii)
                End If
            Else
                If Len(cValid) > 0 Then
                    cValid = Left(cValid, Len(cValid) - 1)
                End If
            End If
            
            MSHFLEX1.Text = UCase(cValid)
        End If
        
        bKalk = False
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    lrow = MSHFLEX1.Row
    
    Select Case KeyCode
        Case Is = 46    'Del
            MSHFLEX1.Row = 0
            Select Case MSHFLEX1.Text
                Case Is = "kalkulieren", "KVK - Preis neu", "Bestand"
                MSHFLEX1.Row = lrow
                MSHFLEX1.Text = ""
                Case Else
                MSHFLEX1.Row = lrow
            End Select
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyUp"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_LeaveCell()
    On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_LeaveCell"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub updateARTUEB()
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim lcol            As Long
    Dim lrow            As Long
    Dim cArtNr          As String
    Dim cKVKN           As String
    Dim cBestand        As String
    Dim rsKalk          As Recordset
    Dim sSQL            As String
    Dim cKalk           As String
    Dim dWert           As Double
    

    For j = 2 To MSHFLEX1.Rows - 1
        For i = 0 To byAnzahlSpalten - 1
            MSHFLEX1.Row = 0
            MSHFLEX1.Col = i
            If MSHFLEX1.Text = "Artnr" Then
                MSHFLEX1.Row = j
                cArtNr = MSHFLEX1.Text
                
                For k = 0 To byAnzahlSpalten - 1
                    MSHFLEX1.Row = 0
                    MSHFLEX1.Col = k
                    If MSHFLEX1.Text = "Bestand" Then
                        MSHFLEX1.Row = j
                        cBestand = MSHFLEX1.Text
                        Bestandsveraenderung cArtNr, CLng(cBestand), "Auto.Kalkulation"
                    ElseIf MSHFLEX1.Text = "KVK - Preis neu" Then
                        MSHFLEX1.Row = j
                        cKVKN = MSHFLEX1.Text
                        
                        If cKVKN <> "" Then
                            cKVKN = fnMoveComma2Point$(cKVKN)
                            
                            sSQL = "Update ARTUEB set KVKN "
                            sSQL = sSQL & " = " & cKVKN
                            sSQL = sSQL & " where artnr = " & cArtNr
                            gdBase.Execute sSQL, dbFailOnError
                        End If
                        
                    ElseIf MSHFLEX1.Text = "kalkulieren" Then
                        MSHFLEX1.Row = j
                        cKalk = MSHFLEX1.Text
                        
                        If cKalk <> "" Then
                            sSQL = "Update ARTUEB set Etimerk "
                            sSQL = sSQL & " = '" & cKalk & "' "
                            sSQL = sSQL & " where artnr = " & cArtNr
                            gdBase.Execute sSQL, dbFailOnError
                        End If
                    End If
                Next k
                
            End If
        Next i
    Next j

    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "updateARTUEB"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub posimerk()
    On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row = 1 Then
        Label0(0).Caption = "1"
        Exit Sub
    End If
    
    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSHFLEX1.Row
    lcol = MSHFLEX1.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSHFLEX1.Rows Then
        lrow = lrow - 1
    End If
   
    MSHFLEX1.Row = lrow
    MSHFLEX1.Col = lcol
    
    Label0(0).Caption = Trim$(Str$(lrow))
    Label0(1).Caption = Trim$(Str$(lcol))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "posimerk"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyReturn Then
        cmdGo_Click
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        Select Case Index
            Case Is = 4
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
                
            Case 7
                ctmp = Text1(4).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Text1(4).SetFocus
                    Exit Sub
                End If
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = ctmp
                    
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                Text1(7).Text = Trim(gF2Prompt.cWahl)

        End Select
        Text1(Index).SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus()
    On Error GoTo LOKAL_ERROR

    Text2.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Text2_Change()
    On Error GoTo LOKAL_ERROR
    
    If Len(Text2.Text) > 0 Then
        Command2.Enabled = True
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
    Else
        Command2.Enabled = False
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_Change"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = glSelBack1
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_Change()
    On Error GoTo LOKAL_ERROR

    If Len(Text3.Text) > 0 Then
        Command2.Enabled = True
        Text2.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
    Else
        Command2.Enabled = False
        Command4.Enabled = False
        Text2.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_Change"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text3.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text2.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_LostFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = glSelBack1
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_Change()
    On Error GoTo LOKAL_ERROR

    If Len(Text4.Text) > 0 Then
        Command2.Enabled = True
        Text2.Enabled = False
        Text3.Enabled = False
        Text5.Enabled = False
    Else
        Command2.Enabled = False
        Command4.Enabled = False
        Text2.Enabled = True
        Text3.Enabled = True
        Text5.Enabled = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_Change"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text4.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text5_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_LostFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = glSelBack1
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_Change()
    On Error GoTo LOKAL_ERROR

    If Len(Text5.Text) > 0 Then
        Command2.Enabled = True
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
    Else
        Command2.Enabled = False
        Command4.Enabled = False
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_Change"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text5.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Im Kalkulator ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

