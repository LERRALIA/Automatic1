VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL186 
   BackColor       =   &H00C0C000&
   Caption         =   "neue XXX - Artikeldaten"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL186.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   960
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   26
      Top             =   2760
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
      Caption         =   "Historie"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "�bernahmeoptionen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8880
      TabIndex        =   20
      Top             =   4440
      Width           =   2775
      Begin sevCommand3.Command Command5 
         Height          =   405
         Index           =   9
         Left            =   1200
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
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
      Begin VB.CheckBox Check1 
         Caption         =   "nicht enthaltene Artikel = EX -Artikel (Auslistung)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtAGN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "AGN f�r neue Artikel"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   18
      Top             =   3360
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
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   17
      Top             =   3840
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
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   16
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
      Caption         =   "Protokoll"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   10200
      Pattern         =   "MASTER!.*"
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   7680
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
      Caption         =   "Schlie�en"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   0
      Top             =   6480
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
      Caption         =   "Einlesen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   27
      Top             =   1440
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
      Caption         =   "L�schen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label lblDatei 
      Caption         =   "Dateiname"
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
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   9255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "n�here Informationen hier: (bitte anklicken)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5400
      MouseIcon       =   "frmWKL186.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   15
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   6960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   10
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Artikel insgesamt"
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
      Index           =   9
      Left            =   1440
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   7
      Left            =   4080
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "0"
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
      Index           =   6
      Left            =   4080
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "EK Preis�nderung"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Auslistung / Ex "
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "neue Artikel"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "M�chten Sie diese �bernehmen, so klicken Sie auf ""Einlesen""."
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
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neue XXX-Artikeldaten stehen bereit. "
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
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "neue XXX - Artikeldaten"
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
      TabIndex        =   3
      Top             =   120
      Width           =   11535
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
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL186"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glartv      As Long
Dim glartb      As Long
Dim iSec        As Integer


Private Function XXXDatenEinlesen_neu(sPfad As String, sDatei As String) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim rsagn   As Recordset
    
    XXXDatenEinlesen_neu = False
    
    loeschNEW "REWE_AGN", gdApp
    CreateTableT2 "REWE_AGN", gdApp
    
    'vorbereitung der Importtabelle
    '1. erst l�schen

    loeschNEW "MEISTER", gdApp
    CreateTable "MEISTER", gdApp

    Dim iFileNr As Integer
    Dim cSatz1 As String
    Dim lPos As Long
    Dim cEinzelsatz As String
    Dim lLinr As Long
    Dim lfnr1 As Long
    Dim cBezeich As String
    Dim lPosEnde As Long
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lposSemiEnde As Long
    Dim cWert As String
    
    Dim cAgn As String
    Dim cAGNBEZEICH As String
    Dim cWarengruppe As String
    
    Label1(2).Visible = True
    Label1(3).Visible = True
    Label1(5).Visible = True
    Label1(9).Visible = True
    Label1(6).Visible = True
    Label1(7).Visible = True
    Label1(8).Visible = True
    Label1(10).Visible = True
    
    Command5(3).Enabled = False
    Command5(6).Enabled = False
    
    lfnr1 = 0
    lPos = 1
    lPosEnde = 1
    lposSemiEnde = 1

    Set rsrs = gdApp.OpenRecordset("MEISTER")
    Set rsagn = gdApp.OpenRecordset("REWE_AGN")

    iFileNr = FreeFile
    Open sPfad & "\" & sDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then

        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1


        lLenfil = Len(cSatz1)

        lPosEnde = InStr(lPos, cSatz1, vbCrLf)
        

        Do
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)

            cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
            lPos = lPos + lPosEnde - lPos + 2

            lposSemi = 1

            rsrs.AddNew
            rsagn.AddNew
            
            lfnr1 = lfnr1 + 1
            rsrs!lfnr = lfnr1
            
            'WGRU = AGN
            cWert = ""
            cAgn = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cAgn = cWert
            rsrs!AGN = CLng(cAgn)
            rsagn!RAGN = CLng(cAgn)

            'WGRU_BEZ
            cWert = ""
            cAGNBEZEICH = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            cAGNBEZEICH = Left(cWert, 30)
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "�", "�")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "!", "�")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "\", "�")
            rsagn!RAGTEXT = cAGNBEZEICH

            'Libesnr
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!LIBESNR = cWert
            
            'EAN
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!EAN = cWert
            
            'Bezeich
            cWert = ""
            cBezeich = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cBezeich = cWert
            
            While InStr(1, cBezeich, "  ")
                cBezeich = SwapStr(cBezeich, "  ", " ")
            Wend

            rsrs!BEZEICH = Left(cBezeich, 35)
            
            'Einh = Minmen
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!MINMEN = cWert

            'EK = EKPR
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If IsNumeric(cWert) Then
                rsrs!lekpr = cWert
            Else
                rsrs!lekpr = 0
            End If

            'UVP = VKPR
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If IsNumeric(cWert) Then
                rsrs!vkpr = cWert
            Else
                rsrs!vkpr = 0
            End If
            
            'INH_Wert
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If cWert <> "" Then
                rsrs!INHALT = CLng(cWert)
            Else
                rsrs!INHALT = 0
            End If
            
            'INH_BEZ
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            rsrs!INHALTBEZ = Left(cWert, 3)
            
            'MwSt
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            
            If lposSemiEnde > 0 Then
                cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi + 1)
                cWert = Val(cWert)
                Select Case cWert
                    Case "0"
                        rsrs!MWST = "O"
                    Case "1"
                        rsrs!MWST = "E"
                    Case "2"
                        rsrs!MWST = "V"
                    Case Else
                        rsrs!MWST = "V"
                End Select
                'UVP-Pflicht
                cWert = ""
                cWert = Right(cEinzelsatz, 1)
                If cWert = "1" Then
                    rsrs!Status = "J"
                Else
                    rsrs!Status = "N"
                End If
            Else
                cWert = ""
                cWert = Right(cEinzelsatz, 1)
                Select Case cWert
                    Case "0"
                        rsrs!MWST = "O"
                    Case "1"
                        rsrs!MWST = "E"
                    Case "2"
                        rsrs!MWST = "V"
                    Case Else
                        rsrs!MWST = "V"
                End Select
            
                rsrs!Status = "N"
            End If
            
            rsrs!NOTIZEN = ""
            rsrs!RKZ = "N"
            rsrs!EAN2 = ""
            rsrs!EAN3 = ""
            rsrs!GRUNDPREIS = "J"
'            rsrs!MARKE = ""
            rsrs!LPZ = 1
            rsrs!linr = 0
            rsrs!GEFUEHRT = "J"
            
            cWarengruppe = ""

            Select Case Left(cAgn, 2)
                Case "09", "13", "15", "16", "21", "26", "27", "28", "31", "32", "33", "34", "35", "36"
                    cWarengruppe = "Lebensmittel"
                Case "29", "68"
                    cWarengruppe = "Babynahrung/-bedarf"
                Case "18", "72"
                    cWarengruppe = "Gartenbedarf"
                Case "30", "53", "54", "56", "57", "58", "59", "61", "62", "74", "75", "76", "77", "82"
                    cWarengruppe = "Drogeriewaren"
                Case "38", "39", "40", "64"
                    cWarengruppe = "S��waren, Chips"
                Case "42", "43", "44", "45", "46", "47", "48"
                    cWarengruppe = "Getr�nke"
                Case "49", "50", "51"
                    cWarengruppe = "Kaffee/Tee/Kakao/Tabak"
                Case "65"
                    cWarengruppe = "Tiernahrung/-bedarf"
                Case "95", "96", "98", "99"
                    cWarengruppe = "Sonstiges"
                Case Else
                    cWarengruppe = "unbekannt"
            End Select
            rsrs!MARKE = cWarengruppe
            
            rsagn.Update
            rsrs.Update
            
            Label1(10).Caption = CInt(Label1(10).Caption) + 1
            Label1(10).Refresh

        Loop While lLenfil >= lPos
    End If

    Close iFileNr
    rsrs.Close: Set rsrs = Nothing
    rsagn.Close: Set rsagn = Nothing
    
    '****************
    'hier Reweagn's auf neue pr�fen

    loeschNEW "REWE_AGN_DIS", gdApp

    sSQL = "Select ragn as agn, ragtext as agtext into REWE_AGN_DIS from REWE_AGN group by ragn,ragtext "
    gdApp.Execute sSQL, dbFailOnError

    loeschNEW "REWE_AGN", gdApp

    loeschNEW "REWE_AGN_DIS", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "REWE_AGN_DIS"

    sSQL = "Delete from REWE_AGN_DIS where REWE_AGN_DIS.agn in (Select agn from agndbf)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into agndbf select agn,agtext from REWE_AGN_DIS"
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "REWE_AGN_DIS", gdBase

    'Ende hier Reweagn's auf neue pr�fen

    
    anzeige "normal", "EAN Duplikate werden entfernt...", Label1(4)

    '6. EAN - Duplikats�berpr�fung in der Importtabelle Anzahl ermitteln
    ErmittlungReweDuplisPlusDel
        
    sSQL = "Delete from Meister where LEKPR = 0 "
    gdApp.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Die Artikeldaten werden �berpr�ft...", Label1(4)
   
    '7. diverse Feld�berpr�fungen vornehmen
    feldcheckRewe
    anzeige "normal", "Die Sortiments�bersicht wird erstellt...", Label1(4)

    '8. Tabelle IMPORTPRI zur Datenbank kopieren
    loeschNEW "ImportpriREWE", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "ImportpriREWE"

    sSQL = "Create Index ARTNR on ImportpriREWE (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ImportpriREWE set marke ='' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    '9. Ermittle LINR und ordne zu
    anzeige "normal", "Die Lieferantenzuordnung wird �berpr�ft...", Label1(4)
    
    lLinr = checkLinrForKISS(Label1(4))
    
    If lLinr = 0 Then
        Screen.MousePointer = 0
        anzeige "rot", "Keine auswertbaren Lieferantennummern enthalten.", Label1(4)
        Exit Function
    Else
        sSQL = "Update ImportpriREWE SET LINR = " & lLinr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Check1.Value = vbChecked Then 'alle anderen sind EX
    
        sSQL = "Update Artlief set RKZ = 'J'"
        sSQL = sSQL & ", EXDAT = '" & DateValue(Now) & "' "
        sSQL = sSQL & " where LINR = " & lLinr & " "
        gdBase.Execute sSQL, dbFailOnError
    
    
    
    
'        sSQL = "Update Artikel inner join Artlief on Artikel.artnr = artlief.artnr "
'        sSQL = sSQL & " set Artikel.RKZ = 'J'"
'        sSQL = sSQL & " where Artlief.LINR = " & lLinr & " "
'        gdBase.Execute sSQL, dbFailOnError
    End If

    FormatiereBildschirmdatenXXX
    
    schreibeStreckeProtokoll "Datei: " & lblDatei.Caption & " eingelesen, gew�hlter Lieferant: " & lLinr & " "
    
    'jetzt �bernehmen
    Uebernahme_XXX_Delta_neu sDatei, lLinr
    
    If NewTableSuchenDBKombi("ImportpriREWE_dauer", gdBase) = False Then
        sSQL = "select * into ImportpriREWE_dauer from ImportpriREWE  "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Insert into ImportpriREWE_dauer select * from ImportpriREWE  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Command5(2).Enabled = True
    Command5(3).Enabled = True
    Command5(6).Enabled = True
    
    anzeige "normal", "Fertig! Artikel�bernahme beendet, Protokoll und Etiketten ausdrucken!", Label1(4)
    
    XXXDatenEinlesen_neu = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "XXXDatenEinlesen_neu"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Function
Private Sub Uebernahme_XXX_Delta_neu(sDatei As String, lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsArtlief       As DAO.Recordset
    Dim sArtnr          As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen...", Label1(4)
    
    loeschNEW "STADAPROSTRECKE", gdBase
    CreateTableT2 "STADAPROSTRECKE", gdBase
    
    'Ex Artikel, und Exdatum
    cSQL = "Update Artikel inner join ImportpriREWE i on artikel.artnr = i.artnr "
    cSQL = cSQL & " set "
    cSQL = cSQL & " artikel.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", artikel.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", artikel.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artlief inner join ImportpriREWE i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError
    
   
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(1)...", Label1(4)
    
    
    If Val(Label1(7).Caption) > 0 Then
        Command5(3).Visible = True
        Command5(3).Enabled = False
    End If
    
    'Ex Artikel
    'Protokoll f�llen
    cSQL = "Insert into STADAPROSTRECKE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Auslistung/Ex' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
    cSQL = cSQL & " from ImportpriREWE where rkz = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(2)...", Label1(4)
    
    'neue Artikel
    cSQL = "Select * from ImportpriREWE where awm = '98'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = Trim(rsrs!artnr)
            End If
            
            Sicherheitsl�schen sArtnr 'artlief

            rsrs.MoveNext
            
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(3)...", Label1(4)
    
    cSQL = "Insert into Artikel Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", AGN "
    cSQL = cSQL & ", PGN "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", MWST "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", EAN2"
    cSQL = cSQL & ", EAN3 "
    cSQL = cSQL & ", ETIMERK "
'    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", INHALT "
    cSQL = cSQL & ", INHALTBEZ "
    cSQL = cSQL & ", GRUNDPREIS "
    cSQL = cSQL & ", MINBEST "
    cSQL = cSQL & ", 'J' as RABATT_OK "
    cSQL = cSQL & ", 'J' as GEFUEHRT "
    cSQL = cSQL & ", EKPR "
    cSQL = cSQL & ", 'N' as PREISSCHU "
    cSQL = cSQL & ", 'J' as BONUS_OK "
    cSQL = cSQL & ", 'J' as UMS_OK "
    cSQL = cSQL & ", AWM "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as AUFDAT "
    cSQL = cSQL & ", EXDAT "
    cSQL = cSQL & ", GROESSE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", NOTIZEN "
    cSQL = cSQL & ", KVKNEU as KVKPR1 "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(4)...", Label1(4)
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(5)...", Label1(4)
    
    
    'Neuheiten
    'Protokoll f�llen
    cSQL = "Insert into STADAPROSTRECKE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Neuheiten/wieder verf�gbar' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
'    cSQL = cSQL & ", VKPR_ALT real"
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
'    cSQL = cSQL & ", LEK_ALT real"
'    cSQL = cSQL & ", LEK_NEW real"
'    cSQL = cSQL & ", RKZ "
'    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
'    cSQL = cSQL & ", AGTEXT varchar(30)"
'    cSQL = cSQL & ", AGN int "
    cSQL = cSQL & " from ImportpriREWE where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(6)...", Label1(4)
    
    'EK-Preis�nderungen
    'Protokoll f�llen
    cSQL = "Insert into STADAPROSTRECKE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'EK-Preis�nderungen' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
'    cSQL = cSQL & ", VKPR_ALT real"
'    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
    cSQL = cSQL & ", LEKALT as LEK_ALT "
    cSQL = cSQL & ", LEKNEU as LEK_NEW "
'    cSQL = cSQL & ", RKZ "
'    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
'    cSQL = cSQL & ", AGTEXT varchar(30)"
'    cSQL = cSQL & ", AGN int "
    cSQL = cSQL & " from ImportpriREWE where LEKALT <> 0 and RKZ = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(7)...", Label1(4)
    
    'VK-Preisvorschl�ge
    'Protokoll f�llen
    cSQL = "Insert into STADAPROSTRECKE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'LVK-Preisvorschl�ge' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
'    cSQL = cSQL & ", VKPR_ALT real"
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
'    cSQL = cSQL & ", LEKALT as LEK_ALT "
'    cSQL = cSQL & ", LEKNEU as LEK_NEW "
'    cSQL = cSQL & ", RKZ "
'    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
'    cSQL = cSQL & ", AGTEXT varchar(30)"
'    cSQL = cSQL & ", AGN int "
    cSQL = cSQL & " from ImportpriREWE where KVKALT <> KVKNEU and mnotizen = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    'alle anderen �nderungen + Artliefeintrag
    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EAN  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EAN <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    
    

    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(8)...", Label1(4)
    
    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.BEZEICH = i.BEZEICH "
    cSQL = cSQL & " , a.LIBESNR = i.LIBESNR "
    cSQL = cSQL & " , a.EAN = i.EAN "
    cSQL = cSQL & " , a.MWST = i.MWST "
    cSQL = cSQL & " , a.MINMEN = i.MINMEN "
    cSQL = cSQL & " , a.INHALT = i.INHALT "
    cSQL = cSQL & " , a.INHALTBEZ = i.INHALTBEZ "
    cSQL = cSQL & " , a.AGN = i.AGN "
    cSQL = cSQL & " , a.NOTIZEN = i.NOTIZEN "
    cSQL = cSQL & " , a.lekpr = i.lekneu "
    cSQL = cSQL & " , a.rkz = i.rkz "
    cSQL = cSQL & " , a.VKPR = i.KVKNEU "
    
    cSQL = cSQL & ", a.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", a.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", a.SYNSTATUS = 'E' "
    
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(9)...", Label1(4)
    
    
    cSQL = "Delete from artlief where artnr in (Select artnr from ImportpriREWE) and Linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", NULL as EXDAT "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'E' as SYNSTATUS "
    cSQL = cSQL & " from ImportpriREWE " 'where awm <> '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artlief inner join ImportpriREWE i on Artlief.artnr = i.artnr  "
    cSQL = cSQL & " set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where i.rkz = 'J' "
    cSQL = cSQL & "  and Artlief.linr = " & lLinr
    gdBase.Execute cSQL, dbFailOnError
        
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(10)...", Label1(4)
    
    cSQL = "Update STADAPROSTRECKE s inner join Artikel a  on s.artnr = a.artnr "
    cSQL = cSQL & " set s.farbnr = val(a.awm) "
    cSQL = cSQL & " , s.agn = a.agn "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(11)...", Label1(4)
    
    cSQL = "Update STADAPROSTRECKE s inner join AGNDBF a  on s.agn = a.agn "
    cSQL = cSQL & " set s.agtext = a.agtext "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(12)...", Label1(4)
    
    'Hier VK-Preisanpassungen
    'Protokoll f�llen
    cSQL = "Insert into STADAPROSTRECKE Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'KVK-Preisanpassung' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "

    cSQL = cSQL & ", KVKALT as VKPR_ALT "
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
    
'    cSQL = cSQL & ", LEKALT as LEK_ALT "
'    cSQL = cSQL & ", LEKNEU as LEK_NEW "
'    cSQL = cSQL & ", RKZ "
'    cSQL = cSQL & " , '" & DateValue(Now) & "' as EXDat "
'    cSQL = cSQL & ", AGTEXT varchar(30)"
'    cSQL = cSQL & ", AGN int "
    cSQL = cSQL & " from ImportpriREWE where mnotizen = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    cSQL = "Update Artikel a inner join ImportpriREWE i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.kvkpr1 = i.vkpr  "
    cSQL = cSQL & " where i.mnotizen = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz �bernommen(13)...", Label1(4)
    
    BringFarbeInsSpiel "STADAPROSTRECKE", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme_XXX_Delta_neu"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
    

End Sub
Private Sub Sicherheitsl�schen_mitLinr(sArtnr As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from artlief where artnr = " & sArtnr & " and Linr = " & sLinr
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sicherheitsl�schen_mitLinr"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Status_Ermitteln()
On Error GoTo LOKAL_ERROR

    Dim rs As Recordset
    Dim sSQL As String
    
    'neue Artikel
    sSQL = "Select count(*) as maxi from ImportpriREWE where awm = '98'"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(6).Caption = CInt(Label1(6).Caption) + Val(Trim(rs!maxi))
            If Val(Label1(6).Caption) > 0 Then
                Command5(6).Visible = True
                Command5(6).Enabled = False
            End If
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
    'EK Preis�nderungen
    sSQL = "Select count(*) as maxi from ImportpriREWE where LEKALT <> 0"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(8).Caption = CInt(Label1(8).Caption) + Val(Trim(rs!maxi))
        End If
        
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "status_ermitteln"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErmittlungReweDuplisPlusDel()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsArt       As Recordset
    Dim cEAN        As String
    Dim lcount      As Long
    
    lcount = 0
    
    loeschNEW "ImportDupli", gdApp
    
    sSQL = "select count(ean) as count ,ean into ImportDupli from Meister group by ean having count(ean) > 1"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where ean is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where trim(ean) = ''"
    gdApp.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdApp.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsrs.EOF Then
        
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            

            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
            End If

            sSQL = "Select * from Meister where ean = '" & cEAN & "'"
            Set rsArt = gdApp.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungReweDuplisPlusDel"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub feldcheckRewe()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMeister       As Recordset
    Dim rsIMPORTPRI     As Recordset
    Dim sSQL            As String

    Dim sBez            As String
    Dim sEAN            As String
    
    
    loeschNEW "ImportpriREWE", gdApp
    CreateTableT2 "IMPORTPRIREWE", gdApp
    
    Set rsIMPORTPRI = gdApp.OpenRecordset("ImportpriREWE")
    
    Set rsMeister = gdApp.OpenRecordset("Meister")
    If Not rsMeister.EOF Then
    
        
        rsMeister.MoveFirst
        Do While Not rsMeister.EOF
        
            rsIMPORTPRI.AddNew
            rsIMPORTPRI!artnr = rsMeister!artnr
            
            'Bezeichnung auf * und ' checken
            If Not IsNull(rsMeister!BEZEICH) Then
                sBez = Left(rsMeister!BEZEICH, 35)
                sBez = SwapStr(sBez, "*", " ")  'stern
                sBez = SwapStr(sBez, Chr(34), "Z") '"
                sBez = SwapStr(sBez, "'", " ")  'Hochkommata
                sBez = SwapStr(sBez, ",", ".")  'komma
                sBez = SwapStr(sBez, "�", "�")  '�
                sBez = SwapStr(sBez, "�", "�")  '
                sBez = SwapStr(sBez, "�", "�")  '
                sBez = SwapStr(sBez, "�", "�")
                sBez = SwapStr(sBez, "�", "�")
                sBez = SwapStr(sBez, "�", "�")
                sBez = SwapStr(sBez, "`", " ")  '
                sBez = SwapStr(sBez, "|", " ")  '
            Else
                sBez = ""
            End If
            If gbTagAkt = True Then
                rsIMPORTPRI!BEZEICH = UCase(sBez)
            Else
                rsIMPORTPRI!BEZEICH = sBez
            End If
            'Standardm��ig auf Ziffer 0
            
            If Not IsNull(rsMeister!AGN) Then
                rsIMPORTPRI!AGN = rsMeister!AGN
            Else
                rsIMPORTPRI!AGN = 0
            End If
            
            If Not IsNull(rsMeister!PGN) Then
                rsIMPORTPRI!PGN = rsMeister!PGN
            Else
                rsIMPORTPRI!PGN = 0
            End If
            
            If Not IsNull(rsMeister!RKZ) Then
                rsIMPORTPRI!RKZ = rsMeister!RKZ
            Else
                rsIMPORTPRI!RKZ = "N"
            End If
            
            rsIMPORTPRI!lekneu = Round(rsMeister!lekpr, 2)
            rsIMPORTPRI!vkpr = rsMeister!vkpr
            
            If rsMeister!MWST = "1" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf rsMeister!MWST = "2" Then
                rsIMPORTPRI!MWST = "E"
            ElseIf rsMeister!MWST = "" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf IsNull(rsMeister!MWST) Then
                rsIMPORTPRI!MWST = "V"
            Else
                rsIMPORTPRI!MWST = rsMeister!MWST
            End If
            
            rsIMPORTPRI!linr = rsMeister!linr
            rsIMPORTPRI!LIBESNR = rsMeister!LIBESNR
            
            If Not IsNull(rsMeister!EAN) Then
                sEAN = rsMeister!EAN
            Else
                sEAN = "0"
            End If
            
            sEAN = SwapStr(sEAN, "00000", "")
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN = "0"
            Else
                rsIMPORTPRI!EAN = sEAN
            End If

            rsIMPORTPRI!EAN2 = rsMeister!EAN2
            rsIMPORTPRI!EAN3 = rsMeister!EAN3
            rsIMPORTPRI!ETIMERK = rsMeister!ETIMERK
            rsIMPORTPRI!MOPREIS = rsMeister!MOPREIS
            
            'standardm��ig auf "N"

            rsIMPORTPRI!NOTIZEN = rsMeister!NOTIZEN
            rsIMPORTPRI!BESTAND = rsMeister!BESTAND
            rsIMPORTPRI!VKMENGE = rsMeister!VKMENGE
            rsIMPORTPRI!VKDATUM = rsMeister!VKDATUM
            rsIMPORTPRI!MINMEN = rsMeister!MINMEN
            rsIMPORTPRI!INHALT = rsMeister!INHALT
            rsIMPORTPRI!INHALTBEZ = Trim(UCase(rsMeister!INHALTBEZ))
            rsIMPORTPRI!GRUNDPREIS = rsMeister!GRUNDPREIS
            rsIMPORTPRI!MINBEST = rsMeister!MINBEST
            rsIMPORTPRI!RABATT_OK = rsMeister!RABATT_OK
            rsIMPORTPRI!GEFUEHRT = rsMeister!GEFUEHRT
            rsIMPORTPRI!KVKPR1 = rsMeister!KVKPR1
            rsIMPORTPRI!ekpr = rsMeister!ekpr
            rsIMPORTPRI!PREISSCHU = rsMeister!PREISSCHU
            rsIMPORTPRI!BONUS_OK = rsMeister!BONUS_OK
            rsIMPORTPRI!UMS_OK = rsMeister!UMS_OK
            rsIMPORTPRI!AWM = rsMeister!AWM
            rsIMPORTPRI!LASTDATE = rsMeister!LASTDATE
            rsIMPORTPRI!LASTTIME = rsMeister!LASTTIME
            rsIMPORTPRI!MNOTIZEN = rsMeister!Status
            rsIMPORTPRI!KVKalt = 0
            rsIMPORTPRI!KVKNEU = rsMeister!vkpr
            
            If Not IsNull(rsMeister!AUFDAT) Then
                rsIMPORTPRI!AUFDAT = DateValue(Right(rsMeister!AUFDAT, 2) & "." & Mid(rsMeister!AUFDAT, 5, 2) & "." & Left(rsMeister!AUFDAT, 4))
            Else
                rsIMPORTPRI!AUFDAT = Null
            End If
            
            If Not IsNull(rsMeister!EXDAT) Then
                rsIMPORTPRI!EXDAT = DateValue(Right(rsMeister!EXDAT, 2) & "." & Mid(rsMeister!EXDAT, 5, 2) & "." & Left(rsMeister!EXDAT, 4))
            Else
                rsIMPORTPRI!EXDAT = Null
            End If
            
            rsIMPORTPRI!LPZ = rsMeister!LPZ
            rsIMPORTPRI!FARBNR = rsMeister!FARBNR
            rsIMPORTPRI!MARKE = rsMeister!MARKE
            rsIMPORTPRI!GROESSE = rsMeister!GROESSE
            rsIMPORTPRI!SPANNE = rsMeister!SPANNE
            rsIMPORTPRI!AUFSCHLAG = rsMeister!AUFSCHLAG
            rsIMPORTPRI!SYNStatus = rsMeister!SYNStatus
           
            rsIMPORTPRI.Update
            rsMeister.MoveNext
        Loop
    End If
    rsMeister.Close
    rsIMPORTPRI.Close
    
    sSQL = "Delete from ImportpriREWE where EAN is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriREWE where bezeich is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriREWE where bezeich = ''"
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "feldcheckRewe"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Function FormatiereBildschirmdatenXXX() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim lreservArtnr    As Long
    Dim lvergebeArtnr   As Long
   
    FormatiereBildschirmdatenXXX = False
    
    anzeige "normal", "Neue Artikel werden ermittelt...", Label1(4)
    'Farbe alle auf neu
    sSQL = "Update ImportpriREWE set AWM = '98' "
'    sSQL = sSQL & " ,AGN = " & Val(txtAGN.Text) & " "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(1)......", Label1(4)
    
    sSQL = "Create Index AWM on ImportpriREWE (AWM)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(2).........", Label1(4)
    
    sSQL = "Create Index ean on ImportpriREWE (ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(3)............", Label1(4)
    
    sSQL = "Create Index linr on ImportpriREWE (linr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(4)...", Label1(4)
    
    sSQL = "Create Index libesnr on ImportpriREWE (libesnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(5)......", Label1(4)
    
    'Artikel mit EAN �bereinstimmung auf standard
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    
    
    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    sSQL = "Update ImportpriREWE i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.LIBESNR = i.LIBESNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.artnr = ARTLIEF.artnr  "
    sSQL = sSQL & " where i.artnr is null "
    sSQL = sSQL & " and i.LIBESNR <> '0000000' "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Neue Artikel werden ermittelt(9)......", Label1(4)

    sSQL = "Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
'    sSQL = sSQL & " ,i.agn = Artikel.agn  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  = '0' and i.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(10).........", Label1(4)
    
    'Lekpreisver�nderungen anzeigen
    
    sSQL = "Update ImportpriREWE i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.artnr = i.artnr "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.LEKALT = ARTLIEF.LEKPR "
    sSQL = sSQL & " where not i.artnr is null "
    sSQL = sSQL & " and i.rkz  = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LEKALT from ImportpriREWE where not LEKALT is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!LEKALT = Round(rsrs!LEKALT, 2)
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update ImportpriREWE set LEKALT = 0 "
    sSQL = sSQL & " where not artnr is null "
    sSQL = sSQL & " and rkz  = 'N' "
    sSQL = sSQL & " and Round(lekneu,2) = Round(LEKALT,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Lekpreisver�nderungen anzeigen
    
    anzeige "normal", "Neue Artikel werden ermittelt(12)...", Label1(4)
    
    sSQL = " Update ImportpriREWE i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.PREISSCHU = ARTIKEL.PREISSCHU "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(13)......", Label1(4)
    
    Status_Ermitteln
    
    'Zieh doch mal  den freien Artikelnummernkreis hoch
    anzeige "normal", "F�r neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)
    
    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)
    
    sSQL = "Select * from ImportpriREWE where awm = '98' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!artnr = lreservArtnr
            rsrs.Update
            
            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verf�gung (Einstellungen �berpr�fen).", Label1(4)
                Exit Function
            Else
                lreservArtnr = lvergebeArtnr
                
            End If
            rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lvergebeArtnr > 0 Then
        sSQL = "Update FFE set ARTNRV = " & lvergebeArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    FormatiereBildschirmdatenXXX = True
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereBildschirmdatenXXX"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function
Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
 
    Dim sdatname    As String
    Dim i           As Integer
    Dim sSQL        As String
    Dim lLfnr       As Long
    Dim cLfnr       As String
    Dim rsrs        As DAO.Recordset
    Dim iRet        As Integer
    Dim ctmp        As String
 
    Select Case Index
    
        Case 0
            If Command5(0).Caption = "Schlie�en" Then
                Unload frmWKL186
            ElseIf Command5(0).Caption = "Weitere" Then
            
                Command5(0).Caption = "Schlie�en"
                Command5(5).Enabled = True
                Command5(1).Enabled = True
                
                Command5(1).BackColorTo = vbRed
                Command5(1).BackColorFrom = vbWhite
                
                Ablaufprotokoll_f�llen
                
                Command5(0).BackColorTo = glButtonHintergrund_to
                Command5(0).BackColorFrom = glButtonHintergrund_from
            
            End If
        Case 1      'xxx Stammdaten einlesen
        
            Timer1.Enabled = False
            
            Command5(1).Enabled = False
            
            Command5(1).BackColorTo = glButtonHintergrund_to
            Command5(1).BackColorFrom = glButtonHintergrund_from
                
            Command5(2).Enabled = False
        
            'Ablaufprotokoll f�llen
            'Etiketten erstellen
            'dem Anwender ein �bernahmeergebnis zeigen
            
            If XXXDatenEinlesen_neu(gsKinPfad, lblDatei.Caption) = True Then

                Kill gsKinPfad & "\" & lblDatei.Caption
            End If
            
            
            sSQL = "Delete from SORDER where datname = '" & lblDatei.Caption & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            If Datendrin("SORDER", gdBase) Then
                Command5(0).Caption = "Weitere"
                
                Command5(0).BackColorTo = vbRed
                Command5(0).BackColorFrom = vbWhite
                
            Else
                Command5(0).Caption = "Schlie�en"
            End If
            
        Case 2
            anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
            reportbildschirm "", "aWKL186"
            
            
        Case 3
            'EX Artikel als Etiketten
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", KVKNEU as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from ImportpriREWE_Dauer "
            sSQL = sSQL & " Where rkz = 'J'"
            gdBase.Execute sSQL, dbFailOnError

            gbEtiExArtikel = True

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
            
            gbEtiExArtikel = False
            
        Case 4
            zeigeHilfeDabapfad "LPROTOK", "StreckeProtokoll.txt"
            
        Case 5
        
            'L�schen �berspringen
            
            
            Timer1.Enabled = False
            
            Command5(5).Enabled = False
            Command5(1).Enabled = False
            
            Command5(1).BackColorTo = glButtonHintergrund_to
            Command5(1).BackColorFrom = glButtonHintergrund_from
                
            Command5(2).Enabled = False
        
            'Ablaufprotokoll f�llen
            'Etiketten erstellen
            'dem Anwender ein �bernahmeergebnis zeigen
            
            ctmp = "M�chten Sie wirklich diese aktuelle Stammdatendatei" & vbCrLf & vbCrLf
            ctmp = ctmp & lblDatei.Caption & vbCrLf & vbCrLf
            ctmp = ctmp & "l�schen?"
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
                'l�schen der Datei
                Kill gsKinPfad & "\" & lblDatei.Caption
            End If
            
            sSQL = "Delete from SORDER where datname = '" & lblDatei.Caption & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            If Datendrin("SORDER", gdBase) Then
                Command5(0).Caption = "Weitere"
                
                Command5(0).BackColorTo = vbRed
                Command5(0).BackColorFrom = vbWhite
                
            Else
                Command5(0).Caption = "Schlie�en"
            End If
            
            
            
            
            
            
            
            
            
            
        Case 6
        
            'neue Artikel als Etiketten
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", KVKNEU as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from ImportpriREWE_Dauer "
            sSQL = sSQL & " where awm = '98' "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
        Case Is = 9     'F2 Linie
            Screen.MousePointer = 0
            txtAGN_KeyUp vbKeyF2, 0
    
    End Select
    
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    lesenEinstellungen
    
    Ablaufprotokoll_f�llen
    
    iSec = 0
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Ablaufprotokoll_f�llen()
    On Error GoTo LOKAL_ERROR
    
    Dim sdatname        As String
    Dim i               As Integer
    Dim sSQL            As String
    Dim lLfnr           As Long
    Dim cLfnr           As String
    Dim rsrs            As DAO.Recordset
    Dim sUeberschrift   As String
    Dim sDateidatum     As String
    Dim lHeute          As Long
    
    
    lblUeberschrift.Caption = "bitte warten..."
    lblUeberschrift.Refresh
    frmWKL186.Caption = "bitte warten..."
    
    anzeige "normal", "bitte warten...", Label1(0)
    
    lHeute = Fix(Now)
    
'    lblUeberschrift.Caption = ""
'    frmWKL186.Caption = ""
'    anzeige "normal", "", Label1(0)
    
    loeschNEW "SORDER", gdBase
    CreateTableT2 "SORDER", gdBase
            
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "Strecke*.csv"
    File1.Refresh
    
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sDateidatum = Right(File1.list(i), 14)
            sDateidatum = Left(sDateidatum, 10)
            sDateidatum = SwapStr(sDateidatum, "-", ".")
            
            
            sdatname = File1.list(i)
            cLfnr = Mid(sdatname, 10, 3)
            lLfnr = Val(cLfnr)
            
            
            If IsDate(sDateidatum) = True Then
                If lHeute >= CLng(DateValue(sDateidatum)) Then
                    sSQL = "Insert into SORDER (lfnr,DATNAME)"
                    sSQL = sSQL & " Values ( "
                    sSQL = sSQL & " " & lLfnr & " "
                    sSQL = sSQL & ", '" & sdatname & "' "
                    sSQL = sSQL & " ) "
                    gdBase.Execute sSQL, dbFailOnError
                End If
            Else
                sSQL = "Insert into SORDER (lfnr,DATNAME)"
                sSQL = sSQL & " Values ( "
                sSQL = sSQL & " " & lLfnr & " "
                sSQL = sSQL & ", '" & sdatname & "' "
                sSQL = sSQL & " ) "
                gdBase.Execute sSQL, dbFailOnError
            End If
            
        Next i
    
'''        'Datei/en stehen an
'''        For i = 0 To File1.ListCount - 1
'''            sdatname = File1.list(i)
'''            cLfnr = Mid(sdatname, 10, 3)
'''            lLfnr = Val(cLfnr)
'''
'''            sSQL = "Insert into SORDER (lfnr,DATNAME)"
'''            sSQL = sSQL & " Values ( "
'''            sSQL = sSQL & " " & lLfnr & " "
'''            sSQL = sSQL & ", '" & sdatname & "' "
'''            sSQL = sSQL & " ) "
'''            gdBase.Execute sSQL, dbFailOnError
'''        Next i
    End If
    
    lblUeberschrift.Caption = "bitte warten... ..."
    lblUeberschrift.Refresh
    frmWKL186.Caption = "bitte warten... ..."
    anzeige "normal", "bitte warten... ...", Label1(0)
    
    sSQL = "Select Top 1 datname from SORDER order by lfnr asc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!Datname) Then
            lblDatei.Caption = rsrs!Datname
            sUeberschrift = Mid(lblDatei.Caption, 9, Len(lblDatei.Caption) - 12)
        End If
            
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    lblUeberschrift.Caption = "neue " & sUeberschrift & " - Artikeldaten"
    lblUeberschrift.Refresh
    frmWKL186.Caption = "neue " & sUeberschrift & " - Artikeldaten"
    anzeige "erfolg", "neue " & sUeberschrift & " - Artikeldaten stehen bereit.", Label1(0)
    
    Command5(1).BackColorTo = vbRed
    Command5(1).BackColorFrom = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Ablaufprotokoll_f�llen"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
   
End Sub
Private Sub lesenEinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    glartv = 600000
    glartb = 700000
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!ARTNRV) Then
                glartv = rsrs!ARTNRV
            Else
                glartv = 600000
            End If
            
            If Not IsNull(rsrs!ARTNRB) Then
                glartb = rsrs!ARTNRB
            Else
                glartb = 700000
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesenEinstellungen"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "REWE_AGN", gdApp
    loeschNEW "MEISTER", gdApp
'    loeschNEW "STADAPROSTRECKE", gdBase
    loeschNEW "ImportDupli", gdApp
    loeschNEW "ImportpriREWE", gdBase
    loeschNEW "ImportpriREWE_dauer", gdBase
    loeschNEW "ImportpriREWE", gdApp
    loeschNEW "SORDER", gdBase
    
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
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Timer1.Enabled = False

If Index = 11 Then
    URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/231-rewe-stammdaten-einlesen.html"
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    iSec = iSec + 1
    
    If iSec >= 10 Then
        Unload frmWKL186
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub txtAGN_Change()
On Error GoTo LOKAL_ERROR
    
    If Len(txtAGN.Text) >= 3 Then
        Label10.Caption = Ermittleagntext(txtAGN.Text)
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAGN_Change"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAGN_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    cZeichen = Chr$(KeyAscii)
    cValid = "1234567890" & Chr$(8)
       
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
    Fehler.gsFunktion = "txtAGN_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAGN_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "AGN"
        frmWK00a.Show 1
        If gF2Prompt.cWahl <> "" Then
            txtAGN.Text = gF2Prompt.cWahl
        End If
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAGN_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil neue XXX Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
