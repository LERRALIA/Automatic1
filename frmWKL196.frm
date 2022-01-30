VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL196 
   BackColor       =   &H00C0C000&
   Caption         =   "neue Budni - Artikeldaten"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWKL196.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   10560
      MaxLength       =   5
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox chk_KVK 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "UVP ¸bernehmen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9600
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "nur relevante Artikel, die schon einmal verkauft wurden"
      Height          =   495
      Left            =   8400
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CheckBox chk_vk 
      Caption         =   "nur relevante Artikel, die schon einmal verkauft wurden"
      Height          =   495
      Left            =   8400
      TabIndex        =   22
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11160
      Top             =   240
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   6
      Left            =   6840
      TabIndex        =   21
      Top             =   3720
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
      Index           =   5
      Left            =   6840
      TabIndex        =   20
      Top             =   5160
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
      TabIndex        =   19
      Top             =   4200
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
      Left            =   10320
      Pattern         =   "MASTER!.*"
      TabIndex        =   6
      Top             =   1200
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
      Caption         =   "Schlieﬂen"
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
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   25
      Top             =   5640
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
   Begin VB.Label Label1 
      Caption         =   "Aufschlag in % auf den Listeneinkaufspreis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   9600
      TabIndex        =   30
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Unverbindliche Preisempfehlung als Kassenverkaufspreis ¸bernehmen? "
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
      Index           =   16
      Left            =   840
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "KVK Preis‰nderung"
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
      Index           =   15
      Left            =   1440
      TabIndex        =   27
      Top             =   5640
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
      Index           =   14
      Left            =   4080
      TabIndex        =   26
      Top             =   5640
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
      Index           =   13
      Left            =   4080
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "Umleitung"
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
      Index           =   12
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "n‰here Informationen hier: (bitte anklicken)"
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
      Left            =   4680
      MouseIcon       =   "frmWKL196.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   15
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3735
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
      Top             =   3240
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
      Top             =   3240
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
      Top             =   4680
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
      Top             =   4200
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
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "EK Preis‰nderung"
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
      Top             =   4680
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
      Top             =   4200
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
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Mˆchten Sie diese ¸bernehmen, so klicken Sie auf ""Einlesen""."
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
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neue Budni-Artikeldaten stehen bereit. "
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "neue Budni - Artikeldaten"
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
Attribute VB_Name = "frmWKL196"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim glartv      As Long
Dim glartb      As Long
Dim iSec        As Integer
Private Function checkBUDNIinLISRT() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsLi As Recordset
    
    checkBUDNIinLISRT = 0

    sSQL = "Select linr from BUDNIE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!linr) Then
            checkBUDNIinLISRT = rsrs!linr
            
            sSQL = "Select * from LISRT where LINR = " & checkBUDNIinLISRT
            sSQL = sSQL & " and ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
            Set rsLi = gdBase.OpenRecordset(sSQL)
            If rsLi.RecordCount = 0 Then
                checkBUDNIinLISRT = 0
            End If
            rsLi.Close
        
        End If
    End If
    
    If checkBUDNIinLISRT = 0 Then
    
        Screen.MousePointer = 0
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
        If gF2Prompt.cFeld <> "" Then
            gsAnzeige00a = "Bitte den Budni - Lieferant ausw‰hlen!"
            frmWK00a.Show 1
        End If
        gsAnzeige00a = ""
        
        anzeige "normal", "Der Lieferant: " & gF2Prompt.cWahl & " wurde zugeordnet.", Label1(4)
        
        If gF2Prompt.cWahl <> "" Then
             checkBUDNIinLISRT = CDbl(gF2Prompt.cWahl)
        End If
        
        If checkBUDNIinLISRT <> 0 Then
            sSQL = "update BUDNIE set linr = " & checkBUDNIinLISRT
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkBUDNIinLISRT"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen: Fremdformate ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub BudniDatenEinlesen(sPfad As String, sDatei As String, bKomplett As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    loeschNEW "REWE_AGN", gdApp
    CreateTableT2 "REWE_AGN", gdApp

    'vorbereitung der Importtabelle
    '1. erst lˆschen

    loeschNEW "MEISTER", gdApp
    CreateTable "MEISTER", gdApp

    Dim iFileNr As Integer
    Dim cSatz1 As String
    Dim lPos As Long
    Dim cEinzelsatz As String
    Dim lLinr As Long
    Dim lfnr1 As Long
    Dim rsagn As Recordset
    Dim lPosEnde As Long
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lposSemiEnde As Long
    Dim cWert As String
    Dim cAgn As String
    Dim cAGNBEZEICH As String
    Dim dBudniFaktor As Double
    
    dBudniFaktor = 0
    
    If Text1(2).Text <> "" Then
        If IsNumeric(Text1(2).Text) = True Then
            dBudniFaktor = CDbl(Text1(2).Text)
        End If
    End If
    
    Label1(2).Visible = True
    Label1(3).Visible = True
    Label1(5).Visible = True
    Label1(9).Visible = True
    Label1(6).Visible = True
    Label1(7).Visible = True
    Label1(8).Visible = True
    Label1(10).Visible = True
    Label1(12).Visible = True
    Label1(13).Visible = True

    Command5(3).Enabled = False
    Command5(5).Enabled = False
    Command5(6).Enabled = False
    
    lfnr1 = 0
    lPos = 1
    lPosEnde = 1
    lposSemiEnde = 1
    
    lLinr = checkBUDNIinLISRT()
    
    If lLinr = 0 Then
        Screen.MousePointer = 0
        MsgBox "Keine auswertbare Lieferantennummer zugewiesen.", vbInformation, "Winkiss Hinweis:"

        Exit Sub
    End If
    
    If bKomplett = True Then
        sSQL = "Update Artikel inner join Artlief on artikel.artnr = Artlief.artnr "
        sSQL = sSQL & " set Artikel.RKZ = 'J' where Artlief.linr = " & lLinr & " "
        gdBase.Execute sSQL, dbFailOnError
    End If

    Set rsrs = gdApp.OpenRecordset("MEISTER")
    Set rsagn = gdApp.OpenRecordset("REWE_AGN")

    iFileNr = FreeFile
    Open sPfad & "\" & sDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then

        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1


        lLenfil = Len(cSatz1)

        lPosEnde = InStr(lPos, cSatz1, vbCrLf)
        lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile ¸berspringen

        Do
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)

            cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
            lPos = lPos + lPosEnde - lPos + 2

            lposSemi = 1

            rsrs.AddNew
            rsagn.AddNew
            lfnr1 = lfnr1 + 1
            rsrs!lfnr = lfnr1

            'Libesnr
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            If Val(cWert) = 0 Then
                rsrs!LIBESNR = "0000000"
            Else
                rsrs!LIBESNR = cWert
            End If

            'Artnr von Ernst ohne Bedeutung
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            'Bezeich
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!BEZEICH = cWert
            
            'EAN
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If cWert <> "" Then
                cWert = Val(cWert)
                If Len(cWert) = 11 Then
                    cWert = "0" & cWert
                End If
            End If
            rsrs!EAN = cWert
            
            'EAN2
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If cWert <> "" Then
                cWert = Val(cWert)
                If Len(cWert) = 11 Then
                    cWert = "0" & cWert
                End If
            End If
            rsrs!EAN2 = cWert
            
            'EAN3
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If cWert <> "" Then
                cWert = Val(cWert)
                If Len(cWert) = 11 Then
                    cWert = "0" & cWert
                End If
            End If
            rsrs!EAN3 = cWert
            
            
            'EK = EKPR
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            'hier wird Aufgeschlagen
            
            
            
            If IsNumeric(cWert) Then
            
                If dBudniFaktor > 0 Then
                    
                    rsrs!lekpr = CDbl(cWert) + (dBudniFaktor * CDbl(cWert) / 100)
                    
                Else
                    rsrs!lekpr = cWert
                End If
                
                
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
            
            'MM = VPE
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!MINMEN = cWert


            'MWST
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            If cWert <> "" Then
                If cWert = "V" Then
                    rsrs!MWST = "V"
                ElseIf cWert = "E" Then
                    rsrs!MWST = "E"
                ElseIf cWert = "O" Then
                    rsrs!MWST = "O"
                Else
                    rsrs!MWST = "V"
                End If
            Else
                rsrs!MWST = "V"
            End If
            
            'Inhalt
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
        
            rsrs!INHALT = Val(cWert)
                

            'Inhaltbez
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
        
            rsrs!INHALTBEZ = cWert
            
            'GP = Grundpreisauszeichnungspflicht
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
        
            If cWert = "True" Then
                rsrs!GRUNDPREIS = "J"
            Else
                rsrs!GRUNDPREIS = "N"
            End If
            
            'Notizen
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
        
            rsrs!NOTIZEN = Left(cWert, 25)
            

            'AGN
            cWert = ""
            cAgn = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            cAgn = cWert
            If cAgn = "" Then cAgn = "0"
            rsrs!AGN = CLng(cAgn)
            rsagn!RAGN = CLng(cAgn)

            'AGN_BEZ
            cWert = ""
            cAGNBEZEICH = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            cAGNBEZEICH = Left(cWert, 30)
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "Ω", "ƒ")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "!", "‹")
            cAGNBEZEICH = SwapStr(cAGNBEZEICH, "\", "÷")
            rsagn!RAGTEXT = cAGNBEZEICH
            
            'Marke
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!MARKE = cWert
            
            'Auslistung EX
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            If UCase(cWert) = "TRUE" Then
                rsrs!RKZ = "J"
                Label1(7).Caption = CInt(Label1(7).Caption) + 1
                Label1(7).Refresh
            Else
                rsrs!RKZ = "N"
            End If
            
            'PGN
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!PGN = Val(cWert)
            

            'PGN_BEZ
            cWert = ""
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

            
            
            
            rsrs!LPZ = 1
            rsrs!linr = lLinr
            rsrs!GEFUEHRT = "J"


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

    '****************
    
    'Umleitung rausziehen
    loeschNEW "Budni_Umleitung", gdApp

    sSQL = "Select libesnr, Bezeich, NOTIZEN as UML_libesnr into Budni_Umleitung from Meister where NOTIZEN <> '' "
    gdApp.Execute sSQL, dbFailOnError

    
    
    
    'hier Reweagn's auf neue pr¸fen
    
    

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

    'Ende hier Reweagn's auf neue pr¸fen
    
    anzeige "normal", "EAN Duplikate werden entfernt...", Label1(4)

    '6. EAN - Duplikats¸berpr¸fung in der Importtabelle Anzahl ermitteln
    ErmittlungBUDNIDuplisPlusDel
        
    sSQL = "Delete from Meister where LEKPR = 0 "
    gdApp.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Die Artikeldaten werden ¸berpr¸ft...", Label1(4)
   
    '7. diverse Feld¸berpr¸fungen vornehmen
    feldcheckBUDNI
    anzeige "normal", "Die Sortiments¸bersicht wird erstellt...", Label1(4)

    '8. Tabelle IMPORTPRI zur Datenbank kopieren
    loeschNEW "ImportpriBUDNI", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "ImportpriBUDNI"
    
    '8. Tabelle BUDNI_UMLEITUNG zur Datenbank kopieren
    loeschNEW "BUDNI_UMLEITUNG", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "BUDNI_UMLEITUNG"

    sSQL = "Create Index ARTNR on ImportpriBUDNI (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update ImportpriBUDNI set marke ='' "
    gdBase.Execute sSQL, dbFailOnError
    
    FormatiereBildschirmdatenBUDNI
    
    'jetzt ¸bernehmen
    Uebernahme_BUDNI_Delta sDatei, lLinr
    
    If NewTableSuchenDBKombi("ImportpriBUDNI_dauer", gdBase) = False Then
        sSQL = "select * into ImportpriBUDNI_dauer from ImportpriBUDNI  "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Insert into ImportpriBUDNI_dauer select * from ImportpriBUDNI  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Kill sPfad & "\" & sDatei
    
    Command5(2).Enabled = True
    
    Command5(3).Enabled = True
    Command5(5).Enabled = True
    Command5(6).Enabled = True
    
'    Command5(4).Enabled = True
    
    
    anzeige "normal", "Fertig! Artikel¸bernahme beendet, Protokoll und Etiketten ausdrucken!", Label1(4)
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BudniDatenEinlesen"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
Resume Next
    
End Sub
Private Sub Uebernahme_BUDNI_Delta(sDatei As String, lLieferant As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim rsArtlief       As DAO.Recordset
    Dim sArtnr          As String

    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen...", Label1(4)
    
    If NewTableSuchenDBKombi("STADAPROBUDNI", gdBase) = False Then
        CreateTableT2 "STADAPROBUDNI", gdBase
    End If
    
    'Alle EANS auffangen************************************************

    loeschNEW "artean_BU", gdBase

    cSQL = "Create Table artean_BU  "
    cSQL = cSQL & "( ARTNR int"
    cSQL = cSQL & ", EANCH varchar(13)"
    cSQL = cSQL & ", OTHERARTNR1 int"
    cSQL = cSQL & ", OTHERARTNR1_Bestand int"
    
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Insert into artean_BU Select artnr, val(ean) as EANCH from ImportpriBUDNI  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_BU Select artnr, val(ean2) as EANCH from ImportpriBUDNI  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_BU Select artnr, val(ean3) as EANCH from ImportpriBUDNI  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU set EANCH = '0' & EANCH  "
    cSQL = cSQL & " where len(EANCH)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU set OTHERARTNR1 = 0 , OTHERARTNR1_Bestand = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where EANCH = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    loeschNEW "artean_Artikel", gdBase

    cSQL = "Create Table artean_Artikel  "
    cSQL = cSQL & "( ARTNR int"
    cSQL = cSQL & ", EANCH varchar(13)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN = '' where EAN is null  "
    gdBase.Execute cSQL, dbFailOnError


    cSQL = "Insert into artean_Artikel Select artnr, val(ean) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN2 = '' where EAN2 is null  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_Artikel Select artnr, val(ean2) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN3 = '' where EAN3 is null  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into artean_Artikel Select artnr, val(ean3) as EANCH from Artikel  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
        'nich vergessen ARTEAN_K
        cSQL = "Update ARTEAN_K set EAN = '' where EAN is null  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Insert into artean_Artikel Select artnr, val(ean) as EANCH from ARTEAN_K  "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    cSQL = "Delete from artean_Artikel where EANCH = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update artean_Artikel set EANCH = '0' & EANCH  "
    cSQL = cSQL & " where len(EANCH)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'mit dieser Information holen wir uns aus der Datenbank alle ARTNR(ungleich), die schon solch eine EAN besitzen
    
    cSQL = "Update artean_BU inner join artean_Artikel on artean_BU.EANCH = artean_Artikel.EANCH "
    cSQL = cSQL & " set artean_BU.OTHERARTNR1 = artean_Artikel.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where Artnr = OTHERARTNR1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_BU where OTHERARTNR1 = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_BU inner join Artikel on artean_BU.OTHERARTNR1 = Artikel.ARTNR "
    cSQL = cSQL & " set artean_BU.OTHERARTNR1_Bestand = Artikel.Bestand "
    gdBase.Execute cSQL, dbFailOnError
    
    
    'nun zur Vollendung
    'bei Otherartnr1 die EAN rausnehmen und von Otherartnr1 auf Artnr den Bestand addieren
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN2 = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean2 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.OTHERARTNR1 and Artikel.EAN3 = artean_BU.EANCH "
    cSQL = cSQL & " set Artikel.ean3 = '' "
    gdBase.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
        cSQL = "Update ARTEAN_K inner join artean_BU on ARTEAN_K.ARTNR = artean_BU.OTHERARTNR1 and ARTEAN_K.EAN = artean_BU.EANCH "
        cSQL = cSQL & " set ARTEAN_K.ean = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from ARTEAN_K where ean = '' "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    'bestand addieren
    
    cSQL = "Update Artikel inner join artean_BU on Artikel.ARTNR = artean_BU.ARTNR  "
    cSQL = cSQL & " set Artikel.Bestand = Artikel.Bestand + artean_BU.OTHERARTNR1_Bestand"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    

    'Ende ********************** Alle EANS auffangen
    
    
    
    
    
    
    'Ex Artikel, und Exdatum
    cSQL = "Update Artlief set Artlief.RKZ = 'J'"
    cSQL = cSQL & ", Artlief.EXDAT = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", Artlief.SYNSTATUS = 'E' "
    cSQL = cSQL & " where  Artlief.linr = " & lLieferant
    cSQL = cSQL & " and Artlief.RKZ = 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(1)...", Label1(4)
    
    'neue Artikel
    cSQL = "Select * from ImportpriBUDNI where awm = '98'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!artnr) Then
                sArtnr = Trim(rsrs!artnr)
            End If
            
            Sicherheitslˆschen sArtnr 'artlief

            rsrs.MoveNext
            
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(3)...", Label1(4)
    
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
    cSQL = cSQL & " from ImportpriBUDNI where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(4)...", Label1(4)
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'A' as SYNSTATUS "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & " from ImportpriBUDNI where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(5)...", Label1(4)
    
    
    'Neuheiten
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROBUDNI Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Neuheiten/wieder verf¸gbar' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
    cSQL = cSQL & " from ImportpriBUDNI where awm = '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(6)...", Label1(4)
    
    'EK-Preis‰nderungen
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROBUDNI Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'EK-Preis‰nderungen' as AKT "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", LEKALT as LEK_ALT "
    cSQL = cSQL & ", LEKNEU as LEK_NEW "
    cSQL = cSQL & " from ImportpriBUDNI where LEKALT <> 0 and RKZ = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    
'''    'KVK-Preis‰nderungen
'''    'Protokoll f¸llen
'''    cSQL = "Insert into STADAPROBUDNI Select "
'''    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
'''    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
'''    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
'''    cSQL = cSQL & ", 'KVK-Preis‰nderungen' as AKT "
'''    cSQL = cSQL & ", ARTNR "
'''    cSQL = cSQL & ", BEZEICH "
'''    cSQL = cSQL & ", EAN "
'''    cSQL = cSQL & ", libesnr "
'''    cSQL = cSQL & ", KVKALT as VKPR_ALT "
'''    cSQL = cSQL & ", KVKNEU as VKPR_NEW "
'''    cSQL = cSQL & " from ImportpriBUDNI where RKZ = 'N' "
'''    cSQL = cSQL & " and Round(KVKALT, 2) <> Round(KVKNEU, 2)"
'''    gdBase.Execute cSQL, dbFailOnError
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(7)...", Label1(4)
    
    
    
    

    
    
    
    
    
    
    'alle anderen ƒnderungen + Artliefeintrag
    
    
    
    
    'test fang von hinten an
    '3.EAN
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.1)...", Label1(4)
    loeschNEW "REWEZWEAN", gdBase
    CreateTableT2 "REWEZWEAN", gdBase
    
    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN3) as EANTO from ImportpriBUDNI "
    cSQL = cSQL & " where (not ean3 is null or ean3 = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
    cSQL = cSQL & " where len(EANTO)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from REWEZWEAN  "
    cSQL = cSQL & " where EANTO = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EANTO  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EANTO <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    '2.EAN
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.2)...", Label1(4)
    loeschNEW "REWEZWEAN", gdBase
    CreateTableT2 "REWEZWEAN", gdBase
    
    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN2) as EANTO from ImportpriBUDNI "
    cSQL = cSQL & " where (not ean2 is null or ean2 = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
    cSQL = cSQL & " where len(EANTO)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from REWEZWEAN  "
    cSQL = cSQL & " where EANTO = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EANTO  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EANTO <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    '1.EAN
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.3)...", Label1(4)
    loeschNEW "REWEZWEAN", gdBase
    CreateTableT2 "REWEZWEAN", gdBase
    
    cSQL = "Insert into REWEZWEAN select ARTNR,val(EAN) as EANTO from ImportpriBUDNI "
    cSQL = cSQL & " where (not ean is null or ean = '')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update REWEZWEAN set EANTO = '0' & EANTO  "
    cSQL = cSQL & " where len(EANTO)= 11 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from REWEZWEAN  "
    cSQL = cSQL & " where EANTO = '0' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel a inner join REWEZWEAN i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.EAN = i.EANTO  "
    cSQL = cSQL & " , a.EAN2 = a.EAN "
    cSQL = cSQL & " , a.EAN3 = a.EAN2 "
    cSQL = cSQL & " where i.EANTO <> a.ean "
    cSQL = cSQL & " and not a.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(8.4)...", Label1(4)
    
    cSQL = "Update Artikel a inner join ImportpriBUDNI i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.BEZEICH = i.BEZEICH "
    cSQL = cSQL & " , a.LIBESNR = i.LIBESNR "
    cSQL = cSQL & " , a.LPZ = i.LPZ "
    cSQL = cSQL & " , a.RKZ = i.RKZ "
'    cSQL = cSQL & " , a.EAN = i.EAN "
    cSQL = cSQL & " , a.MWST = i.MWST "
    cSQL = cSQL & " , a.MINMEN = i.MINMEN "
    cSQL = cSQL & " , a.INHALT = i.INHALT "
    cSQL = cSQL & " , a.INHALTBEZ = i.INHALTBEZ "
    cSQL = cSQL & " , a.AGN = i.AGN "
    cSQL = cSQL & " , a.NOTIZEN = i.NOTIZEN "
    cSQL = cSQL & " , a.lekpr = i.lekneu"
    cSQL = cSQL & ", a.LASTDATE = '" & DateValue(Now) & "' "
    cSQL = cSQL & ", a.LASTTIME = '" & TimeValue(Now) & "' "
    cSQL = cSQL & ", a.SYNSTATUS = 'E' "
    
'    If chk_KVK.Value = vbChecked Then
'        cSQL = cSQL & " , a.KVKPR1 = i.KVKNEU "
'    End If
    
    
    
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(9)...", Label1(4)
    
    cSQL = "Delete from artlief where artnr in (Select artnr from ImportpriBUDNI) and Linr = " & lLieferant & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into ARTLIEF Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LEKNEU as LEKPR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", MINMEN "
    cSQL = cSQL & ", 0 as SPANNE "
    cSQL = cSQL & ", 'E' as SYNSTATUS "
    cSQL = cSQL & ", RKZ "
    cSQL = cSQL & ", NULL as EXDAT "
    cSQL = cSQL & " from ImportpriBUDNI " 'where awm <> '98' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN = '' where EAN = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriBUDNI) "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update Artikel set EAN2 = '' where EAN2 = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriBUDNI) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Artikel set EAN3 = '' where EAN3 = '0' "
    cSQL = cSQL & " and artnr in (Select artnr from ImportpriBUDNI) "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    

    cSQL = "update artikel set ean3 = '' where ean3 = ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update artikel set ean2 = '' where ean2 = ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update artikel set ean3 = '' where ean3 = ean2 "
    gdBase.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("artean_K", gdBase) Then

        If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = False Then
            cSQL = " Alter table artean_K add erkannt Text(1)  "
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        
        
        'gibt es den EAN aus Artean_K auch an EAN1 der Artikel
    
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean2 and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        cSQL = "Update artean_K set erkannt = 'N'  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean3 and artean_K.artnr =  ARTIKEL.artnr "
        cSQL = cSQL & " set artean_K.erkannt = 'J'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from artean_K where erkannt = 'J'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
        
    
'        cSQL = "Update artean_K set erkannt = 'N'  "
'        gdBase.Execute cSQL, dbFailOnError
'
'        cSQL = "Update artean_K inner join  ARTIKEL on artean_K.artnr =  ARTIKEL.artnr "
'        cSQL = cSQL & " set artean_K.erkannt = 'J'"
'        gdBase.Execute cSQL, dbFailOnError
'
'        cSQL = "Delete from artean_K where erkannt = 'N'  "
'        gdBase.Execute cSQL, dbFailOnError
        
        If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = True Then
            cSQL = " Alter table artean_K drop erkannt   "
            gdBase.Execute cSQL, dbFailOnError
        End If
    
    End If
    
    
    
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(10)...", Label1(4)
    
    cSQL = "Update STADAPROBUDNI s inner join Artikel a  on s.artnr = a.artnr "
    cSQL = cSQL & " set s.farbnr = val(a.awm) "
    cSQL = cSQL & " , s.agn = a.agn "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(11)...", Label1(4)
    
    cSQL = "Update STADAPROBUDNI s inner join AGNDBF a  on s.agn = a.agn "
    cSQL = cSQL & " set s.agtext = a.agtext "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(12)...", Label1(4)
    
    'Hier VK-Preisanpassungen
    cSQL = "Update Artikel a inner join ImportpriBUDNI i on a.artnr = i.artnr "
    cSQL = cSQL & " set a.kvkpr1 = i.vkpr  "
    cSQL = cSQL & " where i.mnotizen = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz ¸bernommen(13)...", Label1(4)
    
    
    'Ex Artikel
    'Protokoll f¸llen
    cSQL = "Insert into STADAPROBUDNI Select "
    cSQL = cSQL & " '" & DateValue(Now) & "' as Datum"
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as UHRZEIT "
    cSQL = cSQL & ", '" & sDatei & "' as Quelldat "
    cSQL = cSQL & ", 'Auslistung/Ex' as AKT "
    cSQL = cSQL & ", Artikel.ARTNR "
    cSQL = cSQL & ", Artikel.BEZEICH "
    cSQL = cSQL & ", Artikel.EAN "
    cSQL = cSQL & ", Artlief.libesnr "
    cSQL = cSQL & ", Artlief.RKZ "
    cSQL = cSQL & ", Artlief.EXDAT "
    cSQL = cSQL & " from Artlief inner join Artikel on "
    cSQL = cSQL & " Artlief.artnr = Artikel.artnr "
    cSQL = cSQL & " where Artlief.rkz = 'J' "
    cSQL = cSQL & " and Artlief.linr = " & lLieferant
    cSQL = cSQL & " and Artlief.EXDAT = " & CLng(DateValue(Now))
    gdBase.Execute cSQL, dbFailOnError
    
    Label1(7).Caption = ermittleAnzBudEx
    
    If Val(Label1(7).Caption) > 0 Then
        Command5(3).Visible = True
        Command5(3).Enabled = False
        
        chk_vk.Visible = True
    End If
    
    BringFarbeInsSpiel "STADAPROBUDNI", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme_BUDNI_Delta"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermittleAnzBudEx() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    ermittleAnzBudEx = 0
    
    cSQL = "Select * from STADAPROBUDNI where RKZ = 'J' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        ermittleAnzBudEx = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleAnzBudEx"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub Sicherheitslˆschen_mitLinr(sArtnr As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from artlief where artnr = " & sArtnr & " and Linr = " & sLinr
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sicherheitslˆschen_mitLinr"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Status_Ermitteln()
On Error GoTo LOKAL_ERROR

    Dim rs As Recordset
    Dim sSQL As String
    
    'neue Artikel
    sSQL = "Select count(*) as maxi from ImportpriBUDNI where awm = '98'"
    
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
    
    'EK Preis‰nderungen
    sSQL = "Select count(*) as maxi from ImportpriBUDNI where LEKALT <> 0"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!maxi) Then
            Label1(8).Caption = CInt(Label1(8).Caption) + Val(Trim(rs!maxi))
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
''    'KVK Preis‰nderungen
''    sSQL = "Select count(*) as maxi from ImportpriBUDNI where Round(KVKALT, 2) <> Round(KVKNEU, 2)"
''
''    Set rs = gdBase.OpenRecordset(sSQL)
''    If Not rs.EOF Then
''
''        If Not IsNull(rs!maxi) Then
''            Label1(14).Caption = CInt(Label1(14).Caption) + Val(Trim(rs!maxi))
''            If Val(Label1(14).Caption) > 0 Then
''                Command5(4).Visible = True
''                Command5(4).Enabled = False
''            End If
''        End If
''
''    End If
''    rs.Close: Set rs = Nothing
    
   
   'Umleitung
    sSQL = "Select count(*) as maxi from BUDNI_UMLEITUNG "
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            Label1(13).Caption = CInt(Label1(13).Caption) + Val(Trim(rs!maxi))
            If Val(Label1(13).Caption) > 0 Then
                Command5(5).Visible = True
                Command5(5).Enabled = False
                Check1.Visible = True
            End If
        End If
        
    End If
    rs.Close: Set rs = Nothing
   
'    'VK Preis‰nderungen
'    sSQL = "Select count(*) as maxi from ImportpriBUDNI where Round(KVKALT,2) <> Round(KVKNEU,2)"
'    sSQL = sSQL & " and mnotizen = 'J' "
'
'    Set rs = gdBase.OpenRecordset(sSQL)
'    If Not rs.EOF Then
'
'        If Not IsNull(rs!maxi) Then
'            Label1(13).Caption = CInt(Label1(13).Caption) + Val(Trim(rs!maxi))
'            If Val(Label1(13).Caption) > 0 Then
'                Command5(5).Visible = True
'                Command5(5).Enabled = False
'            End If
'        End If
'
'    End If
'    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "status_ermitteln"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErmittlungBUDNIDuplisPlusDel()
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
    Fehler.gsFunktion = "ErmittlungBUDNIDuplisPlusDel"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub feldcheckBUDNI()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMeister       As Recordset
    Dim rsIMPORTPRI     As Recordset
    Dim sSQL            As String

    Dim sBez            As String
    Dim sEAN            As String
    
    
    loeschNEW "ImportpriBUDNI", gdApp
    CreateTableT2 "IMPORTPRIBUDNI", gdApp
    
    Set rsIMPORTPRI = gdApp.OpenRecordset("ImportpriBUDNI")
    
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
                sBez = SwapStr(sBez, "·", "ﬂ")  'ﬂ
                sBez = SwapStr(sBez, "î", "ˆ")  '
                sBez = SwapStr(sBez, "Ñ", "‰")  '
                sBez = SwapStr(sBez, "Å", "¸")
                sBez = SwapStr(sBez, "ô", "÷")
                sBez = SwapStr(sBez, "ö", "‹")
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
            'Standardm‰ﬂig auf Ziffer 0
            
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
            
            'EAN
            If Not IsNull(rsMeister!EAN) Then
                sEAN = rsMeister!EAN
            Else
                sEAN = "0"
            End If
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN = "0"
            Else
                rsIMPORTPRI!EAN = sEAN
            End If
            
            'EAN2
            If Not IsNull(rsMeister!EAN2) Then
                sEAN = rsMeister!EAN2
            Else
                sEAN = "0"
            End If
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN2 = "0"
            Else
                rsIMPORTPRI!EAN2 = sEAN
            End If
            
            'EAN3
            If Not IsNull(rsMeister!EAN3) Then
                sEAN = rsMeister!EAN3
            Else
                sEAN = "0"
            End If
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN3 = "0"
            Else
                rsIMPORTPRI!EAN3 = sEAN
            End If

'            rsIMPORTPRI!EAN2 = rsMeister!EAN2
'            rsIMPORTPRI!EAN3 = rsMeister!EAN3
            rsIMPORTPRI!ETIMERK = rsMeister!ETIMERK
            rsIMPORTPRI!MOPREIS = rsMeister!MOPREIS
            
            'standardm‰ﬂig auf "N"

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
    
    sSQL = "Delete from ImportpriBUDNI where EAN is null "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriBUDNI where bezeich is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ImportpriBUDNI where bezeich = ''"
    gdApp.Execute sSQL, dbFailOnError
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "feldcheckBUDNI"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Function FormatiereBildschirmdatenBUDNI() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim lreservArtnr    As Long
    Dim lvergebeArtnr   As Long
   
    FormatiereBildschirmdatenBUDNI = False
    
    anzeige "normal", "Neue Artikel werden ermittelt...", Label1(4)
    'Farbe alle auf neu
    sSQL = "Update IMPORTPRIBUDNI set AWM = '98' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    anzeige "normal", "Neue Artikel werden ermittelt(1)......", Label1(4)
    
    sSQL = "Create Index AWM on IMPORTPRIBUDNI (AWM)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(2).........", Label1(4)
    
    sSQL = "Create Index ean on IMPORTPRIBUDNI (ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(3)............", Label1(4)
    
    sSQL = "Create Index linr on IMPORTPRIBUDNI (linr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(4)...", Label1(4)
    
    sSQL = "Create Index libesnr on IMPORTPRIBUDNI (libesnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(5)......", Label1(4)
    
    
    'Artikel mit EAN ¸bereinstimmung auf standard
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = i.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    
    
    'Teil 2
    
    'Artikel mit EAN ¸bereinstimmung auf standard
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = i.EAN2 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean2  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = i.EAN2 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean2  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = i.EAN2 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean2  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    
    
    
    'Teil 3
    
    'Artikel mit EAN ¸bereinstimmung auf standard
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = i.EAN3 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean3  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(6).........", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = i.EAN3 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean3  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(7)............", Label1(4)
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = i.EAN3 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.ean3  <> '0' "
    sSQL = sSQL & " and i.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(8)...", Label1(4)
    
    
    sSQL = "Update IMPORTPRIBUDNI inner join ARTEAN_K on "
    sSQL = sSQL & " ARTEAN_K.EAN = IMPORTPRIBUDNI.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRIBUDNI.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRIBUDNI.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where IMPORTPRIBUDNI.ean  <> '0' "
    sSQL = sSQL & " and IMPORTPRIBUDNI.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update IMPORTPRIBUDNI inner join ARTEAN_K on "
    sSQL = sSQL & " ARTEAN_K.EAN = IMPORTPRIBUDNI.EAN2 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRIBUDNI.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRIBUDNI.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where IMPORTPRIBUDNI.ean2  <> '0' "
    sSQL = sSQL & " and IMPORTPRIBUDNI.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update IMPORTPRIBUDNI inner join ARTEAN_K on "
    sSQL = sSQL & " ARTEAN_K.EAN = IMPORTPRIBUDNI.EAN3 "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRIBUDNI.AWM = '0' "
    sSQL = sSQL & " ,IMPORTPRIBUDNI.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where IMPORTPRIBUDNI.ean3  <> '0' "
    sSQL = sSQL & " and IMPORTPRIBUDNI.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update IMPORTPRIBUDNI i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.LIBESNR = i.LIBESNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.AWM = '0' "
    sSQL = sSQL & " ,i.artnr = ARTLIEF.artnr  "
    sSQL = sSQL & " where i.artnr is null "
'    sSQL = sSQL & " and i.LIBESNR <> '0000000' "

    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Neue Artikel werden ermittelt(9)......", Label1(4)

    sSQL = "Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,i.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,i.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where i.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(10).........", Label1(4)
    
    'Lekpreisver‰nderungen anzeigen
    
    sSQL = "Update IMPORTPRIBUDNI i inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = i.LINR and ARTLIEF.artnr = i.artnr "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.LEKALT = ARTLIEF.LEKPR "
    sSQL = sSQL & " where not i.artnr is null "
    sSQL = sSQL & " and i.rkz  = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LEKALT from IMPORTPRIBUDNI where not LEKALT is null "
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
    
    sSQL = "Update IMPORTPRIBUDNI set LEKALT = 0 "
    sSQL = sSQL & " where not artnr is null "
    sSQL = sSQL & " and rkz  = 'N' "
    sSQL = sSQL & " and Round(lekneu,2) = Round(LEKALT,2) "
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Lekpreisver‰nderungen anzeigen
    
   
    
    anzeige "normal", "Neue Artikel werden ermittelt(12)...", Label1(4)
    
    sSQL = " Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.PREISSCHU = ARTIKEL.PREISSCHU "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    anzeige "normal", "Neue Artikel werden ermittelt(12a)...", Label1(4)
    
    sSQL = " Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.Inhalt = ARTIKEL.Inhalt "
    sSQL = sSQL & " where i.Inhalt = 0 and ARTIKEL.Inhalt > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Neue Artikel werden ermittelt(12b)...", Label1(4)
    
    sSQL = " Update IMPORTPRIBUDNI i inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = i.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " i.Inhaltbez = ARTIKEL.Inhaltbez "
    sSQL = sSQL & " where i.Inhaltbez = '' and ARTIKEL.Inhaltbez <> '' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    anzeige "normal", "Neue Artikel werden ermittelt(13)......", Label1(4)
    
    
    
    Status_Ermitteln
    
    
    'Zieh doch mal  den freien Artikelnummernkreis hoch
    
    
    anzeige "normal", "F¸r neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)
    
    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)
    
    
    sSQL = "Select * from IMPORTPRIBUDNI where awm = '98' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!artnr = lreservArtnr
            rsrs.Update
            
            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verf¸gung (Einstellungen ¸berpr¸fen).", Label1(4)
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
    
    
    
    FormatiereBildschirmdatenBUDNI = True
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereBildschirmdatenBUDNI"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."

    Fehlermeldung1

End Function

Private Sub Command5_Click(Index As Integer)
 On Error GoTo LOKAL_ERROR
 
    Dim sdatname        As String
    Dim i               As Integer
    Dim sSQL            As String
    Dim lLfnr           As Long
    Dim cLfnr           As String
    Dim rsrs            As DAO.Recordset
    Dim bfound_komplett As Boolean
    
    bfound_komplett = False
 
    Select Case Index
    
        Case 0
            Unload frmWKL196
        Case 1      'Budni Stammdaten einlesen
        
            Timer1.Enabled = False
            Command5(1).Enabled = False
        
            'Ablaufprotokoll f¸llen
            'Etiketten erstellen
            'dem Anwender ein ‹bernahmeergebnis zeigen
            
            CreateTableT2 "BUORDER", gdBase
            
            File1.Path = gsKinPfad 'Standard In Pfad
            File1.Pattern = "BUDNI*.DRO"
            File1.Refresh
            
            'Ist eine komplett-Datei enthalten?
            If File1.ListCount > 0 Then
                'Datei/en stehen an
                For i = 0 To File1.ListCount - 1
                    sdatname = File1.list(i)
                    If UCase(sdatname) = "BUDNI_KOMPLETT.DRO" Then
                        bfound_komplett = True
                        Exit For
                    End If
                Next i
            End If
            
            If bfound_komplett = True Then
            
                sdatname = "BUDNI_KOMPLETT.DRO"
                lLfnr = 0
                        
                sSQL = "Insert into BUORDER (lfnr,DATNAME)"
                sSQL = sSQL & " Values ( "
                sSQL = sSQL & " " & lLfnr & " "
                sSQL = sSQL & ", '" & sdatname & "' "
                sSQL = sSQL & " ) "
                gdBase.Execute sSQL, dbFailOnError
            
            Else
            
                If File1.ListCount > 0 Then
                    'Datei/en stehen an
                    For i = 0 To File1.ListCount - 1
                        sdatname = File1.list(i)
                        cLfnr = Mid(sdatname, 7, 4)
                        lLfnr = Val(cLfnr)
                        
                        sSQL = "Insert into BUORDER (lfnr,DATNAME)"
                        sSQL = sSQL & " Values ( "
                        sSQL = sSQL & " " & lLfnr & " "
                        sSQL = sSQL & ", '" & sdatname & "' "
                        sSQL = sSQL & " ) "
                        gdBase.Execute sSQL, dbFailOnError
                    Next i
                End If
            End If
            
            sSQL = "Select * from BUORDER order by lfnr asc"
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                Do While Not rsrs.EOF
                    If Not IsNull(rsrs!Datname) Then
                        sdatname = rsrs!Datname
                        
                        BudniDatenEinlesen gsKinPfad, sdatname, bfound_komplett
                        
                    End If
                
                rsrs.MoveNext
                Loop
            End If
            rsrs.Close: Set rsrs = Nothing
            
            loeschNEW "BUORDER", gdBase
            
        Case 2
            anzeige "normal", "Druckvorschau wird erstellt...", Label1(4)
            reportbildschirm "", "aWKL196"
            
            
        Case 3
        
            'EX Artikel als Etiketten
            Dim lLinr As Long
            lLinr = checkBUDNIinLISRT()
            
            Screen.MousePointer = 11
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artikel.Artnr "
            sSQL = sSQL & ", Artikel.BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", Artikel.KVKPR1 as VKPR "
            sSQL = sSQL & ", Artlief.LIBESNR "
            sSQL = sSQL & ", Artikel.EAN "
            sSQL = sSQL & ", Artikel.LPZ "
            sSQL = sSQL & ", Artlief.LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from Artlief inner join Artikel on Artlief.artnr = artikel.artnr "
            sSQL = sSQL & " Where Artlief.rkz = 'J'"
            sSQL = sSQL & " and Artlief.linr = " & lLinr
            sSQL = sSQL & " and Artlief.Exdat = " & CLng(DateValue(Now))
            
            If chk_vk.Value = vbChecked Then
                sSQL = sSQL & " and Artikel.Artnr in (Select Artnr from Kassjour)"
            End If
            
            gdBase.Execute sSQL, dbFailOnError
            
            gbEtiExArtikel = True
            glEtiExArtikel_linr = lLinr

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
            
            glEtiExArtikel_linr = 0
            gbEtiExArtikel = False
            
        Case 4
''''            'KVK ƒnderung Artikel als Etiketten
''''
''''            Screen.MousePointer = 11
''''
''''            loeschNEW "LSTEETI", gdBase
''''            CreateTableT2 "LSTEETI", gdBase
''''
''''            sSQL = "Insert into LSTEETI select Artnr "
''''            sSQL = sSQL & ", BEZEICH "
''''            sSQL = sSQL & ", 1 as BESTAND "
''''            sSQL = sSQL & ", 1 as ANZAHL "
''''            sSQL = sSQL & ", KVKNEU as VKPR "
''''            sSQL = sSQL & ", LIBESNR "
''''            sSQL = sSQL & ", EAN "
''''            sSQL = sSQL & ", LPZ "
''''            sSQL = sSQL & ", LINR "
''''            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
''''            sSQL = sSQL & " from ImportpriBUDNI_Dauer "
''''            sSQL = sSQL & " Where Round(KVKALT, 2) <> Round(KVKNEU, 2) and RKZ = 'N'"
''''            gdBase.Execute sSQL, dbFailOnError
''''
''''            gsETILS = "aus Lieferschein"
''''
''''            frmWKL30.Show 1
            
        Case 5
        
            'VK Pflichtpreisanpassungen als Etiketten
            
            Screen.MousePointer = 11
            
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            
            
            Dim rsLi As DAO.Recordset
            Dim sBudniLinr As String
            sBudniLinr = ""
    
            sSQL = "select LINR from LISRT where FORMAT = 'EDIBUDNI' "
            Set rsLi = gdBase.OpenRecordset(sSQL)
            If Not rsLi.EOF Then
                sBudniLinr = Trim(rsLi!linr)
            End If
            rsLi.Close: Set rsLi = Nothing
            
            
            
            
            sSQL = "Insert into LSTEETI select Artlief.Artnr "
            sSQL = sSQL & ", '' as BEZEICH "
            sSQL = sSQL & ", 1 as BESTAND "
            sSQL = sSQL & ", 1 as ANZAHL "
            sSQL = sSQL & ", 0 as VKPR "
            sSQL = sSQL & ", Artlief.LIBESNR "
            sSQL = sSQL & ", '' as EAN "
            sSQL = sSQL & ", 0 as LPZ "
            sSQL = sSQL & ", Artlief.LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from Artlief inner join  BUDNI_UMLEITUNG on Artlief.libesnr = BUDNI_UMLEITUNG.uml_libesnr "
            sSQL = sSQL & " Where Artlief.linr = " & sBudniLinr
'            sSQL = sSQL & " and Artlief.rkz = 'N'"
            
            If Check1.Value = vbChecked Then
                sSQL = sSQL & " and Artlief.Artnr in (Select Artnr from Kassjour)"
            End If
'            MsgBox sSQL
            
            
            gdBase.Execute sSQL, dbFailOnError
            
            
            sSQL = "Update LSTEETI t inner join ARTIKEL a on t.artnr = a.artnr "
            sSQL = sSQL & " set t.VKPR = a.kvkpr1 "
            sSQL = sSQL & ", t.BEZEICH = a.BEZEICH "
            sSQL = sSQL & ", t.EAN = a.EAN "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
            
        Case 6
        
            'neue Artikel als Etiketten
            
            Screen.MousePointer = 11
            
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
            sSQL = sSQL & " from ImportpriBUDNI_Dauer "
            sSQL = sSQL & " where awm = '98' "
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
        
            frmWKL30.Show 1
        
    
    End Select
    
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    loeschNEW "STADAPROBUDNI", gdBase
    CreateTableT2 "STADAPROBUDNI", gdBase
    
    lesenEinstellungen
    iSec = 0
    
    If Not NewTableSuchenDBKombi("BUDNIE", gdBase) Then 'das erste Mal
        CreateTableT2 "BUDNIE", gdBase
        
        sSQL = "Insert into BUDNIE (linr) values (0)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
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
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "REWE_AGN", gdApp
    loeschNEW "MEISTER", gdApp
'    loeschNEW "STADAPROREWE", gdBase
    loeschNEW "ImportDupli", gdApp
    loeschNEW "IMPORTPRIBUDNI", gdBase
    loeschNEW "IMPORTPRIBUDNI_dauer", gdBase
    loeschNEW "IMPORTPRIBUDNI", gdApp
    
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
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Timer1.Enabled = False
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        
        Case Is = 2
            cValid = "1234567890," & Chr$(8)
        
    End Select
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If Index = 2 And cZeichen = "," Then
        If InStr(Text1(Index).Text, ",") > 0 Then
            KeyAscii = 0
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    iSec = iSec + 1
    
    If iSec >= 10 Then
        Unload frmWKL196
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil neue Budni Artikeldaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

