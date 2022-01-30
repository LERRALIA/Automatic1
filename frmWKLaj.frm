VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKLaj 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Filialtausch"
   ClientHeight    =   8625
   ClientLeft      =   1335
   ClientTop       =   1620
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
   Icon            =   "frmWKLaj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   375
      Left            =   11400
      TabIndex        =   65
      ToolTipText     =   "Einstellungen"
      Top             =   120
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "E"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame4 
      Height          =   2775
      Left            =   4800
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ohne Druck - enthaltene Kundenbestellungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   4815
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ohne Druck - Tauschkisteninhalt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   78
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ohne Druck - Kistenübersicht"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   3495
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Left            =   5880
         TabIndex        =   67
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "P"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "verteilte Artikel als Datei speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   3615
      End
      Begin sevCommand3.Command Command3 
         Height          =   345
         Left            =   5880
         TabIndex        =   45
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "x"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Left            =   5280
         TabIndex        =   43
         Top             =   2280
         Width           =   375
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
         Caption         =   "..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   5760
         TabIndex        =   40
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Die Datei einer Filiale unter diesem Pfad speichern."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   6255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pfad"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Filiale"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Einstellungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   35
      Top             =   5040
      Width           =   11655
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
      Height          =   855
      Left            =   0
      TabIndex        =   26
      Top             =   7560
      Width           =   8175
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "5"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "6"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   6
         Left            =   3720
         TabIndex        =   11
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "7"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   7
         Left            =   4320
         TabIndex        =   12
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "8"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   8
         Left            =   4920
         TabIndex        =   13
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "9"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   9
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "0"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   10
         Left            =   6120
         TabIndex        =   15
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   11
         Left            =   6720
         TabIndex        =   16
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   ">"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   615
         Index           =   12
         Left            =   7320
         TabIndex        =   17
         Top             =   120
         Width           =   600
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "C"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   9600
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   7815
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Datei versenden"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   33
         Top             =   2520
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Datei leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   7
         Left            =   4440
         TabIndex        =   36
         Top             =   2520
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Einstellungen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artnr   Artikelbezeichnung                 Menge an Fil          Datum/Zeit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   7695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tausch-Protokoll"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1455
      End
   End
   Begin Threed.SSCommand Command1 
      Height          =   495
      Index           =   1
      Left            =   10080
      TabIndex        =   29
      ToolTipText     =   "Schließen"
      Top             =   7680
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Schließen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   10080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   51
      Top             =   4800
      Width           =   11655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   3960
      Width           =   11895
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   8640
         MaxLength       =   3
         TabIndex        =   59
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "akt. Monat"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   55
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gestern"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   54
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Heute"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   53
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin sevCommand3.Command Command9 
         Height          =   375
         Index           =   0
         Left            =   6780
         TabIndex        =   52
         Top             =   360
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
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   7200
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   2160
         TabIndex        =   47
         Tag             =   "2"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   50
         Top             =   360
         Width           =   375
      End
      Begin Threed.SSCommand Command1 
         Height          =   350
         Index           =   2
         Left            =   10080
         TabIndex        =   3
         Top             =   15
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   617
         _StockProps     =   78
         Caption         =   "Suchen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Command1 
         Height          =   350
         Index           =   4
         Left            =   10080
         TabIndex        =   56
         Top             =   385
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   617
         _StockProps     =   78
         Caption         =   "Drucken"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin sevCommand3.Command Command9 
         Height          =   420
         Index           =   20
         Left            =   1680
         TabIndex        =   69
         ToolTipText     =   "Kalender"
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   741
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   420
         Index           =   21
         Left            =   3720
         TabIndex        =   70
         ToolTipText     =   "Kalender"
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   741
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   165
         Index           =   1
         Left            =   1320
         TabIndex        =   72
         Top             =   600
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
         Index           =   0
         Left            =   1320
         TabIndex        =   73
         Top             =   360
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
      Begin sevCommand3.Command Command7 
         Height          =   165
         Index           =   3
         Left            =   3360
         TabIndex        =   74
         Top             =   600
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
         Index           =   2
         Left            =   3360
         TabIndex        =   75
         Top             =   360
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8640
         TabIndex        =   58
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9360
         TabIndex        =   28
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel-Nr:"
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
         Left            =   7200
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant:"
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
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   2160
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Empfänger bestimmen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   3855
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         MultiSelect     =   2  'Erweitert
         TabIndex        =   64
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   62
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   1
         Top             =   840
         Width           =   3615
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Tauschen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   8
         Left            =   2880
         TabIndex        =   61
         Top             =   2520
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "freie Nr"
      End
      Begin Threed.SSCommand Command1 
         Height          =   375
         Index           =   9
         Left            =   2880
         TabIndex        =   63
         Top             =   1680
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Senden"
         Enabled         =   0   'False
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   68
         Top             =   2600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Transportverpackung Nr"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   60
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Empfänger bestimmen"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "keine Filiale festgelegt"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3615
      End
   End
   Begin Threed.SSCommand Command1 
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   57
      Top             =   7800
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Löschen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand Command1 
      Height          =   495
      Index           =   10
      Left            =   5760
      TabIndex        =   76
      Top             =   1320
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "nochmals Senden"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Die Datei ist nicht angekommen?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5760
      TabIndex        =   77
      Top             =   840
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   11760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filialtausch"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmWKLaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function fnPruefeFilialTauschWKLaj() As Integer
    On Error GoTo LOKAL_ERROR

    fnPruefeFilialTauschWKLaj = 0

    If Val(Left(Label2.Caption, 2)) = 0 Then
        fnPruefeFilialTauschWKLaj = 1
        Exit Function
    End If
    
    
    If frmWKL20!List1.ListCount = 0 Then
        fnPruefeFilialTauschWKLaj = 2
        Exit Function
    End If

    If Val(Left(Label2.Caption, 2)) = Val(gcFilNr) Then
        fnPruefeFilialTauschWKLaj = 3
        Exit Function
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeFilialTauschWKLaj"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub LeereDialogWKLaj()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    For i = 0 To 4
        Text1(i).Text = ""
    Next i
    List1.Clear
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKLaj"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseFilialenWKLaj()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim cFeld As String
    
    cSQL = "Select * from FILIALEN order by FILIALNR "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
                cFeld = rsrs!FILIALNR
            Else
                cFeld = "-1"
            End If
            If Val(cFeld) > 0 Then
                cLBSatz = cFeld & " "
                If Not IsNull(rsrs!Filialname) Then
                    cFeld = rsrs!Filialname
                Else
                    cFeld = ""
                End If
                cLBSatz = cLBSatz & cFeld
                
                Combo1.AddItem cLBSatz
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseFilialenWKLaj"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TauscheArtikelWKLaj()
    On Error GoTo LOKAL_ERROR
        
    Dim lAnzSatz As Long
    Dim lAktSatz    As Long
    Dim cLBSatz     As String
    
    Dim rsrs        As Recordset
    Dim rsTausch    As Recordset
    Dim rsZB        As Recordset
    Dim cSQL        As String
    
    Dim lDatum      As Long
    Dim czeit       As String
    Dim ctmp        As String
    Dim lMenge      As Long
    Dim lartnr      As Long
    Dim cBezeich    As String
    Dim lLinr       As Long
    Dim lLpz        As Long
    Dim lFilVon     As Long
    Dim lFilAn      As Long
    Dim lKasNum     As Long
    Dim dEkpr       As Double
    Dim dLEKPR      As Double
    Dim cVKPR       As String
    
    Dim lKJADate    As Long
    Dim cKJAZeit    As String
    Dim lKJBediener As Long
    Dim dKJBest1    As Double
'    Dim dBestand    As Double
    
    
    lKJBediener = 0
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")
    
    cSQL = "Select * from TAUSCH where ARTNR = -1"
    Set rsTausch = gdBase.OpenRecordset(cSQL)
    
    lDatum = Fix(Now)
    czeit = Format$(Now, "HH:MM:SS")
    lFilVon = Val(gcFilNr)
    lFilAn = Val(Left(Label2.Caption, 2))
    lKasNum = Val(gcKasNum)
    
    lAnzSatz = frmWKL20!List1.ListCount
    
    'Grund.cfg
    Dim cpfaddb As String
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    If FileExists(cpfaddb & "Grund.cfg") Then
        schreibe_xml_file gcFilNr, CStr(lFilAn)
    End If
    'Ende Grund.cfg
    

    'Hier den Warenkorb vorher zusammenfassen
    Dim ltranspack As Double
    ltranspack = 0
    
    If Check1.Value = vbChecked Then
        'Transpack gesetzt?
        
        If Label12.Caption <> "" Then
            ltranspack = CDbl(Label12.Caption)
        Else
            ltranspack = 0
        End If
            
        trageinTrans CLng(ltranspack)
            
        'Transpack gesetzt? --- Ende
    End If
    
    
    
    
    
    'Ende
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        lMenge = Val(Left(cLBSatz, 5))
        lartnr = Val(Mid(cLBSatz, 7, 6))
        cVKPR = Trim(Mid(cLBSatz, 50, 9))
        
        ctmp = Mid(cLBSatz, 148, 3)
        ctmp = Trim$(ctmp)
        lKJBediener = Val(ctmp)
        
        ctmp = Mid(cLBSatz, 138, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)
        dKJBest1 = Val(ctmp)
        
        
        
        If glAutoKundnrforKundBest > 0 Then
        
            If glAutoAusSchFiliale = lFilAn Then
            
            Else
                If dKJBest1 <= 0 Then
'                    If HatArtikelVerkäufe(lartnr) = False Then
'                        insertKundBest glAutoKundnrforKundBest, lartnr, "1", lKJBediener
'                    End If

                    If MBgleichNull(lartnr) = True Then
                        insertKundBest glAutoKundnrforKundBest, lartnr, "1", lKJBediener
                    End If


                End If
            End If
        End If
        
        
        
        
        cSQL = "Select * from ARTIKEL where ARTNR = " & Trim$(Str$(lartnr))
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
        
            rsTausch.AddNew
            rsTausch!ADATE = lDatum
            rsTausch!AZEIT = czeit
            rsTausch!Menge = lMenge
            rsTausch!artnr = lartnr
            
            If Not IsNull(rsrs!BEZEICH) Then
                rsTausch!BEZEICH = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!linr) Then
                rsTausch!linr = rsrs!linr
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                rsTausch!LPZ = rsrs!LPZ
            End If
            
            If Not IsNull(rsrs!ekpr) Then
                rsTausch!ekpr = rsrs!ekpr
            End If
            
            
            rsTausch!lekpr = 0
            
            Dim sMin_Linr As String
            Dim rsLIEF As DAO.Recordset
    
            sMin_Linr = ermLiefmitGroesstemLEKPR(Str$(lartnr))

            Set rsLIEF = gdBase.OpenRecordset("select * from Artlief where artnr = " & Str$(lartnr) & " and linr = " & sMin_Linr)
            If Not rsLIEF.EOF Then
                If Not IsNull(rsLIEF!lekpr) Then
                    rsTausch!lekpr = rsLIEF!lekpr
                End If
            End If
            rsLIEF.Close: Set rsLIEF = Nothing
            
            
            
            
            
'            If Not IsNull(rsrs!lekpr) Then
'                rsTausch!lekpr = rsrs!lekpr
'            End If
'
            If IstDasEineWGN(lartnr) Then
                rsTausch!vkpr = cVKPR
            Else
                If Not IsNull(rsrs!vkpr) Then
                    rsTausch!vkpr = rsrs!vkpr
                End If
            End If
            
            rsTausch!BEDIENER = lKJBediener
            rsTausch!FIL_VON = lFilVon
            rsTausch!FIL_AN = lFilAn
            
            'hier nicht mehr Kassennummer sondern Bestand - für Bestandsprotokollierung wichtig
            rsTausch!kasnum = dKJBest1
            
            If Not IsNull(rsrs!KVKPR1) Then
                If CDbl(rsrs!KVKPR1) > 0 Then
                    rsTausch!KVKPR1 = rsrs!KVKPR1
                Else
                    If IstDasEineWGN(lartnr) Then
                        rsTausch!KVKPR1 = cVKPR
                    Else
                        rsTausch!KVKPR1 = rsrs!KVKPR1
                    End If
                End If
            End If
            
            
            rsTausch!TRANSPACK = ltranspack
            
            rsTausch!SENDOK = False
            rsTausch.Update
            
            insertUNTERWF lKJADate, cKJAZeit, CStr(lartnr), CStr(lMenge), CStr(lFilAn)
            
            'Jetzt Kassenbon für Protokoll schreiben
        End If
        rsrs.Close: Set rsrs = Nothing
        

        If Check1.Value = vbChecked Then
            'Transpack gesetzt?
            
            If ltranspack <> 0 Then
                schreibeVerteiluFilbesteZ CStr(lartnr), lMenge, CInt(lFilAn), ltranspack, lKJBediener, lDatum, czeit
            End If
            
            fülleliste List4
            'Transpack gesetzt? --- Ende
        End If
    Next lAktSatz
    
    rsTausch.Close
    
    frmWKL20.Label2(6).Caption = ""
    frmWKL20.Label2(5).Caption = ""
    
    Label4.Caption = "Artikelanzahl: " & ermittleAnzVerteil
    Label4.Refresh
    
    SendeDaten2DruckerNeuWKLaj
    
    frmWKL20!List1.Clear
    frmWKL20!List3.Nodes.Clear
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TauscheArtikelWKLaj"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
    
End Sub
Public Sub schreibeVerteiluFilbesteZ(cART As String, lspeich As Long, inewzu As Integer, ltranspack As Double, lbed As Long, lDatum As Long, czeit As String)
On Error GoTo LOKAL_ERROR

Dim sSQL    As String

Dim rsFB    As Recordset
Dim rsUv    As Recordset
Dim rsArt   As Recordset

Dim sBez        As String
Dim slibesnr    As String
Dim lLinr       As Long
Dim lLpz        As Long
Dim dKVK        As Double
Dim lFarbnr     As Long

sSQL = "Select * from artikel where artnr = " & cART
Set rsArt = gdBase.OpenRecordset(sSQL)
If Not rsArt.EOF Then

    If Not IsNull(rsArt!BEZEICH) Then
        sBez = rsArt!BEZEICH
    Else
        sBez = ""
    End If
    
    If Not IsNull(rsArt!LIBESNR) Then
        slibesnr = rsArt!LIBESNR
    Else
        slibesnr = ""
    End If
    
    If Not IsNull(rsArt!linr) Then
        lLinr = rsArt!linr
    Else
        lLinr = 0
    End If
    
    If Not IsNull(rsArt!LPZ) Then
        lLpz = rsArt!LPZ
    Else
        lLpz = 0
    End If
    
    If Not IsNull(rsArt!KVKPR1) Then
        dKVK = rsArt!KVKPR1
    Else
        dKVK = 0
    End If
    
    If Not IsNull(rsArt!AWM) Then
        lFarbnr = Val(rsArt!AWM)
    Else
        lFarbnr = 0
    End If
    
    Set rsFB = gdBase.OpenRecordset("FILZ")
    rsFB.AddNew
    rsFB!artnr = cART
    rsFB!BEZEICH = sBez
    rsFB!linr = lLinr
    rsFB!LPZ = lLpz
    
    rsFB!LIBESNR = slibesnr
    rsFB!KVKPR1 = dKVK
    rsFB!BESTVOR = lspeich
    rsFB!FILIALE = inewzu
    rsFB!FILVON = gcFilNr
    
    rsFB!FARBNR = lFarbnr
    rsFB!TRANSPACK = ltranspack
    
    rsFB!BEDNU = lbed
    rsFB!bedname = ermfromBed("BEDNAME", CStr(lbed))
    rsFB!openart = "OFFEN"
    rsFB!AENART = "Filialtausch WK"
    rsFB!Pcname = srechnertab
    
    
    rsFB!ADATE = lDatum 'DateValue(Now)
    rsFB!AZEIT = czeit 'TimeValue(Now)
    
    rsFB.Update
    rsFB.Close
    
        

    
End If
rsArt.Close

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "schreibeVerteiluFilbesteZ"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Public Sub trageinTrans(ltranspack As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL            As String
    Dim rec             As Recordset
    
    cSQL = "Delete from  TRANS where transpack = " & ltranspack
    gdBase.Execute cSQL, dbFailOnError
    
    Set rec = gdBase.OpenRecordset("TRANS")
    rec.AddNew
    rec!TRANSPACK = ltranspack
    rec.Update
    rec.Close
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "trageinTrans"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub loescheTausch(sZeile As String)
    On Error GoTo LOKAL_ERROR
        
    Dim rsTausch As Recordset
    Dim cSQL As String

    Dim cdat As String
    Dim lDatum As Long
    Dim czeit As String
    Dim lMenge As Long
    Dim lartnr As Long
    Dim lFilAn As Long
    Dim lBestand As Long
    
    Dim cMenge As String
    
    Dim ctemp As String
    
    cdat = Mid(sZeile, 3, 8)
    lDatum = DateValue(cdat) 'DateValue(Left(sZeile, 8))
    czeit = Mid(sZeile, 12, 8)
    lartnr = Val(Mid(sZeile, 31, 6))
    lMenge = Val(Mid(sZeile, 27, 3))
    lFilAn = Val(Mid(sZeile, 87, 2))
    
    cSQL = "Select * from TAUSCH "
    cSQL = cSQL & " where artnr = " & lartnr
    cSQL = cSQL & " and  menge = " & lMenge
    cSQL = cSQL & " and  adate = " & lDatum
    cSQL = cSQL & " and  azeit = '" & czeit & "'"
    cSQL = cSQL & " and  Fil_An = " & lFilAn
    cSQL = cSQL & " and  SENDOK = False "
'    MsgBox cSQL
    Set rsTausch = gdBase.OpenRecordset(cSQL)
    If Not rsTausch.EOF Then
        rsTausch.MoveFirst
        rsTausch.delete
    Else
        ctemp = "Dieser Artikel kann nicht mehr gelöscht werden." & vbCrLf & vbCrLf
        ctemp = ctemp & "Artikel, die mit einem 'V'(versendet) gekennzeichnet sind, können nicht mehr gelöscht werden."
        MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
        rsTausch.Close
        Exit Sub
    End If
    
    rsTausch.Close
    
    cSQL = "Select * from UNTERWF "
    cSQL = cSQL & " where artnr = " & lartnr
    cSQL = cSQL & " and  menge = " & lMenge
    cSQL = cSQL & " and  adate = " & lDatum
    cSQL = cSQL & " and  azeit = '" & czeit & "'"
    cSQL = cSQL & " and  ZIELFILIALE = " & lFilAn
    cSQL = cSQL & " and  SENDOK = False "
    
    Set rsTausch = gdBase.OpenRecordset(cSQL)
    If Not rsTausch.EOF Then
        rsTausch.MoveFirst
        rsTausch.delete
    End If
    
    rsTausch.Close
    
    'Bestand zurück !!!!!
    lBestand = 0
    cSQL = "Select Bestand from Artikel "
    cSQL = cSQL & " where artnr = " & lartnr
    Set rsTausch = gdBase.OpenRecordset(cSQL)
    If Not rsTausch.EOF Then
        If Not IsNull(rsTausch!BESTAND) Then
            lBestand = rsTausch!BESTAND
        End If
    End If
    
    lBestand = lBestand + lMenge
    
    Bestandsveraenderung CStr(lartnr), lBestand, "Rücknahme Tausch"
    
    If Check1.Value = vbChecked Then 'auch zurück schreiben
        cSQL = "Select * from Filz "
        cSQL = cSQL & " where artnr = " & lartnr
        cSQL = cSQL & " and  BESTVOR = " & lMenge
        cSQL = cSQL & " and  ADATE = " & lDatum
        cSQL = cSQL & " and  AZEIT = '" & czeit & "'"
        cSQL = cSQL & " and  FILIALE = " & lFilAn

        Set rsTausch = gdBase.OpenRecordset(cSQL)
        If Not rsTausch.EOF Then
            rsTausch.MoveFirst
            rsTausch.delete
        End If

        rsTausch.Close

    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheTausch"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR

    If Check1.Value = vbUnchecked Then
        schreibeFilProt "entfernt Haken (verteilte Artikel als Datei speichern)", "FILPROT"
        
        Label5.Visible = False
        Command1(8).Visible = False
        Label12.Visible = False
        
        Command1(9).Visible = False
        
    Else
    
        If gbFtpYes Then
            Command1(9).Enabled = True
        End If
        
        Command1(9).Visible = True
        Label5.Visible = True
        Command1(8).Visible = True
        Label12.Visible = True
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Combo1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lAktKiste As Long
    
    If Trim$(Combo1.Text) = "" Then
        Label2.Caption = "keine Filiale festgelegt"
        Label2.Refresh
    Else
        Label2.Caption = Combo1.Text
        Label2.Refresh
        
        lAktKiste = ermAktKiste(Val(Left(Label2.Caption, 2)))
        
        If lAktKiste = 0 Then
            Label12.Caption = freieTranspack
        Else
            Label12.Caption = lAktKiste
        End If
        
    End If
    Label0.Caption = "-1"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub alleKisten_Close_forspezFil(iFil As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL As String

sSQL = "Update Filz SET Openart = 'GESCHLOSSEN'"
sSQL = sSQL & " where Filiale = " & iFil & " "

gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alleKisten_Close_forspezFil"
    Fehler.gsFehlertext = "Im Programmteil Transportverpackung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermAktKiste(iFil As Integer) As Long
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermAktKiste = 0

sSQL = "Select distinct(TRANSPACK) as KISTENNR from Filz where Openart = 'OFFEN'"
sSQL = sSQL & " and Filiale = " & iFil & " "

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!KISTENNR) Then
        ermAktKiste = rsrs!KISTENNR
    End If
End If
rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAktKiste"
    Fehler.gsFehlertext = "Im Programmteil Transportverpackung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub Combo1_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Combo1.BackColor = glSelBack1
    Label0.Caption = "-1"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_GotFocus"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Shift
        Case Is = 0     'ohne Zusatztaste
        Case Is = 1     'SHIFT ist gedrückt
        Case Is = 2     'STRG ist gedrückt
        Case Is = 3     'SHIFT + STRG sind gedrückt
        Case Is = 4     'ALT ist gedrückt
            
            If KeyCode = 84 Then              'Taste Alt + t tauschen
                Command1_Click 0
                Exit Sub
            End If
            
            If KeyCode = 83 Then              'Taste Alt + S schließen
                Command1_Click 1
                Exit Sub
            End If

            
        Case Is = 5     'SHIFT + ALT sind gedrückt
            
        Case Is = 6     'ALT GR bzw. STRG + ALT sind gedrückt
            
        Case Is = 7     'SHIFT + STRG + ALT sind gedrückt
            
        Case Else       'was noch?
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_KeyUp"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    If Trim$(Combo1.Text) = "" Then
        Label2.Caption = "keine Filiale festgelegt"
        Label2.Refresh
    Else
        Label2.Caption = Combo1.Text
        Label2.Refresh
    End If
    Combo1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_LostFocus"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFeld As Integer
    
    Dim cZeichen As String
    
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "+{Tab}", True
    
    iFeld = Val(Label0.Caption)
    
    If iFeld < 0 Then
        Text1(0).SetFocus
        Combo1.BackColor = vbWhite
        Exit Sub
    End If
    
    Select Case Index
        Case 0 To 9
            Text1(iFeld).Text = Text1(iFeld).Text & Command0(Index).Caption
        Case Is = 10
            WshShell.SendKeys "+{Tab}", True
        Case Is = 11
            WshShell.SendKeys "{Tab}", True
        Case Is = 12
            Text1(iFeld).Text = ""
    End Select
    
    Text1(iFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command5_Click()
On Error GoTo LOKAL_ERROR

    Frame4.Visible = True
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lDat As Long
    
    Select Case Index
        Case Is = 1
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                lDat = lDat - 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case Is = 0
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                lDat = lDat + 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case Is = 3
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat - 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case Is = 2
            If IsDate(Text1(2).Text) = False Then
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(2).Text) = True Then
                    lDat = CLng(DateValue(Text1(2).Text))
                End If
                lDat = lDat + 1
                Text1(2).Text = Format(lDat, "DD.MM.YYYY")
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command9_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    Select Case Index
        Case 0
            Text1_KeyUp 4, vbKeyF2, 0
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
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
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

        Select Case Index
                    
            Case Is = 5     'ak monat
                Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
            Case Is = 1     'gestern
                Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
                Text1(2).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            
            Case Is = 0     'heute
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
                Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
        End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

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
            Case Is = 4
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        Text1(Index).Text = ctmp
                    End If
                    Text1(Index).SetFocus
                End If
        End Select
        
    ElseIf KeyCode = vbKeyReturn Then
        Command1_Click 2
    End If
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim iRet     As Integer
    Dim cDrucker As String
    Dim ctmp     As String
    Dim bReturn  As Boolean
    Dim bDruck As Boolean
    Dim cSQL     As String
    Dim cLBSatz  As String
    Dim lcount   As Long
    Dim bFound  As Boolean
    
    bDruck = True
    Select Case Index
        Case Is = 0     'Tauschen
            iRet = fnPruefeFilialTauschWKLaj()
            Select Case iRet
                
                Case Is = 0     'alles okay
                    If gbBONNEIN = True Then
                        If MsgBox("Bondruck ist ausgeschaltet. Möchten Sie den Bon drucken?", vbYesNo + vbQuestion, "WINKISS-Frage:") = vbYes Then
                            gbBONNEIN = False
                            bDruck = False
                        End If
                    End If
                    
                    Command1(0).Enabled = False
                    TauscheArtikelWKLaj
                    
                    If bDruck = False Then
                        gbBONNEIN = True
                    End If
                    
                    'suche drücken
                    Command1_Click 2
                    
                    Command1(0).Enabled = True
                    
                Case Is = 1     'kein Ziel
                    MsgBox "Bitte die Zielfiliale festlegen!", vbInformation, "Winkiss Hinweis:"
                    Combo1.SetFocus
                Case Is = 2     'keine Artikel
                    MsgBox "Es liegen keine Artikel zum Filialtausch vor!", vbInformation, "Winkiss Hinweis:"
                    Command1(1).SetFocus
                Case Is = 3     'Tausch in die eigene Filiale
                    MsgBox "Sie haben die eigene Filiale als Zielfiliale gewählt!", vbInformation, "Winkiss Hinweis:"
                    Combo1.SetFocus
            End Select
        Case 7
            Frame4.Visible = True
        Case Is = 1     'Schließen
            Unload frmWKLaj
            frmWKL20!Label2(5).Caption = "1"
            
        Case Is = 2     'Suchen
            LeseProtokollFilialTauschWKLaj
        Case Is = 3     'Löschen
        
            bFound = False
            For lcount = 0 To List1.ListCount - 1
                If List1.Selected(lcount) = True Then
                    bFound = True
                End If
            Next lcount
        
            If Not bFound Then
                MsgBox "Bitte einen Eintrag in der Liste markieren!", vbInformation, "Winkiss Hinweis:"
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            
            loescheTausch List1.list(List1.ListIndex)
            
            'Suche drücken
            Command1_Click 2
        Case Is = 4     'Drucken
            
            If Datendrin("TauZ", gdBase) Then
                reportbildschirm "", "ZaEN15"
            End If
            
        Case Is = 5 'Datei versenden
            Command1(5).Enabled = False
            uebertragenVerteil
            Command1(5).Enabled = True
            
        Case Is = 6 'Datei leeren

            If NewTableSuchenDBKombi("ANFILM", gdBase) Then
                cSQL = "Delete from ANFILM"
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            
            Label4.Caption = "Artikelanzahl: " & ermittleAnzVerteil
            Label4.Refresh
            List2.Clear
        Case 8
        
            iRet = fnPruefeFilialTauschWKLaj()
            Select Case iRet
                
                Case Is = 0     'alles okay
                    
                    alleKisten_Close_forspezFil Val(Left(Label2.Caption, 2))
            
                    Label12.Caption = freieTranspack
                    Label12.Refresh
                    
                Case Is = 1     'kein Ziel
                    MsgBox "Bitte die Zielfiliale festlegen!", vbInformation, "Winkiss Hinweis:"
                    Combo1.SetFocus
                Case Is = 2     'keine Artikel
                    MsgBox "Es liegen keine Artikel zum Filialtausch vor!", vbInformation, "Winkiss Hinweis:"
                    Command1(1).SetFocus
                Case Is = 3     'Tausch in die eigene Filiale
                    MsgBox "Sie haben die eigene Filiale als Zielfiliale gewählt!", vbInformation, "Winkiss Hinweis:"
                    Combo1.SetFocus
            End Select
        Case 9
            'übertragen
            
            Dim ltranspack  As Long
            Dim lFil        As Long
            Dim lFilVon     As Long
            Dim lMenge      As Long
            Dim sSQL        As String
            
            
            If Check3.Value = vbUnchecked Then
                ctmp = "Ist der Drucker funktionsbereit?" & vbCrLf & vbCrLf
                ctmp = ctmp & "Drucker an?" & vbCrLf
                ctmp = ctmp & "Druckerpapier?" & vbCrLf & vbCrLf
                ctmp = ctmp & "(Möchten Sie den Lieferschein in Zukunft nicht mehr drucken, so klicken Sie auf 'E' und aktivieren Sie die entsprechenden Häkchen.)" & vbCrLf
                iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")

                If iRet = vbYes Then

                Else
                    Screen.MousePointer = 0
                    ctmp = "Der Vorgang wurde abgebrochen." & vbCrLf
                    MsgBox ctmp, vbInformation + vbOKOnly, "Winkiss Hinweis:"
                    Exit Sub

                End If
            End If
            
            
            If List4.ListCount > 0 Then

                loeschNEW "KBLATT", gdBase
                CreateTableT2 "KBLATT", gdBase
                
                loeschNEW "KP" & srechnertab, gdBase
                CreateTableT2 "KP" & srechnertab, gdBase

                For lcount = 0 To List4.ListCount - 1
                    ltranspack = Left(Trim(List4.list(lcount)), 15)
                    lFil = CLng(Mid(Trim(List4.list(lcount)), 15, 8))
                    lFilVon = CLng(gcFilNr)
                    lMenge = CLng(Mid(Trim(List4.list(lcount)), 43, 7))
                    
                    If Check3.Value = vbUnchecked Then
                        'Kisteninhalt
                        DruckLiefschein ltranspack
                    End If
                    
                    InsertKBLATT ltranspack, 0, lFil, lFilVon, lMenge
                    
                    uebertragenVerteilZ ltranspack
                    
                Next lcount
                
                UpdateKBLATT
                If Check2.Value = vbUnchecked Then
                    'Kistenübersicht
                    reportbildschirm "", "aZEN112b"
                End If
                
                If Datendrin("KP" & srechnertab, gdBase) Then
                    'enthaltene Kundenbestellungen
                
                    loeschNEW "KPPRINT", gdBase
                    CreateTableT2 "KPPRINT", gdBase
                    
                    sSQL = "Insert into KPPRINT select * from KP" & srechnertab
                    gdBase.Execute sSQL
                    
                    sSQL = "Update KPPRINT set Art = 'Kdw/SB nicht geliefert' where transpack = 0 "
                    gdBase.Execute sSQL
                    
                    sSQL = "Update KPPRINT set Art = 'Kdw/SB geliefert' where transpack > 0 "
                    gdBase.Execute sSQL
                    
                    sSQL = "Update KPPRINT inner join kunden on KPPRINT.KUNDNR = KUNDEN.KUNDNR"
                    sSQL = sSQL & " set KPPRINT.kdname = kunden.name "
                    gdBase.Execute sSQL
                    
                    sSQL = "Update KPPRINT inner join Artikel on KPPRINT.ARTNR  = Artikel.ARTNR "
                    sSQL = sSQL & " set KPPRINT.RKZ = ARTIKEL.RKZ "
                    gdBase.Execute sSQL
    
                    If Check4.Value = vbUnchecked Then
                        'enthaltene Kundenbestellung
                        reportbildschirm "", "aZEN112d"
                    End If
                End If
                                
                If gbFtpYes Then
                    giKissFtpMode = 16
                    frmWKL38.Show 1
                End If
                
                Me.Refresh
                
                'bestimmt senden
    
                fülleliste List4
                
                Me.Refresh
            End If
            
        Case 10
            Screen.MousePointer = 0
            dlgKopieren.Show 1
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckLiefschein(ltranspack As Long)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    loeschNEW "PRINTLIEF", gdBase
    CreateTableT2 "PRINTLIEF", gdBase
    
    'BARCODE
    Dim cEANCode As String
    Dim cEAN As String
    
    cEANCode = ""
    If Len(CStr(ltranspack)) < 7 Then
        cEAN = CStr(ltranspack)
        cEAN = fnMoveArtNr2EAN8(cEAN)

        cEANCode = fnCodiereEANCode(cEAN)
    End If
    'BARCODE ENDE
    
    sSQL = "Insert into PRINTLIEF select *  from FILZ"
    sSQL = sSQL & " where transpack = " & ltranspack
    gdBase.Execute sSQL, dbFailOnError
    
    'EAN
    
    sSQL = " Update PRINTLIEF inner join Artikel on PRINTLIEF.artnr = Artikel.artnr "
    sSQL = sSQL & " set PRINTLIEF.EAN = Artikel.EAN "
    gdBase.Execute sSQL, dbFailOnError
    
   
    sSQL = "Update PRINTLIEF set BARCODE = '" & cEANCode & "'"
    sSQL = sSQL & " where transpack = " & ltranspack
    gdBase.Execute sSQL, dbFailOnError
    
    'Kisteninhalt
    reportbildschirmToPrinter "aZEN112c"
              
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckLiefschein"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Function erm_XML_count() As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As DAO.Recordset
    
    erm_XML_count = "0"
    
    If NewTableSuchenDBKombi("COUNTXML", gdBase) = False Then
        sSQL = "Create table COUNTXML (MAXNR LONG) "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Insert into COUNTXML (MAXNR) Values (0) "
        gdBase.Execute sSQL, dbFailOnError
    End If
    

    sSQL = "Select MAXNR from COUNTXML "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!MAXNR) Then
            erm_XML_count = CStr(rsrs!MAXNR + 1)
            rsrs.Edit
            rsrs!MAXNR = erm_XML_count
            rsrs.Update
        End If
    End If
    rsrs.Close
    
    While Len(erm_XML_count) < 13
        erm_XML_count = "0" & erm_XML_count
    Wend
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erm_XML_count"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub schreibe_xml_file(sVonFil As String, sAnFil As String)
On Error GoTo LOKAL_ERROR

    Dim lrow            As Long
    Dim lPos            As Long
    Dim cSatz           As String
    Dim cSQL            As String
    Dim cMenge          As String
    Dim cArtNr          As String
    Dim cMwst           As String
    Dim cKVkPr1         As String
    Dim cKVkPr1_netto   As String
    
    Dim dKVkPr1         As Double
    Dim dKVkPr1_netto   As Double
    
    Dim cPfad           As String
    Dim iFileNr         As Integer
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim sTime           As String
    Dim sDate           As String
    Dim lcount          As Long
    Dim lanzseg         As Long
    Dim cDatname        As String
    
    Dim sFilID          As String
    Dim sFILname        As String
    Dim sFilStrasse     As String
    Dim sFilOrt         As String
    Dim sFilPLZ         As String
    Dim cOrderNr        As String
    Dim sTimestamp      As String
    
    
    Select Case sAnFil
        Case "1"
            sFilID = "F001"
            sFILname = "SLG Service Logistik Günthersdorf GmbH"
            sFilStrasse = "Nordpark 7"
            sFilPLZ = "06237"
            sFilOrt = "Leuna"
        Case "2"
            sFilID = "F002"
            sFILname = "Point Rouge Parfümerie"
            sFilStrasse = "Unter den Linden 16"
            sFilPLZ = "14542"
            sFilOrt = "Werder"
        Case "3"
            sFilID = "F003"
            sFILname = "Point Rouge Parfümerie"
            sFilStrasse = "Brandenburger Str. 29"
            sFilPLZ = "14467"
            sFilOrt = "Potsdam"
    End Select

    lanzseg = 0
    lcount = 0
    
    sTime = Format$(TimeValue(Now), "HHMMSS")
    sDate = Format$(DateValue(Now), "DDMMYYYY")
    


    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    cPfad = cPfad & "XML\"

    cOrderNr = erm_XML_count
    cDatname = sVonFil & sAnFil & "W" & cOrderNr & ".xml"

    Kill cPfad & cDatname

    iFileNr = FreeFile
    Open cPfad & cDatname For Binary As #iFileNr
    
    cSatz = ""
    cSatz = cSatz & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & " ?>" & vbCrLf
    cSatz = cSatz & "<Timezone_B2Orders>" & vbCrLf
    cSatz = cSatz & "<ILN>4043714000009</ILN>" & vbCrLf
    
    cSatz = cSatz & "<PurchaseOrder>" & vbCrLf
    cSatz = cSatz & "<KZ>B2C</KZ>" & vbCrLf
    cSatz = cSatz & "<ILNB2>4043714000207</ILNB2>" & vbCrLf
    
    cSatz = cSatz & "<REAdresse>" & vbCrLf
    cSatz = cSatz & "<Customer_ID>" & sFilID & "</Customer_ID>" & vbCrLf
    cSatz = cSatz & "<Name1>" & sFILname & "</Name1>" & vbCrLf
    cSatz = cSatz & "<Name2></Name2>" & vbCrLf
    cSatz = cSatz & "<Strasse>" & sFilStrasse & "</Strasse>" & vbCrLf
    cSatz = cSatz & "<Ort>" & sFilOrt & "</Ort>" & vbCrLf
    cSatz = cSatz & "<PLZ>" & sFilPLZ & "</PLZ>" & vbCrLf
    cSatz = cSatz & "<Land>DE</Land>" & vbCrLf
    cSatz = cSatz & "<Geburtstag>01.01.1970</Geburtstag>" & vbCrLf
    cSatz = cSatz & "<Telefon></Telefon>" & vbCrLf
    cSatz = cSatz & "<Mobil></Mobil>" & vbCrLf
    cSatz = cSatz & "<Fax></Fax>" & vbCrLf
    cSatz = cSatz & "<EMail></EMail>" & vbCrLf
    cSatz = cSatz & "</REAdresse>" & vbCrLf
    
    cSatz = cSatz & "<LEAdresse>" & vbCrLf
    cSatz = cSatz & "<Customer_ID>" & sFilID & "</Customer_ID>" & vbCrLf
    cSatz = cSatz & "<Name1>" & sFILname & "</Name1>" & vbCrLf
    cSatz = cSatz & "<Name2></Name2>" & vbCrLf
    cSatz = cSatz & "<Strasse>" & sFilStrasse & "</Strasse>" & vbCrLf
    cSatz = cSatz & "<Ort>" & sFilOrt & "</Ort>" & vbCrLf
    cSatz = cSatz & "<PLZ>" & sFilPLZ & "</PLZ>" & vbCrLf
    cSatz = cSatz & "<Land>DE</Land>" & vbCrLf
    cSatz = cSatz & "<Geburtstag>01.01.1970</Geburtstag>" & vbCrLf
    cSatz = cSatz & "<Telefon></Telefon>" & vbCrLf
    cSatz = cSatz & "<Mobil></Mobil>" & vbCrLf
    cSatz = cSatz & "<Fax></Fax>" & vbCrLf
    cSatz = cSatz & "<EMail></EMail>" & vbCrLf
    cSatz = cSatz & "</LEAdresse>" & vbCrLf
    
    cSatz = cSatz & "<Teillieferung>FALSE</Teillieferung>" & vbCrLf
    cSatz = cSatz & "<PaymentID></PaymentID>" & vbCrLf
    

    cSatz = cSatz & "<OrderNo>W" & cOrderNr & "</OrderNo>" & vbCrLf

    sTimestamp = Format(DateValue(Now), "YYYY-MM-DD") & "T" & Format(TimeValue(Now), "HH:MM:SS")
    cSatz = cSatz & "<OrderDate>" & sTimestamp & "</OrderDate>" & vbCrLf
'    cSatz = cSatz & "<OrderDate>2013-01-14T04:40:59</OrderDate>" & vbCrLf
    cSatz = cSatz & "<OrderCurrency>EUR</OrderCurrency>" & vbCrLf
    cSatz = cSatz & "<TermOfShipment>9000</TermOfShipment>" & vbCrLf
    cSatz = cSatz & "<CostOfShipmentNetto>0.0000</CostOfShipmentNetto>" & vbCrLf
    cSatz = cSatz & "<CostOfShipmentBrutto>0.0000</CostOfShipmentBrutto>" & vbCrLf
    cSatz = cSatz & "<CostOfPaymentNetto>0.0000</CostOfPaymentNetto>" & vbCrLf
    cSatz = cSatz & "<CostOfPaymentBrutto>0.0</CostOfPaymentBrutto>" & vbCrLf
    cSatz = cSatz & "<TermsOfPayment>0000</TermsOfPayment>" & vbCrLf
    cSatz = cSatz & "<TotalPaymantNetto>0.0000</TotalPaymantNetto>" & vbCrLf
    cSatz = cSatz & "<TotalPaymentBrutto>0.0000</TotalPaymentBrutto>" & vbCrLf
    cSatz = cSatz & "<TotalQuantity>0</TotalQuantity>" & vbCrLf
    cSatz = cSatz & "<OrderItem>" & vbCrLf
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    Dim lAktSatz As Long
    Dim cLBSatz As String
    Dim lAnzSatz As Long
    lAnzSatz = frmWKL20!List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        cMenge = CStr(Val(Left(cLBSatz, 5)))
        cArtNr = CStr(Val(Mid(cLBSatz, 7, 6)))

        cMwst = "V"
        dKVkPr1 = 0
        dKVkPr1_netto = 0
        
        cSQL = "Select * from Artikel where artnr = " & cArtNr
        Set rsArt = gdBase.OpenRecordset(cSQL)
        If Not rsArt.EOF Then
            
            If Not IsNull(rsArt!ekpr) Then
                dKVkPr1 = rsArt!ekpr
            End If
            
            If Not IsNull(rsArt!MWST) Then
                cMwst = rsArt!MWST
            End If
            
            Select Case cMwst
                Case Is = "V"
                   dKVkPr1_netto = (dKVkPr1 / (100 + gdMWStV)) * 100
                Case Is = "E"
                   dKVkPr1_netto = (dKVkPr1 / (100 + gdMWStE)) * 100
                Case Is = "O"
                   dKVkPr1_netto = (dKVkPr1 / 100) * 100
            End Select

        End If
        rsArt.Close
            
        cKVkPr1 = Format(CStr(dKVkPr1), "######0.0000")
        cKVkPr1 = SwapStr(cKVkPr1, ",", ".")
        
        cKVkPr1_netto = Format(CStr(dKVkPr1_netto), "######0.0000")
        cKVkPr1_netto = SwapStr(cKVkPr1_netto, ",", ".")
        
        cMenge = Format(CStr(cMenge), "######0.0000")
        cMenge = SwapStr(cMenge, ",", ".")
        
        cSatz = ""
        cSatz = cSatz & "<OrderSKU>" & vbCrLf
        cSatz = cSatz & "<EAN>" & cArtNr & "</EAN>" & vbCrLf
        cSatz = cSatz & "<refGS>0</refGS>" & vbCrLf
        cSatz = cSatz & "<BuyingPriceNetto>" & cKVkPr1_netto & "</BuyingPriceNetto>" & vbCrLf
        cSatz = cSatz & "<BuyingPriceBrutto>" & cKVkPr1 & "</BuyingPriceBrutto>" & vbCrLf
        cSatz = cSatz & "<Quantity>" & cMenge & "</Quantity>" & vbCrLf
        cSatz = cSatz & "</OrderSKU>" & vbCrLf

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
            
    Next lAktSatz

    cSatz = ""
    cSatz = cSatz & "</OrderItem>" & vbCrLf
    cSatz = cSatz & "</PurchaseOrder>" & vbCrLf
    cSatz = cSatz & "</Timezone_B2Orders>" & vbCrLf
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz

    Close iFileNr
    
    Dim bmerke  As Boolean
    bmerke = gbFTPautomatic
    
    If gbFtpYes Then
    
        gbFTPautomatic = True
        giKissFtpMode = 32 'FTPMODE= 32 , XML - Ordner leeren abschicken
        frmWKL38.Show 1
        gbFTPautomatic = bmerke
    
    End If


Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "schreibe_xml_file"
        Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
        Fehlermeldung1
    End If
End Sub

Public Sub abschickenVerteilZ()
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rec         As Recordset
    
    

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "abschickenVerteilZ"
    Fehler.gsFehlertext = "Beim Erstellen/Versenden der WE - Datei ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
End Sub
Private Sub SpeichernFBZproKiste(Pfad As String, ltranspack As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim db          As Database
    Dim rsFB        As Recordset
    Dim iFil        As Integer
    Dim cDatname    As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim cPfad       As String
    Dim cEAN        As String
    Dim cEANCode    As String
    
    cPfad = Pfad
        
    If Right$(cPfad, 1) = "\" Then
        cPfad = Left$(cPfad, Len(cPfad) - 1)
    End If
        
    sSQL = "select distinct(filiale)as fili from FILZ where transpack = " & ltranspack
    Set rsFB = gdBase.OpenRecordset(sSQL)
    If Not rsFB.EOF Then
        rsFB.MoveFirst
        Do While Not rsFB.EOF
        
            If Not IsNull(rsFB!fili) Then
                iFil = rsFB!fili
                cDatname = ermittlenextNumberZ(iFil, ltranspack)
                
                If Not FindFile(cPfad, cDatname & ".MDB") Then
                    Set db = CreateDatabase(Pfad & cDatname & ".mdb", dbLangGeneral, dbVersion40)
                    db.Close
                End If
                
                loeschNEW "PRINTFILZ", gdBase
                CreateTableT2 "PRINTFILZ", gdBase
                
                sSQL = "Insert into PRINTFILZ select *  from FILZ"
                sSQL = sSQL & " where filiale = " & iFil
                gdBase.Execute sSQL, dbFailOnError
                
                cEANCode = ""
                If Len(cDatname) < 12 Then
                    cEAN = Right(cDatname, Len(cDatname) - 5)
                    cEAN = fnMoveArtNr2EAN8(cEAN)

                    cEANCode = fnCodiereEANCode(cEAN)

                    sSQL = "Update PRINTFILZ set BARCODE = '" & cEANCode & "'"
                    gdBase.Execute sSQL, dbFailOnError
                End If
                    
                sSQL = "Update PRINTFILZ set Datname = '" & cDatname & "'"
                gdBase.Execute sSQL, dbFailOnError
                
                UpdateKBLATTBCODE cEANCode, cDatname, ltranspack
                
                sSQL = "select * into " & cDatname & " in '" & Pfad & cDatname & ".mdb' from FILZ"
                sSQL = sSQL & " where filiale = " & iFil & " and transpack = " & ltranspack
                gdBase.Execute sSQL, dbFailOnError
                
                cQuelle = Pfad & cDatname & ".mdb"
                cZiel = cPfad & "SIC\" & cDatname & ".mdb"
                lRet = CopyFile(cQuelle, cZiel, lfail)
                                
            End If
        rsFB.MoveNext
        Loop
    End If
    rsFB.Close
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        MsgBox "Es wurde keine Datei zum Speichern erzeugt.", vbInformation, "Winkiss Hinweis:"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SpeichernFBZproKiste"
        Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Function ermittlenextNumberZ(iFil As Integer, ltranspack As Long) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rec         As Recordset
    Dim aktNr       As Long
    Dim caktnr      As String
    Dim cFil        As String
    Dim cQuellfil   As String
    
    ermittlenextNumberZ = ""
    
    If Not NewTableSuchenDBKombi("WLFNR", gdBase) Then
        CreateTableT2 "WLFNR", gdBase
    End If
    
    If Not NewTableSuchenDBKombi("TRANS", gdBase) Then
        CreateTableT2 "TRANS", gdBase
        
        cSQL = "Insert into TRANS select transpack from WLFNR"
        gdBase.Execute cSQL, dbFailOnError
        
    End If

    Set rec = gdBase.OpenRecordset("WLFNR")
    rec.AddNew
    rec!lfnr = aktNr
    rec!Datum = Date
    rec!zeit = Time
    rec!FILIALE = iFil
    rec!TRANSPACK = ltranspack
        
    rec.Update
    rec.Close
    
    cSQL = "Delete from TRANS where  transpack = " & ltranspack
    gdBase.Execute cSQL, dbFailOnError
    
    Set rec = gdBase.OpenRecordset("TRANS")
    rec.AddNew
    rec!TRANSPACK = ltranspack
    rec.Update
    rec.Close
    
    'Aufbau Dateiname W ZF QF lfn 0 P
    
    cFil = iFil
    If Len(Trim$(cFil)) = 1 Then
        cFil = "0" & cFil
    End If
    
    cQuellfil = gcFilNr
    If Len(Trim$(cQuellfil)) = 1 Then
        cQuellfil = "0" & cQuellfil
    End If
    
    caktnr = CStr(ltranspack)
    ermittlenextNumberZ = "N" & cFil & cQuellfil & caktnr
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlenextNumberZ"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Private Sub UpdateKBLATTBCODE(cBcode As String, cDatname As String, dTranspack)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "update KBLATT  "
    sSQL = sSQL & " set KBLATT.DATNAME =  '" & cDatname & "' "
    sSQL = sSQL & " , KBLATT.BARCODE =  '" & cBcode & "' "
    sSQL = sSQL & " WHERE Transpack = " & dTranspack
    gdBase.Execute sSQL
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKBLATTBCODE"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub uebertragenVerteilZ(ltranspack As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rec         As Recordset
    Dim iRet        As Integer
    Dim cPfad       As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WVOUT"
    
    SpeichernFBZproKiste cPfad & "\", ltranspack
    
    InUnterwegsZ ltranspack

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "uebertragenVerteilZ"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InUnterwegsZ(ltranspack As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String

    Screen.MousePointer = 11
    
    If NewTableSuchenDBKombi("UNTERW", gdBase) = False Then
        CreateTableT2 "UNTERW", gdBase
        
        cSQL = "Create Index ARTNR on UNTERW (ARTNR)"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Insert into UNTERW Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", BESTVOR "
    cSQL = cSQL & ", Filiale "
    cSQL = cSQL & ", Filvon "
    cSQL = cSQL & ", ADATE  "
    cSQL = cSQL & ", AZEIT  "
    cSQL = cSQL & ", BEDNU  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", BESTiF  "
    cSQL = cSQL & ", OPENART "
    cSQL = cSQL & ", AENART  "
    cSQL = cSQL & ", Pcname "
    cSQL = cSQL & ", BEDNAME  "
    cSQL = cSQL & ", TRANSPACK  "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", FARBTEXT "
    cSQL = cSQL & ", FARBwert  "
    cSQL = cSQL & ", FARBwertS  "
    cSQL = cSQL & ", FARBNR  "
    cSQL = cSQL & ", false as sendok  "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as UDATE ,'" & TimeValue(Now) & "' as UZEIT "
    cSQL = cSQL & ", " & gcBedienerNr & " as UBEDNU  from FILZ"
    cSQL = cSQL & " where transpack = " & ltranspack
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from FILZ"
    cSQL = cSQL & " where transpack = " & ltranspack
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "InUnterwegsZ"
    Fehler.gsFehlertext = "Im Programmteil Warenverteilung aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub KBausblenden(ltranspack As Double)
On Error GoTo LOKAL_ERROR

Dim cSQL        As String
Dim rsrs        As Recordset
Dim rsrs1       As Recordset
Dim rsRs2       As Recordset
Dim cArtNr      As String
Dim cFilnr      As String

Screen.MousePointer = 11

cSQL = "Select * from FILZ"
cSQL = cSQL & " where transpack = " & ltranspack
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            cFilnr = rsrs!FILIALE
        End If
        
'        If KundenbestArtikelforfilAusblend(rsrs!artnr, CLng(cFilNr), ltranspack) Then
'
'        End If

    rsrs.MoveNext
    Loop
End If
rsrs.Close

cSQL = "Select distinct(Filiale) as Mfil from FILZ"
cSQL = cSQL & " where transpack = " & ltranspack
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst

    If Not IsNull(rsrs!Mfil) Then
        cFilnr = rsrs!Mfil
    End If
    
    cSQL = "Select bestelltmenge from KUNDBEST where "
    cSQL = cSQL & "   StatusARTIKEL = 'NICHTGELIEFERT' "
    cSQL = cSQL & " and  Filiale = " & cFilnr
    cSQL = cSQL & " and  sendok = false"
    
    Set rsrs1 = gdBase.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        cSQL = "Insert into KP" & srechnertab & " Select KUNDBEST.*,  0 as transpack from KUNDBEST  "
        cSQL = cSQL & " where  StatusARTIKEL = 'NICHTGELIEFERT' "
        cSQL = cSQL & " and  Filiale = " & cFilnr
        cSQL = cSQL & " and  sendok = false"
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Update KUNDBEST set sendok = true  where "
        cSQL = cSQL & " StatusARTIKEL = 'NICHTGELIEFERT' "
        cSQL = cSQL & " and  Filiale = " & cFilnr
        gdBase.Execute cSQL, dbFailOnError
        
    Else
        cSQL = "Select bestelltmenge from KP" & srechnertab & " where artnr = 0"
        cSQL = cSQL & " and  Filiale = " & cFilnr
        Set rsRs2 = gdBase.OpenRecordset(cSQL)
        If rsRs2.EOF Then
            cSQL = "Insert into KP" & srechnertab & " (transpack,Filiale,artnr) values  (0," & cFilnr & ",0 )"
            gdBase.Execute cSQL, dbFailOnError
        End If
        rsRs2.Close
        
    End If
    rsrs1.Close
End If
rsrs.Close

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KBausblenden"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertKBLATT(lTRANSP As Long, lTournr As Long, lFiliale As Long, lFilVon As Long, lMenge As Long)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Insert into KBLATT (Transpack,Tournr,Filiale,Filvon,Menge)"
    sSQL = sSQL & " values "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " " & lTRANSP & " "
    sSQL = sSQL & " ," & lTournr & " "
    sSQL = sSQL & " ," & lFiliale & " "
    sSQL = sSQL & " ," & lFilVon & " "
    sSQL = sSQL & " ," & lMenge & " "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertKBLATT"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateKBLATT()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lNr As Long
    Dim rsFil As Recordset
    Dim i As Integer
    
'    sSQL = "update KBLATT inner join Tour on KBLATT.TOURNR = TOUR.TOURNR "
'    sSQL = sSQL & " set KBLATT.TOURBEZ =  TOUR.TOURBEZ "
'    sSQL = sSQL & " , KBLATT.TOURBEM =  TOUR.TOURBEM "
'    gdBase.Execute sSQL
    
    Set rsFil = gdBase.OpenRecordset("select * from Filialen order by filialnr")
    
    If Not rsFil.EOF Then
        rsFil.MoveFirst
        Do While Not rsFil.EOF
           
            If Not IsNull(rsFil!FILIALNR) Then
                i = rsFil!FILIALNR
                
                sSQL = "select * from KBLATT where filiale = " & i
                Set rsrs = gdBase.OpenRecordset(sSQL)
                
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    lNr = 0
                    Do While Not rsrs.EOF
                        lNr = lNr + 1
                        rsrs.Edit
                        rsrs!lfnr = lNr
                        rsrs.Update
                        
                        rsrs.MoveNext
                    Loop
                    
                    rsrs.AddNew
                    lNr = lNr + 1
                    rsrs!lfnr = lNr
                    rsrs!FILIALE = i
                    rsrs.Update
                    
                    rsrs.AddNew
                    lNr = lNr + 1
                    rsrs!FILIALE = i
                    rsrs!lfnr = lNr
                    rsrs.Update
                    
                    rsrs.AddNew
                    lNr = lNr + 1
                    rsrs!FILIALE = i
                    rsrs!lfnr = lNr
                    rsrs.Update
                
                End If
                rsrs.Close

            End If
            rsFil.MoveNext
        Loop
    End If
    rsFil.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKBLATT"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function freieTranspack() As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL            As String
    Dim rec             As Recordset
    
    Screen.MousePointer = 11
    
    freieTranspack = -1
    
    If Not NewTableSuchenDBKombi("TRANS", gdBase) Then
        CreateTableT2 "TRANS", gdBase
        
        If Not NewTableSuchenDBKombi("WLFNR", gdBase) Then
            CreateTableT2 "WLFNR", gdBase
        End If
        
        cSQL = "Insert into TRANS select transpack from WLFNR"
        gdBase.Execute cSQL, dbFailOnError
        
        CheckIndex "TRANS", "TRANSPACK", "", gdBase
    End If
    
    cSQL = "Select max(transpack) as maxi from TRANS "
    Set rec = gdBase.OpenRecordset(cSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!maxi) Then
            freieTranspack = rec!maxi
        End If
    End If
    rec.Close
    
    If freieTranspack < 100000 Then
        freieTranspack = freieTranspack + 1
    End If
    
    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "freieTranspack"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub SpeichernFB(Pfad As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cQuelle     As String
    Dim cZiel    As String
    Dim sSQL        As String
    Dim db          As Database
    Dim rsFB        As Recordset
    Dim iFil        As Integer
    Dim cDatname    As String
    Dim cDatum      As String
    Dim czeit       As String
    ReDim cZeilen(0 To 6) As String
    Dim cPfad1 As String
    Dim bytefil As Byte
    Dim lfail As Long
    Dim lRet As Long
    
    
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    sSQL = "select distinct(filiale)as fili from anfilm "
    Set rsFB = gdBase.OpenRecordset(sSQL)
    If Not rsFB.EOF Then
        rsFB.MoveFirst
        Do While Not rsFB.EOF
        
            If Not IsNull(rsFB!fili) Then
                iFil = rsFB!fili
                cDatname = ermittlenextNumber(iFil)
                'Drucke den Beleg
    
                cZeilen(0) = "Filialtausch-Beleg"
                cZeilen(1) = "-----------------"
                cZeilen(2) = "an Fil:   " & iFil
                cZeilen(3) = "Dateiname:" & cDatname
                cZeilen(4) = ""
                cZeilen(5) = "Datum: " & cDatum
                cZeilen(6) = "Zeit:  " & czeit
                
                DruckeArbeitszeitBelegWK20d cZeilen(), 6
                
                If Not FindFile(Pfad, cDatname & ".MDB") Then
                    Set db = CreateDatabase(Pfad & cDatname & ".mdb", dbLangGeneral, dbVersion40)
                    db.Close
                End If
                
                sSQL = "select * into " & cDatname & " in '" & Pfad & cDatname & ".mdb' from anfilm"
                sSQL = sSQL & " where filiale = " & iFil
                gdBase.Execute sSQL, dbFailOnError
                
                cPfad1 = gcDBPfad
                If Right(cPfad1, 1) <> "\" Then
                    cPfad1 = cPfad1 & "\"
                End If
                cPfad1 = cPfad1 & "WVSIC\"
                
                cQuelle = Pfad & cDatname & ".mdb"
                cZiel = cPfad1 & cDatname & ".mdb"
                lRet = CopyFile(cQuelle, cZiel, lfail)
                
                If Text2.Text <> "" And Text3.Text <> "" Then
                    bytefil = CByte(Trim(Text3.Text))
                    If iFil = bytefil Then
                    cPfad1 = Text2.Text
                    If Right(cPfad1, 1) <> "\" Then
                        cPfad1 = cPfad1 & "\"
                    End If
            
                    cQuelle = Pfad & cDatname & ".mdb"
                    cZiel = cPfad1 & cDatname & ".mdb"
                    lRet = CopyFile(cQuelle, cZiel, lfail)
                    
                    If lRet = 1 Then
                        Kill cQuelle
                    End If
                    
                
                    End If
            
                End If
                            
                sSQL = "delete from anfilm"
                sSQL = sSQL & " where filiale = " & iFil
                gdBase.Execute sSQL, dbFailOnError
                
            End If
            
            
        rsFB.MoveNext
        Loop
    End If
    rsFB.Close
    

        
Exit Sub
LOKAL_ERROR:
    If err.Number = 3044 Or err.Number = 3043 Then
        If MsgBox("Laufwerk A nicht bereit", vbRetryCancel + vbInformation, "Hinweis") = vbRetry Then
            Resume
        Else
            Exit Sub
        End If
    ElseIf err.Number = 3051 Then
        If MsgBox("Laufwerk A ist nicht bereit!", vbInformation + vbRetryCancel, "Diskettenspeicherung") = vbRetry Then
            Resume
        Else
            Exit Sub
        End If
    End If
    If err.Number = 53 Then
        MsgBox "Es wurde keine Datei zum Speichern erzeugt", vbInformation, "Winkiss Hinweis:"
    ElseIf err.Number = 20530 Then
        MsgBox "Sie haben keine Diskette eingelegt", vbInformation, "Winkiss Hinweis:"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SpeichernFB"
        Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Function ermittlenextNumber(iFil As Integer) As String
On Error GoTo LOCALERROR

    Dim cSQL As String
    Dim rec As Recordset
    Dim aktNr As Long
    Dim caktnr As String
    Dim cFil As String
    Dim cFromFil As String
    
    If Len(Trim(gcFilNr)) = 1 Then
        cFromFil = Trim(gcFilNr)
    Else
        cFromFil = "0"
    End If
    
    ermittlenextNumber = ""
    
    If Not NewTableSuchenDBKombi("FWVLFNR", gdBase) Then
        CreateTable "FWVLFNR", gdBase
    End If
    
    cSQL = "select max(lfnr) as aktlfnr from FWVLFNR where filiale = " & iFil
    Set rec = gdBase.OpenRecordset(cSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!aktlfnr) Then
            aktNr = rec!aktlfnr
        Else
            aktNr = 0
        End If
    Else
        aktNr = 0
    End If
    rec.Close: Set rec = Nothing
    
    aktNr = aktNr + 1
    
    If aktNr > 999 Then
        cSQL = "Delete from FWVLFNR where filiale = " & iFil
        gdBase.Execute cSQL, dbFailOnError
        aktNr = 1
    End If
    
    Set rec = gdBase.OpenRecordset("FWVLFNR")
    rec.AddNew
    rec!lfnr = aktNr
    rec!Datum = Date
    rec!zeit = Time
    rec!FILIALE = iFil
        
    rec.Update
    rec.Close: Set rec = Nothing
    
    cFil = iFil
    If Len(Trim(cFil)) = 1 Then
    cFil = "0" & cFil
    End If
    
    caktnr = aktNr
    
    If Len(Trim(caktnr)) = 1 Then
        caktnr = cFromFil & "00" & caktnr
    ElseIf Len(Trim(caktnr)) = 2 Then
        caktnr = cFromFil & "0" & caktnr
    ElseIf Len(Trim(caktnr)) = 3 Then
        caktnr = cFromFil & caktnr
    Else
        caktnr = cFromFil & "001"
    End If
    
    ermittlenextNumber = "WV" & cFil & caktnr
    
Exit Function
LOCALERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlenextNumber"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Private Sub uebertragenVerteil()
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rec         As Recordset
    Dim cPfad       As String
    Dim iRet        As Integer
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    

    SpeichernFB cPfad & "WV\"
    
    giKissFtpMode = 14
    frmWKL38.Show 1
    
    iRet = MsgBox("Übertragung erfolgreich? - dann Artikel aus dem Speicher löschen?", vbQuestion + vbYesNoCancel, "Winkiss Frage:")
    If iRet = vbYes Then
        loeschNEW "ANFILM", gdBase
        CreateTable "ANFILM", gdBase
        Label4.Caption = "Artikelanzahl: " & ermittleAnzVerteil
        Label4.Refresh
        List2.Clear
    End If

    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "uebertragenVerteil"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseProtokollFilialTauschWKLaj()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim dSumMenge   As Double
    Dim dSumEkpr    As Double
    Dim dSumLekpr   As Double
    Dim dSumkvkpr1  As Double
    Dim bAnd        As Boolean
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim lLinr       As Long
    Dim lbed        As Long
    Dim lartnr      As Long
    Dim lZfil       As Long
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    dSumMenge = 0
    dSumEkpr = 0
    dSumLekpr = 0
    dSumkvkpr1 = 0
    
    bAnd = False
    
    loeschNEW "TAUZ", gdBase
    CreateTableT2 "TAUZ", gdBase
    
    cSQL = "Insert into TAUZ Select "
    cSQL = cSQL & " SENDOK "
    cSQL = cSQL & ", ADATE "
    cSQL = cSQL & ", AZEIT "
    cSQL = cSQL & ", MENGE "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", FIL_VON "
    cSQL = cSQL & ", FIL_AN "
    cSQL = cSQL & ", KASNUM "
    cSQL = cSQL & ", EKPR "
    cSQL = cSQL & ", LEKPR "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", BEDIENER "
    cSQL = cSQL & " from TAUSCH "
    
    cFeld = Text1(0).Text
    If cFeld <> "" Then
        lDatVon = DateValue(cFeld)
    Else
        lDatVon = -1
    End If
        
    cFeld = Text1(2).Text
    If cFeld <> "" Then
        lDatBis = DateValue(cFeld)
    Else
        lDatBis = -1
    End If
    
    If lDatVon <> -1 Or lDatBis <> -1 Then
        If lDatVon = -1 Then
            lDatVon = lDatBis
        End If
        If lDatBis = -1 Then
            lDatBis = lDatVon
        End If
        
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " ADATE >= " & Trim$(Str$(lDatVon)) & " and ADATE <= " & Trim$(Str$(lDatBis)) & " "
        bAnd = True
    End If
    
    cFeld = Text1(4).Text
    If cFeld <> "" Then
        lLinr = Val(cFeld)
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " LINR = " & Trim$(Str$(lLinr)) & " "
        bAnd = True
    End If
    
    cFeld = Text1(3).Text
    If cFeld <> "" Then
        lartnr = Val(cFeld)
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " ARTNR = " & Trim$(Str$(lartnr)) & " "
        bAnd = True
    End If
    
    cFeld = Text1(1).Text
    If cFeld <> "" Then
        lZfil = Val(cFeld)
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " FIL_AN = " & Trim$(Str$(lZfil)) & " "
        bAnd = True
    End If
    
    cFeld = Text1(5).Text
    If cFeld <> "" Then
        lbed = Val(cFeld)
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " BEDIENER = " & Trim$(Str$(lbed)) & " "
        bAnd = True
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    Dim cArtNr As String
    Dim czeit As String
    Dim cDatum As String
    
    
    
    List1.Clear
    List3.Clear
    List3.AddItem "  Datum    Zeit     Bed Menge ArtNr. Artikelbezeichnung                  LiefNr Linie AN           SEK        LEK        KVK "

    cSQL = " Select * from Tauz order by ADATE desc, AZEIT desc , ARTNR "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            'schon versendet?
            'Bist du in Filz, dann bist du noch nicht versendet
            
            cArtNr = ""
            czeit = ""
            cDatum = ""
    
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                cDatum = Format$(rsrs!ADATE, "DD.MM.YY")
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = rsrs!AZEIT
            End If
            
            
            If inFilz(cArtNr, cDatum, czeit) = True Then
                cLBSatz = "  "
            Else
                cLBSatz = "V "
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                dWert = rsrs!ADATE
                cFeld = Format$(dWert, "DD.MM.YY")
            Else
                cFeld = ""
            End If
            cFeld = Space$(8 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!AZEIT) Then
                cFeld = rsrs!AZEIT
            Else
                cFeld = 0
            End If
            cFeld = Format$(cFeld, "HH:MM:SS")
            cFeld = Space$(8 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!BEDIENER) Then
                dWert = rsrs!BEDIENER
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "##0")
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!Menge) Then
                dWert = rsrs!Menge
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "####0")
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!artnr) Then
                dWert = rsrs!artnr
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "#####0")
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!linr) Then
                dWert = rsrs!linr
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "#####0")
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!LPZ) Then
                dWert = rsrs!LPZ
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "#0")
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!FIL_AN) Then
                dWert = rsrs!FIL_AN
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "#0")
            cFeld = Space$(4 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            cLBSatz = cLBSatz & Space$(4)
            
            If Not IsNull(rsrs!ekpr) Then
                dWert = rsrs!ekpr
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "######0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!lekpr) Then
                dWert = rsrs!lekpr
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "######0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            cFeld = Format(dWert, "######0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            List1.AddItem cLBSatz
            rsrs.MoveNext
        Loop
       

    End If
    
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseProtokollFilialTauschWKLaj"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad   As String

    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Pfad speichern"
'        .Filter = "Access - Dateien (*.mdb)|*.mdb"
        .FileName = "alle Dateien"
        .ShowSave
    End With
    
    
    
    sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
    Text2.Text = sPfad
    
    
        
err:
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR

    Frame4.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click()
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    zeigeHilfe "LPROTOK", "FilProt.txt", cPfad

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    positionierenwklaj
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label1
    
    Frame4.BackColor = glH2
    Check1.BackColor = glH2
    Check2.BackColor = glH2
    Check3.BackColor = glH2
    Check4.BackColor = glH2


    If SpalteInTabellegefundenNEW("EAJ", "BO2", gdBase) = False Then
       SpalteAnfuegenNEW "EAJ", "BO2", "bit", gdBase
       SpalteAnfuegenNEW "EAJ", "BO3", "bit", gdBase
       
       Dim sSQL As String
       sSQL = "Update EAJ set Bo2 = -1 "
       gdBase.Execute sSQL, dbFailOnError
       
       sSQL = "Update EAJ set Bo3 = -1 "
       gdBase.Execute sSQL, dbFailOnError
       
    End If
    
    voreinstellungladen

    If Not NewTableSuchenDBKombi("ANFILM", gdBase) Then
        CreateTable ("ANFILM"), gdBase
    End If
    
    If Not NewTableSuchenDBKombi("FILZ", gdBase) Then
        CreateTableT2 ("FILZ"), gdBase
    End If

    
    
    Text1(0).Text = Format(DateValue(Now) - 3, "DD.MM.YYYY")
    Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    Command1_Click 2
        
    Call SendMessage(List1.hwnd, LB_SETHORIZONTALEXTENT, 1000, 0&)
    
    If Check1.Value = vbChecked Then
    
        Command1(9).Visible = True
    
        If gbFtpYes Then
            Command1(9).Enabled = True
        End If
        
        Label5.Visible = True
        Command1(8).Visible = True
        Label12.Visible = True
    Else
        Command1(9).Visible = False
        Label5.Visible = False
        Command1(8).Visible = False
        Label12.Visible = False
    End If

    fülleliste List4
    
    LeseFilialenWKLaj
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Filialtausch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleliste(Listx As ListBox)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset
Dim sSatz As String
Dim lCount1 As Long
Dim lCount2 As Long


Listx.Clear



sSQL = "Select distinct(TRANSPACK) as KISTENNR ,Filiale,max(pcname)as Maxname,max(azeit)as Maxitime,sum(BESTVOR) as Maximenge,Openart from Filz "


sSQL = sSQL & " group by TRANSPACK ,Filiale,Openart"
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    
    rsrs.MoveFirst
    Do While Not rsrs.EOF
    
    If Not IsNull(rsrs!KISTENNR) Then
        sSatz = rsrs!KISTENNR & Space(15 - Len(rsrs!KISTENNR))
    End If
    
    If Not IsNull(rsrs!FILIALE) Then
        sSatz = sSatz & rsrs!FILIALE & Space(8 - Len(rsrs!FILIALE))
    Else
        sSatz = sSatz & Space(8)
    End If
    
    If Not IsNull(rsrs!Maxitime) Then
        sSatz = sSatz & rsrs!Maxitime & Space(20 - Len(rsrs!Maxitime))
    Else
        sSatz = sSatz & Space(20)
    End If
    
    If Not IsNull(rsrs!Maximenge) Then
        sSatz = sSatz & rsrs!Maximenge & Space(7 - Len(rsrs!Maximenge))
    Else
        sSatz = sSatz & Space(7)
    End If
    
    If Not IsNull(rsrs!openart) Then
        sSatz = sSatz & rsrs!openart & Space(12 - Len(rsrs!openart))
    Else
        sSatz = sSatz & Space(12)
    End If
   
    Listx.AddItem sSatz
    
    rsrs.MoveNext
    Loop
End If
rsrs.Close

If Listx.ListCount > 0 Then
    Command1(9).ForeColor = vbRed
Else
    Command1(9).ForeColor = Command1(0).ForeColor
End If


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fülleliste"
    Fehler.gsFehlertext = "Im Programmteil Transportverpackung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function inFilz(sArtnr As String, sDatum As String, sZeit As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim lDate   As Long
    
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    If sDatum = "" Then
        Exit Function
    End If
    
    If sZeit = "" Then
        Exit Function
    End If
    
    inFilz = False
    
    lDate = DateValue(sDatum)
    
    cSQL = "Select * from Filz where artnr = " & sArtnr & " "
    cSQL = cSQL & " and Adate = " & lDate & " "
    cSQL = cSQL & " and Azeit = '" & sZeit & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        inFilz = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "inFilz"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function ermittleAnzVerteil() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    ermittleAnzVerteil = "0"
    
    cSQL = "Select * from ANFILM"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        ermittleAnzVerteil = rsrs.RecordCount
    Else
        ermittleAnzVerteil = "0"
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleAnzVerteil"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub fuellelist()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    cSQL = "Select * from ANFILM order by lfnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    List2.Clear
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            If Not IsNull(rsrs!BESTVOR) Then
                cFeld = rsrs!BESTVOR
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
            
            If Not IsNull(rsrs!FILIALE) Then
                cFeld = rsrs!FILIALE
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(9 - Len(cFeld))
            
            
            If Not IsNull(rsrs!LASTDATE) Then
                cFeld = rsrs!LASTDATE
            Else
                cFeld = ""
            End If
            
            cFeld = Format$(cFeld, "DD.MM.YY")
            
            cLBSatz = cLBSatz & cFeld & Space(10 - Len(cFeld))
            
            
            If Not IsNull(rsrs!LASTTIME) Then
                cFeld = rsrs!LASTTIME
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld
            
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    List1.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelist"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub positionierenwklaj()
    On Error GoTo LOKAL_ERROR
    
    With Frame1
        .BorderStyle = 0
    End With
    
    With Frame2
        .BorderStyle = 0
    End With
    
    With Frame3
        .BorderStyle = 0
    End With
    
    Frame4.Height = 2775
    Frame4.Left = 5400
    Frame4.Top = 840
    Frame4.Width = 6375
    Frame4.BorderStyle = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwklaj"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub SendeDaten2DruckerNeuWKLaj()
    On Error GoTo LOKAL_ERROR

    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    Dim iLevel As Integer
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim lAnzZeile As Long
    Dim iAktCopy As Integer
    Dim cDaten As String
    Dim iLenZeile As Integer
    Dim dSumme As Double
    Dim dMWStVoll As Double
    Dim dMWStErm As Double
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim iFileNr As Integer
    Dim cText As String
    Dim cLBSatz As String
    Dim cFeld As String
    Dim cMwst As String
    Dim ctmp As String
    Dim dWert As Double
    Dim dMWSt As Double
    Dim dEuro As Double
    Dim lAnzArt As Long
    Dim lHeute As Long
    
    Dim sLEKPR      As String
    Dim dLEK        As Double
    Dim dsumlek     As Double
    Dim lanz1       As Long
    Dim cAn         As String
    Dim cTauschnr   As String
    
    'Bon-Drucker einschalten
    
    setzedrucker gcBonDrucker
    
    iLevel = 0
    
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz

StartPunkt:
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1
    iLevel = 1
    cDaten = ""
    iLenZeile = 32
    dSumme = 0
    dMWStVoll = 0
    dMWStErm = 0
    lAnzArt = 0
    
    '***********************************************
    'Hier geht's los
    '***********************************************
    
    lAnzSatz = frmWKL20!List1.ListCount
    
    iLevel = 2

    '***********************************************
    'Drucker ein- und Kundendisplay ausschalten
    '***********************************************
    
    cEscapeSequenz = gcInit
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'ggf. vorhandenes Logo auf Kassenbon bringen
    '***********************************************
    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Kopfdaten 1.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 2.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 3.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(4)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'KENNZEICHNUNG FILIALTAUSCH
    '***********************************************
    
    cDaten = "F I L I A L T A U S C H"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        

    '***********************************************
    'KENNZIFFERN FILIALEN
    '***********************************************
    
    cAn = Trim$(Str$(Val(Left(Label2.Caption, 2))))
    
    cDaten = "VON: " & gcFilNr & "     AN: " & cAn
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'eventuell Tauschnr
    '***********************************************
    
    cTauschnr = Trim$(Str$(Val(Label12.Caption)))
    
    cDaten = "TauschNr: " & cTauschnr
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Artikelpositionen drucken
    '***********************************************
    
    iLevel = 4
    
    dSumme = 0
    
    
    
    lanz1 = 0
    dsumlek = 0
    dLEK = 0
    sLEKPR = ""
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        cFeld = Mid(cLBSatz, 7, 6)
        
        sLEKPR = ermPREIS(Trim(cFeld), "LEKPR")
        sLEKPR = Format$(sLEKPR, "#####0.00 ")
        
        
        
        lanz1 = CLng(Mid(cLBSatz, 1, 5))
        dLEK = lanz1 * sLEKPR
        sLEKPR = Format$(dLEK, "#####0.00 ")
        dsumlek = dsumlek + dLEK

        
        If cFeld <> "000000" Then
            '1.Zeile: ArtNr + MWSTKz + ArtBezeich
            cDaten = cFeld & " "
            
            cFeld = Mid(cLBSatz, 72, 1)
            cDaten = cDaten & cFeld & "  "
            cMwst = cFeld
            
            cFeld = Mid(cLBSatz, 14, 35)
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 17 Then
                cFeld = Left(cFeld, 17)
            End If
            cDaten = cDaten & cFeld
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***********************************************
            'EKpreis drucken
            '***********************************************
            
            If gbFILMEK = True Then
                cDaten = "EK - Wert:              "
                ctmp = sLEKPR
                
                ctmp = Space(9 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
                
            '***********************************************
            'wenn Artikelermäßigung, dann drucken
            '***********************************************
            
            ctmp = Mid(cLBSatz, 124, 3)
            If Val(ctmp) > 0 And gbRabatt Then
                'Zeile nur bei Artikel-Ermäßigung drucken
                
                Dim dArtikelrabattinEuro As Double
                dArtikelrabattinEuro = CDbl(Trim(Mid(cLBSatz, 84, 9)))
                Dim dRabattierterGesamtPreisinEuro As Double
                dRabattierterGesamtPreisinEuro = CDbl(Trim(Mid(cLBSatz, 60, 9)))
                Dim dErgebnisinProz As Double
                dErgebnisinProz = dArtikelrabattinEuro * 100 / (dRabattierterGesamtPreisinEuro + dArtikelrabattinEuro)
                ctmp = Format$(dErgebnisinProz, "###,##0.00")
                
                
                cDaten = "Rabatt:    " & ctmp & " %"
                ctmp = Mid(cLBSatz, 84, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                ctmp = Space(9 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
            End If
            
            '***********************************************
            'Anzahl, Einzelpreis, Positionspreis drucken
            '***********************************************
            
            ctmp = Mid(cLBSatz, 1, 5)
            ctmp = Trim$(ctmp)
            lAnzArt = lAnzArt + Val(ctmp)
            ctmp = ctmp & Space$(6 - Len(ctmp))
            cDaten = ctmp & " x"
            
            ctmp = Mid(cLBSatz, 74, 9)
            ctmp = fnMoveComma2Point$(ctmp)
            dWert = Val(ctmp)
            If Left(gFirma.FirmaName, 5) = "Stief" Then
                ctmp = Format$((dWert * 100), "########0")
            Else
                ctmp = Format$(dWert, "#####0.00")
            End If
            ctmp = Space(11 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = fnMoveComma2Point$(ctmp)
            dWert = Val(ctmp)
            ctmp = Format$(dWert, "#####0.00")
            dSumme = dSumme + dWert
            ctmp = Space(13 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            '***********************************************
            'MWSt-Summe berechnen
            '***********************************************
    
            
            If cMwst = "V" Then
                dMWSt = dWert / (100 + gdMWStV)
                dMWSt = dMWSt * gdMWStV
                dMWStVoll = dMWStVoll + dMWSt
            ElseIf cMwst = "E" Then
               dMWSt = dWert / (100 + gdMWStE)
                dMWSt = dMWSt * gdMWStE
                dMWStErm = dMWStErm + dMWSt
            Else
                dMWSt = 0
            End If
        Else
        
'            'Zeile mit Zwischensumme drucken
'            cDaten = "Zwischensumme:     "
'
'            ctmp = Mid(cLBSatz, 60, 9)
'            ctmp = fnMoveComma2Point$(ctmp)
'            dWert = Val(ctmp)
'            ctmp = Format$(dWert, "#####0.00")
'            ctmp = Space(13 - Len(ctmp)) & ctmp
'
'
'
'
'
'            cDaten = cDaten & ctmp

            'Zeile mit Zwischensumme drucken
            ctmp = Mid(cLBSatz, 13, Len(cLBSatz) - 13)
            ctmp = Left(Trim(ctmp), 32)
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        End If
    Next lAktSatz

    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    iLevel = 5
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Anzahl Artikel drucken
    '***********************************************
    
    iLevel = 5
    
    ctmp = "Anzahl Artikel:" & Space$(6)
    cDaten = ctmp
    ctmp = Format$(lAnzArt, "########0")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Summenzeile drucken
    '***********************************************
   
    iLevel = 6

    
    ctmp = "Summe:" & Space$(12) & gcWaehrung
    cDaten = ctmp
    ctmp = Format$(dSumme, "#####0.00")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Summenzeile drucken
    '***********************************************
   
    iLevel = 6

    If gbFILMEK = True Then
    
        ctmp = "Summe EK:" & Space$(9) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dsumlek, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    '***********************************************
    'Zeile volle MWSt drucken
    '***********************************************
    If dMWStVoll <> 0 Then
    
        ctmp = "MWSt.-Anteil: " & Format$(gdMWStV, "#0") & "%" & Space$(1) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dMWStVoll, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    End If
    '***********************************************
    'Zeile erm. MWSt drucken
    '***********************************************
    If dMWStErm <> 0 Then
    
        ctmp = "MWSt.-Anteil: " & Format$(gdMWStE, "#0") & "%" & Space$(2) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dMWStErm, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    End If
    '***********************************************
    'Zeile 'Es bediente Sie' drucken
    '***********************************************
    
    ctmp = "Es bediente Sie"
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Bedienername drucken
    '***********************************************
    
    ctmp = gcBediener
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Zeile Datum, BelegNr, Uhrzeit drucken
    '***********************************************
    
    ctmp = Format$(Date, "DD.MM.YYYY")
    cDaten = ctmp
    ctmp = "0"
    ctmp = gcKasNum & "/" & ctmp
    ctmp = Space$(6 - Len(ctmp)) & ctmp
    cDaten = cDaten & Space$(5) & ctmp
    ctmp = Format$(Now, "HH:MM")
    cDaten = cDaten & Space$(6) & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, gsSTERNZEICH)
'    cDaten = String$(iLenZeile, "*")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
  
    '***********************************************
    'Fußzeile 1 drucken
    '***********************************************
    
    'Fußzeilen
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Fußzeile 2 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Fußzeile 3 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(5)
    End If
    
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    
    iLevel = 10
    
    '***********************************************
    'ein paar Leerzeilen drucken
    '***********************************************
    If gbNoGrafik = True Then
        For lcount = 1 To 9
            If lcount = 9 Then
                cEscapeSequenz = "." & vbCrLf
            Else
                cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
    End If

    'OpenDrawer3 benutzt die WindowsAPI
    'OpenDrawer4 geht über das PRINTER-Objekt
    
    iLevel = 12
    
    
BON_DRUCKEN:
    If gbBONNEIN = False Then
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
    End If
    If iAktCopy = 1 Then
    
        If gbBONNEIN = True Then
            For lcount = 1 To 9
                If lcount = 9 Then
                    cEscapeSequenz = "." & vbCrLf
                Else
                    cEscapeSequenz = " " & vbCrLf
                End If
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Next lcount
        End If
        SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False, True
    End If

BON_SCHNEIDEN:


    'Kassenbon abschneiden
    If gbBONNEIN = False Then
        If gbAPI = True Then
        
            If gbNoGrafik = False Then
                lAnzZeile = 0
        
                ReDim cDruckZeile(1 To 1) As String
                
                cDaten = "AN: " & cAn
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            
                OpenDrawer4Groß aDeviceName, cDruckZeile(), lAnzZeile
            Else
                'Kassenbon abschneiden
                If gbAPI = True Then
                    aDeviceName = Printer.DeviceName
                    cEscapeSequenz = gcSchneiden
                    OpenDrawer aDeviceName, cEscapeSequenz
                End If
            End If
    

        End If
    End If
    iLevel = 11
    
ZWEITER_BON:
    If gb2BONFI = True Then
        If iAktCopy < 2 Then
            GoTo StartPunkt
        End If
    End If
    
    Erase cDruckZeile
    
ENDE:

    frmWKL20!Label8(0).Caption = ""
    frmWKL20!Label8(1).Caption = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerNeuWKLaj"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR

    LogtoEnd Me
    voreinstellungspeichern
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Dim bo0 As Integer
    Dim cPfad As String
    Dim bytefil As Byte
    Dim bo1 As Integer
    
    Dim bo2 As Integer
    Dim bo3 As Integer
    
    If Trim(Text2.Text) = "" Then
        cPfad = ""
    Else
        cPfad = Text2.Text
    End If
    
    bytefil = CByte(Val(Trim(Text3.Text)))

    sSQL = "delete from eaj "
    gdBase.Execute sSQL, dbFailOnError
    
    If Check1.Value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    If Check2.Value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    If Check3.Value = vbChecked Then
        bo2 = 0
    Else
        bo2 = -1
    End If
    
    If Check4.Value = vbChecked Then
        bo3 = 0
    Else
        bo3 = -1
    End If

    sSQL = "Insert into eaj ( bo0,Pfad,fil,bo1,bo2,bo3) values "
    sSQL = sSQL & "(" & bo0 & ", '" & cPfad & "' , " & bytefil & ", " & bo1 & ", " & bo2 & ", " & bo3 & " )"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("EAJ")
    If Not rs.EOF Then
    
        If rs!bo0 = True Then
            Check1.Value = vbUnchecked
        Else
            Check1.Value = vbChecked
        End If
        
        If Not IsNull(rs!Pfad) Then
            Text2.Text = rs!Pfad
        Else
            Text2.Text = ""
        End If
        
        If Not IsNull(rs!fil) Then
            Text3.Text = rs!fil
        Else
            Text3.Text = ""
        End If
        
        If Text3.Text = "0" Then
            Text3.Text = ""
        End If
        
        If rs!bo1 = True Then
            Check2.Value = vbUnchecked
        Else
            Check2.Value = vbChecked
        End If
        
        If rs!bo2 = True Then
            Check3.Value = vbUnchecked
        Else
            Check3.Value = vbChecked
        End If
        
        If rs!bo3 = True Then
            Check4.Value = vbUnchecked
        Else
            Check4.Value = vbChecked
        End If
    
    
    End If
    rs.Close: Set rs = Nothing

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Label0.Caption = Trim$(Str$(Index))
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Filialtausch/Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
