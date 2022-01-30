VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL13 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kundendaten"
   ClientHeight    =   8625
   ClientLeft      =   1245
   ClientTop       =   2415
   ClientWidth     =   11910
   Icon            =   "frmWKL13.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   1560
      TabIndex        =   283
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   720
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   8160
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   211
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   11280
      TabIndex        =   118
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picprogress 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   11595
      TabIndex        =   117
      Top             =   480
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   1200
      TabIndex        =   98
      Top             =   6240
      Visible         =   0   'False
      Width           =   3360
      Begin sevCommand3.Command Command41 
         Height          =   375
         Left            =   9600
         TabIndex        =   107
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   0
         TabIndex        =   101
         Top             =   480
         Width           =   9375
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   7
         Left            =   9600
         TabIndex        =   100
         Top             =   240
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
      Begin VB.ListBox List7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   0
         TabIndex        =   99
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Kunden-Umsatz"
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
         Index           =   0
         Left            =   9600
         TabIndex        =   103
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   102
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Kunden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   8175
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   9480
      Begin sevCommand3.Command Command2 
         Height          =   330
         Index           =   5
         Left            =   9120
         TabIndex        =   123
         Top             =   7200
         Width           =   375
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   4
         Left            =   8160
         TabIndex        =   120
         Top             =   7560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
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
         Caption         =   "Listen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   3
         Left            =   6240
         TabIndex        =   119
         Top             =   7560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         Caption         =   "Historie F3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   2
         Left            =   2280
         TabIndex        =   9
         Top             =   7560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   1
         Left            =   9840
         TabIndex        =   10
         Top             =   7560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   8
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   183
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11245
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
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   6
         Left            =   4200
         TabIndex        =   286
         Top             =   7560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
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
         Caption         =   "Datenschutz"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label11 
         Caption         =   "Stand:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   180
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label5 
         Appearance      =   0  '2D
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   122
         Tag             =   "Shape"
         Top             =   7200
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "grüne Kunden = offene Kundenbestellungen(noch nicht bestellt)"
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
         Index           =   33
         Left            =   480
         TabIndex        =   121
         Top             =   7200
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   84
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   217
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "!"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   1
         Left            =   1005
         TabIndex        =   218
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "´"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   2
         Left            =   1770
         TabIndex        =   219
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "§"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   3
         Left            =   2535
         TabIndex        =   220
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   4
         Left            =   3300
         TabIndex        =   221
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   5
         Left            =   4065
         TabIndex        =   222
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "&&"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   6
         Left            =   4830
         TabIndex        =   223
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "/"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   7
         Left            =   5595
         TabIndex        =   224
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "("
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   8
         Left            =   6360
         TabIndex        =   225
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ")"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   9
         Left            =   7125
         TabIndex        =   226
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   10
         Left            =   7890
         TabIndex        =   227
         Top             =   0
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   11
         Left            =   240
         TabIndex        =   228
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   12
         Left            =   1005
         TabIndex        =   229
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   13
         Left            =   1770
         TabIndex        =   230
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   14
         Left            =   2535
         TabIndex        =   231
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   15
         Left            =   3300
         TabIndex        =   232
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   16
         Left            =   4065
         TabIndex        =   233
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   17
         Left            =   4830
         TabIndex        =   234
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   18
         Left            =   5595
         TabIndex        =   235
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   19
         Left            =   6360
         TabIndex        =   236
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   20
         Left            =   7125
         TabIndex        =   237
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   21
         Left            =   7890
         TabIndex        =   238
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "ß"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   59
         Left            =   8655
         TabIndex        =   239
         Top             =   645
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   22
         Left            =   480
         TabIndex        =   240
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Q"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   23
         Left            =   1245
         TabIndex        =   241
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "W"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   24
         Left            =   2010
         TabIndex        =   242
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "E"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   25
         Left            =   2775
         TabIndex        =   243
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "R"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   26
         Left            =   3540
         TabIndex        =   244
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "T"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   27
         Left            =   4305
         TabIndex        =   245
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Z"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   28
         Left            =   5070
         TabIndex        =   246
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "U"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   29
         Left            =   5835
         TabIndex        =   247
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "I"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   30
         Left            =   6600
         TabIndex        =   248
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "O"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   31
         Left            =   7365
         TabIndex        =   249
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "P"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   32
         Left            =   8130
         TabIndex        =   250
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ü"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   33
         Left            =   8895
         TabIndex        =   251
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "*"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   34
         Left            =   9660
         TabIndex        =   252
         Top             =   1290
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   35
         Left            =   720
         TabIndex        =   253
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   36
         Left            =   1485
         TabIndex        =   254
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "S"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   37
         Left            =   2250
         TabIndex        =   255
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   38
         Left            =   3015
         TabIndex        =   256
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "F"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   39
         Left            =   3780
         TabIndex        =   257
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "G"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   40
         Left            =   4545
         TabIndex        =   258
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "H"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   41
         Left            =   5310
         TabIndex        =   259
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "J"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   42
         Left            =   6075
         TabIndex        =   260
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   43
         Left            =   6840
         TabIndex        =   261
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "L"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   44
         Left            =   7605
         TabIndex        =   262
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   45
         Left            =   8370
         TabIndex        =   263
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ä"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   46
         Left            =   9135
         TabIndex        =   264
         Top             =   1935
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   47
         Left            =   960
         TabIndex        =   265
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Y"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   48
         Left            =   1725
         TabIndex        =   266
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   49
         Left            =   2490
         TabIndex        =   267
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   50
         Left            =   3255
         TabIndex        =   268
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "V"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   51
         Left            =   4020
         TabIndex        =   269
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   52
         Left            =   4785
         TabIndex        =   270
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "N"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   53
         Left            =   5550
         TabIndex        =   271
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "M"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   54
         Left            =   6315
         TabIndex        =   272
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ";"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   55
         Left            =   7080
         TabIndex        =   273
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   56
         Left            =   7845
         TabIndex        =   274
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "_"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   57
         Left            =   8610
         TabIndex        =   275
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   58
         Left            =   9375
         TabIndex        =   276
         Top             =   2595
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   2
         Left            =   8655
         TabIndex        =   277
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "<<<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   3
         Left            =   9420
         TabIndex        =   278
         Top             =   645
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ">>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   0
         Left            =   9915
         TabIndex        =   279
         Top             =   1935
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   1
         Left            =   10140
         TabIndex        =   280
         Top             =   2595
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Rückg"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   4
         Left            =   10425
         TabIndex        =   281
         Top             =   1290
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2081
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "A -> a"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Zielfeld:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   9840
         TabIndex        =   87
         Top             =   80
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   86
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "-1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   2
         Left            =   10680
         TabIndex        =   85
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'Kein
      Height          =   1335
      Left            =   9120
      TabIndex        =   11
      Top             =   7440
      Width           =   2535
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   2520
         TabIndex        =   137
         Top             =   1560
         Visible         =   0   'False
         Width           =   11655
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   3360
            TabIndex        =   192
            Text            =   "Combo3"
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   37
            Left            =   8880
            MaxLength       =   10
            TabIndex        =   190
            Top             =   1560
            Width           =   2055
         End
         Begin sevCommand3.Command Command1 
            Height          =   285
            Index           =   7
            Left            =   8400
            TabIndex        =   189
            ToolTipText     =   "Bankleitzahlen aktualisieren"
            Top             =   1560
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
            Caption         =   "a"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   4440
            TabIndex        =   187
            Text            =   "Combo2"
            Top             =   1560
            Width           =   3855
         End
         Begin sevCommand3.Command Command1 
            Height          =   285
            Index           =   6
            Left            =   8400
            TabIndex        =   186
            Top             =   1080
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
            Caption         =   "<"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Index           =   36
            Left            =   8880
            MaxLength       =   7
            TabIndex        =   184
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   4440
            TabIndex        =   181
            Text            =   "Combo1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  '2D
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4440
            ScaleHeight     =   225
            ScaleWidth      =   1905
            TabIndex        =   151
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0C000&
            Caption         =   "nur die, die offene KB haben"
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
            Left            =   120
            TabIndex        =   149
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0C000&
            Caption         =   "nur die, die offene Kredite haben"
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
            Left            =   120
            TabIndex        =   148
            Top             =   840
            Width           =   3255
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0C000&
            Caption         =   "nur die, die jemals gekauft haben"
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
            Left            =   120
            TabIndex        =   147
            Top             =   1200
            Width           =   3255
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0C000&
            Caption         =   "nur die, die Bankeinzug gewähren"
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
            Left            =   120
            TabIndex        =   146
            Top             =   1560
            Width           =   3255
         End
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   5
            Left            =   9960
            TabIndex        =   144
            Top             =   720
            Width           =   1575
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
            Caption         =   "Zurück"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   3
            Left            =   9960
            TabIndex        =   142
            Top             =   240
            Width           =   1575
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
            Caption         =   "Import/Export"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   7
            Left            =   6840
            TabIndex        =   141
            Top             =   720
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
            Caption         =   "nicht gelieferte"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   6
            Left            =   8400
            TabIndex        =   140
            Top             =   720
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
            Caption         =   "gelieferte"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   5
            Left            =   6840
            TabIndex        =   139
            Top             =   360
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
            Caption         =   "bestellte"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   255
            Index           =   8
            Left            =   8400
            TabIndex        =   138
            Top             =   360
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
            Caption         =   "noch offene"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label13 
            Caption         =   "Rabatt"
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
            Left            =   3360
            TabIndex        =   193
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            Caption         =   "Konto"
            Height          =   255
            Index           =   7
            Left            =   9960
            TabIndex        =   191
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C000&
            Caption         =   "BLZ - sortiert nach BLZ"
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
            Left            =   4440
            MouseIcon       =   "frmWKL13.frx":0442
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   188
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00C0C000&
            Caption         =   "Bonus"
            Height          =   255
            Index           =   6
            Left            =   7440
            TabIndex        =   185
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Preiskennzeichen"
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
            Left            =   4440
            TabIndex        =   182
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "besondere Einschränkungen"
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
            Left            =   120
            TabIndex        =   150
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Rechts
            Caption         =   "Kunden mit Kundenbestellungen"
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
            Left            =   6960
            TabIndex        =   145
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   21
         Left            =   120
         MouseIcon       =   "frmWKL13.frx":074C
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":0A56
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   166
         Top             =   3720
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   360
         MouseIcon       =   "frmWKL13.frx":0AE0
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":0DEA
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   165
         Top             =   3360
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   120
         MouseIcon       =   "frmWKL13.frx":0E74
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":117E
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   164
         Top             =   3360
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   600
         MouseIcon       =   "frmWKL13.frx":1208
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":1512
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   163
         Top             =   3000
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   360
         MouseIcon       =   "frmWKL13.frx":159C
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":18A6
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   162
         Top             =   3000
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   120
         MouseIcon       =   "frmWKL13.frx":1930
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":1C3A
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   161
         Top             =   3000
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   840
         MouseIcon       =   "frmWKL13.frx":1CC4
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":1FCE
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   160
         Top             =   2640
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   600
         MouseIcon       =   "frmWKL13.frx":2058
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":2362
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   159
         Top             =   2640
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   360
         MouseIcon       =   "frmWKL13.frx":23EC
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":26F6
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   158
         Top             =   2640
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   120
         MouseIcon       =   "frmWKL13.frx":2780
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":2A8A
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   157
         Top             =   2640
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   1080
         MouseIcon       =   "frmWKL13.frx":2B14
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":2E1E
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   156
         Top             =   2280
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   840
         MouseIcon       =   "frmWKL13.frx":2EA8
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":31B2
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   155
         Top             =   2280
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   600
         MouseIcon       =   "frmWKL13.frx":323C
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":3546
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   154
         Top             =   2280
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   360
         MouseIcon       =   "frmWKL13.frx":35D0
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":38DA
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   153
         Top             =   2280
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmWKL13.frx":3964
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL13.frx":3C6E
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   152
         Top             =   2280
         Width           =   285
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   4
         Left            =   10080
         TabIndex        =   143
         Top             =   1080
         Width           =   1575
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
         Caption         =   "mehr..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   360
         Left            =   1920
         TabIndex        =   127
         ToolTipText     =   "Kalender"
         Top             =   1440
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   125
         Top             =   1440
         Width           =   1695
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   111
         Top             =   840
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   9
         TabIndex        =   0
         Text            =   "JOOP"
         Top             =   360
         Width           =   1420
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "1234567890123"
         Top             =   360
         Width           =   1060
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   2
         Text            =   "1234567890123"
         Top             =   360
         Width           =   3100
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "1234567890123"
         Top             =   360
         Width           =   940
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6720
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "1234567890123"
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzeige nach Merkmal, Name, Vorname, Ort sortiert"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   97
         Top             =   1560
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzeige nach Datum2, Name, Vorname, Ort sortiert"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   96
         Top             =   1320
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzeige nach Datum1, Name, Vorname, Ort sortiert"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   95
         Top             =   1080
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzeige nach Name, Vorname, Ort sortiert"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   94
         Top             =   840
         Value           =   -1  'True
         Width           =   5055
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   10080
         TabIndex        =   83
         Top             =   1560
         Width           =   1575
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   0
         Left            =   10080
         TabIndex        =   5
         Top             =   120
         Width           =   1575
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   10080
         TabIndex        =   6
         Top             =   600
         Width           =   1575
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
         Caption         =   "Neu"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   210
         Index           =   9
         Left            =   3000
         TabIndex        =   173
         Top             =   2040
         Width           =   1215
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
         Caption         =   "Aktualisieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   174
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   175
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   176
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   177
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   178
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
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
         Caption         =   "Zeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   9600
         TabIndex        =   214
         Top             =   1560
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
         Picture         =   "frmWKL13.frx":3CF8
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   210
         Index           =   16
         Left            =   4320
         TabIndex        =   284
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
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
         Caption         =   "gelöschte Kunden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label11 
         Caption         =   $"frmWKL13.frx":438A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   6
         Left            =   5280
         TabIndex        =   179
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label11 
         Caption         =   "Stand:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   172
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "besondere Einschränkungen"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   171
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "besondere Einschränkungen"
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
         Left            =   1440
         TabIndex        =   170
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "besondere Einschränkungen"
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
         Left            =   1440
         TabIndex        =   169
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "besondere Einschränkungen"
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
         Index           =   4
         Left            =   1440
         TabIndex        =   168
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         Caption         =   "besondere Einschränkungen"
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
         Index           =   5
         Left            =   1440
         TabIndex        =   167
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "angelegt seit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Farben"
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
         Index           =   5
         Left            =   120
         TabIndex        =   112
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8880
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "von"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8880
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Kunden-Nr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Kürzel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   14
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "PLZ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Ort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   10080
      TabIndex        =   89
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   10
         Left            =   8760
         TabIndex        =   205
         Top             =   4320
         Width           =   1455
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
      Begin sevCommand3.Command Command6 
         Height          =   495
         Left            =   10320
         TabIndex        =   92
         Top             =   4320
         Width           =   1440
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
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   91
         Top             =   600
         Width           =   11655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Notizen zum Kunden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'Kein
      Height          =   1575
      Left            =   9120
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   3000
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5520
         MouseIcon       =   "frmWKL13.frx":443A
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   209
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5160
         MouseIcon       =   "frmWKL13.frx":4744
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   208
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4680
         MouseIcon       =   "frmWKL13.frx":4A4E
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   207
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   1455
         Left            =   4680
         MouseIcon       =   "frmWKL13.frx":4D58
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   1455
         ScaleWidth      =   1455
         TabIndex        =   206
         Top             =   720
         Width           =   1455
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   10440
         TabIndex        =   62
         Top             =   4320
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
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
         Index           =   0
         Left            =   9120
         TabIndex        =   61
         Top             =   4320
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
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
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   7920
         TabIndex        =   60
         Top             =   4320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
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
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   6720
         TabIndex        =   59
         Top             =   4320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
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
         Caption         =   "Notizen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Left            =   5520
         TabIndex        =   58
         Top             =   4320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
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
         Caption         =   "Kosmetik"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   4
         Left            =   4440
         TabIndex        =   57
         Top             =   4320
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
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
      Begin sevCommand3.Command Command11 
         Height          =   255
         Left            =   2760
         TabIndex        =   131
         Top             =   4560
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Rechts ausgerichtet
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   130
         ToolTipText     =   "bonusfähig"
         Top             =   4560
         Width           =   255
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
         Index           =   35
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   128
         Top             =   2760
         Width           =   975
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
         Index           =   17
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   27
         Top             =   240
         Width           =   495
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
         Index           =   32
         Left            =   960
         MaxLength       =   8
         TabIndex        =   56
         Top             =   4560
         Width           =   1335
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
         Index           =   31
         Left            =   960
         MaxLength       =   18
         TabIndex        =   55
         Top             =   4200
         Width           =   1335
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
         Index           =   29
         Left            =   960
         MaxLength       =   1
         TabIndex        =   52
         Text            =   "N"
         Top             =   3840
         Width           =   495
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
         Index           =   26
         Left            =   960
         MaxLength       =   12
         TabIndex        =   49
         Top             =   3480
         Width           =   1335
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
         Index           =   22
         Left            =   960
         MaxLength       =   10
         TabIndex        =   45
         Top             =   3120
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
         Index           =   16
         Left            =   7800
         MaxLength       =   20
         TabIndex        =   38
         Top             =   1680
         Width           =   3975
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
         Index           =   18
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   40
         Top             =   2040
         Width           =   3975
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
         Index           =   14
         Left            =   9000
         MaxLength       =   35
         TabIndex        =   36
         Top             =   1320
         Width           =   2775
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
         Left            =   7800
         MaxLength       =   35
         TabIndex        =   31
         ToolTipText     =   "Dr. , Dipl. Ing."
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cbo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmWKL13.frx":5062
         Left            =   10440
         List            =   "frmWKL13.frx":506F
         TabIndex        =   29
         ToolTipText     =   "Fräulein, Frau, Herr"
         Top             =   240
         Width           =   1335
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
         Index           =   30
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   53
         Text            =   "N"
         Top             =   3840
         Width           =   375
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
         Index           =   27
         Left            =   4200
         MaxLength       =   6
         TabIndex        =   54
         Top             =   3840
         Width           =   615
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
         Index           =   23
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   46
         Top             =   3120
         Width           =   975
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
         Index           =   20
         Left            =   960
         MaxLength       =   74
         TabIndex        =   43
         Top             =   2760
         Width           =   1815
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
         Index           =   19
         Left            =   960
         MaxLength       =   12
         TabIndex        =   41
         Top             =   2400
         Width           =   3615
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
         Index           =   34
         Left            =   960
         MaxLength       =   15
         TabIndex        =   39
         Top             =   2040
         Width           =   3615
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
         Index           =   15
         Left            =   960
         MaxLength       =   35
         TabIndex        =   37
         Top             =   1680
         Width           =   3615
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
         Index           =   33
         Left            =   7800
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3480
         Width           =   3975
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
         Index           =   25
         Left            =   10440
         MaxLength       =   12
         TabIndex        =   48
         Top             =   3120
         Width           =   1335
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
         Index           =   24
         Left            =   7800
         MaxLength       =   12
         TabIndex        =   47
         Top             =   3120
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
         Index           =   7
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   28
         ToolTipText     =   "M oder W"
         Top             =   240
         Width           =   495
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
         Index           =   12
         Left            =   960
         MaxLength       =   35
         TabIndex        =   34
         Top             =   1320
         Width           =   3615
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
         Index           =   10
         Left            =   960
         MaxLength       =   35
         TabIndex        =   32
         Top             =   960
         Width           =   3615
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
         Left            =   960
         MaxLength       =   35
         TabIndex        =   30
         Top             =   600
         Width           =   3615
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
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   26
         Top             =   240
         Width           =   735
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
         Left            =   960
         MaxLength       =   9
         TabIndex        =   25
         Top             =   240
         Width           =   975
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
         Index           =   21
         Left            =   7800
         MaxLength       =   74
         TabIndex        =   44
         Top             =   2760
         Width           =   3975
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
         Index           =   28
         Left            =   7800
         MaxLength       =   18
         TabIndex        =   42
         Top             =   2400
         Width           =   3975
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
         Index           =   13
         Left            =   7800
         MaxLength       =   7
         TabIndex        =   35
         Top             =   1320
         Width           =   735
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
         Index           =   11
         Left            =   7800
         MaxLength       =   35
         TabIndex        =   33
         Top             =   960
         Width           =   3975
      End
      Begin sevCommand3.Command Command9 
         Height          =   255
         Left            =   4200
         TabIndex        =   113
         Top             =   3480
         Visible         =   0   'False
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   50
         Top             =   3480
         Width           =   495
      End
      Begin sevCommand3.Command Command8 
         Height          =   255
         Left            =   9360
         TabIndex        =   110
         Top             =   240
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
      Begin sevCommand3.Command Command45 
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   108
         Top             =   3840
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Caption         =   "ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame6"
         Height          =   735
         Left            =   4680
         TabIndex        =   194
         Top             =   5040
         Visible         =   0   'False
         Width           =   6615
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Netto-KVK"
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
            Index           =   6
            Left            =   2880
            TabIndex        =   289
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Schnitt-EK"
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
            Left            =   2880
            TabIndex        =   200
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Listen-EK"
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
            Left            =   1680
            TabIndex        =   199
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Listen-VK"
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
            Index           =   1
            Left            =   5520
            TabIndex        =   198
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kassen-VK"
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
            Index           =   0
            Left            =   4200
            TabIndex        =   197
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "spez KVK"
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
            Index           =   4
            Left            =   4200
            TabIndex        =   196
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "LVK m A"
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
            Index           =   5
            Left            =   5520
            TabIndex        =   195
            ToolTipText     =   $"frmWKL13.frx":5088
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            Caption         =   "Preis Kz.:"
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
            Index           =   55
            Left            =   600
            TabIndex        =   202
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            Caption         =   "bei Listen-EK: immer kleinsten EK beziehen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   155
            Index           =   8
            Left            =   0
            MouseIcon       =   "frmWKL13.frx":5128
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   201
            ToolTipText     =   "Daten aktualisieren"
            Top             =   280
            Width           =   2655
         End
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   4200
         TabIndex        =   215
         ToolTipText     =   "Kalender"
         Top             =   3120
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   1
         Left            =   2160
         TabIndex        =   216
         ToolTipText     =   "Kalender"
         Top             =   3120
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   17
         Left            =   3840
         TabIndex        =   285
         Top             =   4320
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   873
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
         ToolTip         =   "Datenschutzblatt drucken"
         ButtonStyle     =   2
         Caption         =   "DS"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check7 
         Alignment       =   1  'Rechts ausgerichtet
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   287
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "DS unterschrieben"
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
         Index           =   40
         Left            =   2760
         TabIndex        =   288
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bild aufnehmen"
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
         Index           =   39
         Left            =   5640
         MouseIcon       =   "frmWKL13.frx":5432
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   282
         Top             =   3400
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bild löschen"
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
         Index           =   38
         Left            =   4680
         MouseIcon       =   "frmWKL13.frx":573C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   213
         Top             =   3400
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bild laden"
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
         Index           =   37
         Left            =   4680
         MouseIcon       =   "frmWKL13.frx":5A46
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   212
         Top             =   3140
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Bildname"
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
         Left            =   4680
         TabIndex        =   210
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   6240
         Top             =   2400
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
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
         Index           =   36
         Left            =   5160
         TabIndex        =   204
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Preis Kz.:"
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
         Index           =   35
         Left            =   5160
         TabIndex        =   203
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "seit::"
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
         Index           =   34
         Left            =   2880
         TabIndex        =   129
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale:"
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
         Index           =   32
         Left            =   3360
         TabIndex        =   124
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Telefon:"
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
         Left            =   0
         TabIndex        =   69
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Straße:"
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
         Index           =   7
         Left            =   0
         TabIndex        =   66
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Name:"
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
         Index           =   5
         Left            =   0
         TabIndex        =   64
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Firma:"
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
         Left            =   0
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "KdNr:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Datum2:"
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
         Index           =   18
         Left            =   2400
         TabIndex        =   75
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Ort:"
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
         Left            =   8520
         TabIndex        =   68
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Datei:"
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
         Index           =   14
         Left            =   0
         TabIndex        =   116
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Mobil:"
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
         Index           =   29
         Left            =   0
         TabIndex        =   115
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bankeinzug:"
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
         Index           =   31
         Left            =   2400
         TabIndex        =   114
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Farbe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   30
         Left            =   8400
         TabIndex        =   109
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Email:"
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
         Index           =   28
         Left            =   7080
         TabIndex        =   106
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Anrede:"
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
         Left            =   9720
         TabIndex        =   105
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bonus:"
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
         Index           =   27
         Left            =   0
         TabIndex        =   93
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kd.-Karte:"
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
         Index           =   26
         Left            =   0
         TabIndex        =   88
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Löschen:"
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
         Index           =   25
         Left            =   1920
         TabIndex        =   82
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Gesperrt:"
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
         Index           =   24
         Left            =   0
         TabIndex        =   81
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EC-Karte:"
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
         Index           =   23
         Left            =   6840
         TabIndex        =   80
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Rabatt:"
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
         Index           =   22
         Left            =   3360
         TabIndex        =   79
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "offen:"
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
         Index           =   21
         Left            =   0
         TabIndex        =   78
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz VJ:"
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
         Index           =   20
         Left            =   9360
         TabIndex        =   77
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz LJ:"
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
         Index           =   19
         Left            =   6840
         TabIndex        =   76
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Gebdatum:"
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
         Index           =   17
         Left            =   0
         TabIndex        =   74
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kurztext2:"
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
         Index           =   16
         Left            =   6840
         TabIndex        =   73
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Land:"
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
         Index           =   15
         Left            =   0
         TabIndex        =   72
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Merkmal:"
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
         Left            =   6840
         TabIndex        =   71
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Fax:"
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
         Left            =   6840
         TabIndex        =   70
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "PLZ:"
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
         Left            =   6840
         TabIndex        =   67
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Vorname:"
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
         Index           =   6
         Left            =   6840
         TabIndex        =   65
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Titel:"
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
         Index           =   4
         Left            =   7080
         TabIndex        =   63
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Geschlecht:"
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
         Left            =   6840
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kürzel:"
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
         Index           =   1
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neukunden AM"
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
      Index           =   20
      Left            =   9000
      TabIndex        =   136
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neukunden VJM"
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
      Index           =   21
      Left            =   9000
      TabIndex        =   135
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neukunden AJ"
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
      Index           =   22
      Left            =   6480
      TabIndex        =   134
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neukunden VJ"
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
      Index           =   23
      Left            =   6480
      TabIndex        =   133
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Neukunden:"
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
      Index           =   24
      Left            =   4680
      MouseIcon       =   "frmWKL13.frx":5D50
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   132
      ToolTipText     =   "Daten aktualisieren"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label0 
      BackColor       =   &H00C0C000&
      Caption         =   "Kundendaten"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   104
      Top             =   0
      Width           =   3945
   End
End
Attribute VB_Name = "frmWKL13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPrueD          As Integer
Dim SpaltennummerAWM As Byte
Dim SpaltennummerKUNDNR As Byte
Dim SpaltennummerSTERNE As Byte
Dim bSterneDa As Boolean

Private Sub cbo1_Click()
On Error GoTo LOKAL_ERROR

If cbo1.Text = "Herr" Then
    Text1(7).Text = "M"
ElseIf cbo1.Text = "Frau" Then
    Text1(7).Text = "W"
ElseIf cbo1.Text = "Familie" Then
    Text1(7).Text = "F"
Else
    Text1(7).Text = ""
End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR

If Check1.Value = vbChecked Then
    Command9.Visible = True
Else
    Command9.Visible = False
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1


End Sub



Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case Is = 0        ' Kalender
            Text1(23).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        
        Case Is = 1        ' Kalender
            Text1(22).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
        End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "

    Fehlermeldung1
   
End Sub

Private Sub Command10_Click()
On Error GoTo LOKAL_ERROR
    
    
    Text3.Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
    Text3.SetFocus
    'fertig
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    zeigeHilfeDabapfad "LPROTOK", "BONUSMANU.txt"
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command41_Click()
On Error GoTo LOKAL_ERROR
    
    reportbildschirm "dWKL001b", "aWKL13b"

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command41_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command45_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Frame6.Visible = True
            Command45(0).Visible = False
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command45_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR
    
    gsBackcolor = Label4(30).BackColor
    gsForecolor = Label4(30).ForeColor
    gsKundenFarbe = Label4(30).Tag
    frmWKL65.Show 1
    
    Label4(30).BackColor = gsBackcolor
    Label4(30).ForeColor = gsForecolor
    Label4(30).Tag = gsKundenFarbe
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command9_Click()
On Error GoTo LOKAL_ERROR
    
    gckuVorname = Text1(11).Text
    gckuname = Text1(10).Text
    gckundnr = Text1(5).Text
    
    frmWKL66.Show 1
    
    gckuVorname = ""
    gckuname = ""
    gckundnr = ""
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "KUNDNR"
    gstab = "BEAKU"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    Label1(24).ForeColor = glS1
    Label1(8).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "TOPK" & srechnertab, gdBase
    loeschNEW "kuntt", gdBase
    LogtoEnd Me
    
    'Muss so bleiben wegen Kasse neuer KUNDE
'    gcKundenNr = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL13Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame0.Top = 800
    Frame0.Left = 0
    Frame0.Height = 5655
    Frame0.Width = 11895
    
    Frame1.Top = 0
    Frame1.Left = 0
    Frame1.Height = 8655
    Frame1.Width = 11895
    
    Frame2.Top = 5200
    Frame2.Left = 0
    Frame2.Height = 4455
    Frame2.Width = 11895
    
    Frame3.Top = 0
    Frame3.Left = 0
    Frame3.Height = 5055
    Frame3.Width = 11895
    
    Frame4.Top = 0
    Frame4.Left = 0
    Frame4.Height = 5175
    Frame4.Width = 11895
    
    Frame5.Top = 2040
    Frame5.Left = 120
    Frame5.Height = 1935
    Frame5.Width = 11655
    
    Frame6.Top = 3840
    Frame6.Left = 4920
    Frame6.Height = 495
    Frame6.Width = 6855
    
    Frame11.Top = 600
    Frame11.Left = 120
    Frame11.Height = 4335
    Frame11.Width = 11895

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL13Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Function OffeneKB(cSuch As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    OffeneKB = False
    
    cSQL = "Select KUNDNR from KUNDBEST where KUNDNR = " & Val(cSuch) & " and StatusARTIKEL = 'INBESTELLUNG' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        OffeneKB = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "OffeneKB"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function


Private Function JEMALSKAUF(cSuch As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    JEMALSKAUF = False
    
    cSQL = "Select * from KASSJOUR where KUNDNR = " & Val(cSuch)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        JEMALSKAUF = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "JEMALSKAUF"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function OffeneKredite(cSuch As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    OffeneKredite = 0
    
    cSQL = "Select SUM(GVKPR) as maxi from Kredit where KUNDNR = " & Val(cSuch)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            OffeneKredite = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "OffeneKredite"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Public Sub HoleDatenWKL13(cSuch As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim dWert       As Double
    Dim iFehler     As Integer
    Dim lJahr       As Long
    Dim dUmsatzLJ   As Double
    Dim dUmsatzVJ   As Double
    
    
    iFehler = 0
    
    lJahr = Year(Now)
    
    dUmsatzLJ = umsatzLFJahr1(cSuch)
    dUmsatzVJ = umsatzVJJahr(cSuch)

    cSQL = "Select * from KUNDEN where KUNDNR = " & cSuch & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Kundnr) Then
            ctmp = rsrs!Kundnr
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(5).Text = ctmp
        
        If Not IsNull(rsrs!Kuerzel) Then
            ctmp = rsrs!Kuerzel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(6).Text = ctmp
        
        If Not IsNull(rsrs!geschlecht) Then
            ctmp = rsrs!geschlecht
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(7).Text = ctmp
        
        If Not IsNull(rsrs!firma) Then
            ctmp = rsrs!firma
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(8).Text = ctmp
        
        If Not IsNull(rsrs!titel) Then
            ctmp = rsrs!titel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(9).Text = ctmp
        
        If Not IsNull(rsrs!name) Then
            ctmp = rsrs!name
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(10).Text = ctmp
        
        If Not IsNull(rsrs!vorname) Then
            ctmp = rsrs!vorname
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(11).Text = ctmp
        
        If Not IsNull(rsrs!strasse) Then
            ctmp = rsrs!strasse
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(12).Text = ctmp
        
        If Not IsNull(rsrs!Plz) Then
            ctmp = rsrs!Plz
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(13).Text = ctmp
        
        If Not IsNull(rsrs!STADT) Then
            ctmp = rsrs!STADT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(14).Text = ctmp
        
        
        
        If Not IsNull(rsrs!Mobiltel) Then
            ctmp = rsrs!Mobiltel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(34).Text = ctmp
        
        If Not IsNull(rsrs!angelegt) Then
            dWert = rsrs!angelegt
        Else
            dWert = 0
        End If
        If dWert > 0 Then
            ctmp = Format$(dWert, "DD.MM.YY")
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(35).Text = ctmp
        
        
        If Not IsNull(rsrs!FILIALNR) Then
            ctmp = rsrs!FILIALNR
        Else
            ctmp = "0"
        End If
        ctmp = Trim$(ctmp)
        Text1(17).Text = ctmp
        
        If Not IsNull(rsrs!Tel) Then
            ctmp = rsrs!Tel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(15).Text = ctmp
        
        If Not IsNull(rsrs!FAXNR) Then
            ctmp = rsrs!FAXNR
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(16).Text = ctmp
        
        If Not IsNull(rsrs!anrede) Then
            ctmp = rsrs!anrede
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        
        cbo1.Text = ctmp
        
        If Not IsNull(rsrs!MERKMAL) Then
            ctmp = rsrs!MERKMAL
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(18).Text = ctmp
        
        If Not IsNull(rsrs!MERKMAL2) Then
            ctmp = rsrs!MERKMAL2
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        If ctmp = "N" Then
            Check6.Value = vbUnchecked
        Else
            Check6.Value = vbChecked
        End If
        
        If Not IsNull(rsrs!FORMATDAT) Then
            ctmp = rsrs!FORMATDAT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(19).Text = ctmp
        
        If Not IsNull(rsrs!KURZTEXT1) Then
            ctmp = rsrs!KURZTEXT1
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(20).Text = ctmp
        
        If Not IsNull(rsrs!KurzTEXT2) Then
            ctmp = rsrs!KurzTEXT2
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(21).Text = ctmp
        
        If Not IsNull(rsrs!BE) Then
            If rsrs!BE = True Then
                Check1.Value = vbChecked
                Command9.Visible = True
            Else
                Check1.Value = vbUnchecked
                Command9.Visible = False
            End If
        Else
            Check1.Value = vbUnchecked
            Command9.Visible = False
        End If
        
        If Not IsNull(rsrs!DS) Then
            If rsrs!DS = True Then
                Check7.Value = vbChecked
            Else
                Check7.Value = vbUnchecked
            End If
        Else
            Check7.Value = vbUnchecked
        End If
        
        If Not IsNull(rsrs!Datum1) Then
            dWert = rsrs!Datum1
        Else
            dWert = 0
        End If
        If dWert > 0 Then
            ctmp = Format$(dWert, "DD.MM.YYYY")
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(22).Text = ctmp
        iFehler = 1
        
        iFehler = 2
        If Not IsNull(rsrs!DATUM2) Then
            dWert = rsrs!DATUM2
        Else
            dWert = 0
        End If
        
        If dWert > 0 Then
            ctmp = Format$(dWert, "DD.MM.YYYY")
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(23).Text = ctmp

        iFehler = 4
        
        dWert = dUmsatzLJ
        ctmp = Format$(dWert, "#######0.00")
        Text1(24).Text = ctmp
        
        dWert = dUmsatzVJ
        ctmp = Format$(dWert, "#######0.00")
        Text1(25).Text = ctmp
        
        dWert = OffeneKredite(cSuch)
        
        ctmp = Format$(dWert, "#######0.00")
        Text1(26).Text = ctmp
        
        If Not IsNull(rsrs!RABATT) Then
            dWert = rsrs!RABATT
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "##0.00")
        Text1(27).Text = ctmp
        
        If Not IsNull(rsrs!ECIDENT) Then
            ctmp = rsrs!ECIDENT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(28).Text = ctmp
        
        If Not IsNull(rsrs!GESPERRT) Then
            ctmp = rsrs!GESPERRT
        Else
            ctmp = "N"
        End If
        ctmp = Trim$(ctmp)
        Text1(29).Text = ctmp
        
        Text1(30).Text = "N"
        
        If Not IsNull(rsrs!KUNDKART) Then
            ctmp = rsrs!KUNDKART
        Else
            ctmp = ""
        End If
        Text1(31).Text = ctmp
        
        If Not IsNull(rsrs!BONUS) Then
            ctmp = Format$(rsrs!BONUS, "####0.00")
        Else
            ctmp = "0,00"
        End If
        
        
        
        Text1(32).Text = ctmp
        
        If Not IsNull(rsrs!Email) Then
            ctmp = rsrs!Email
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(33).Text = ctmp
        
        If Not IsNull(rsrs!NOTIZEN) Then
            ctmp = rsrs!NOTIZEN
        Else
            ctmp = ""
        End If
        Text2.Text = ctmp
        
        If Not IsNull(rsrs!PREISKZ) Then
            dWert = rsrs!PREISKZ
        Else
            dWert = 0
        End If
        Option1(dWert).Value = True
        
        If Option1(0).Value Then
            Label4(36).Caption = Option1(0).Caption
        ElseIf Option1(1).Value Then
            Label4(36).Caption = Option1(1).Caption
        ElseIf Option1(2).Value Then
            Label4(36).Caption = Option1(2).Caption
        ElseIf Option1(3).Value Then
            Label4(36).Caption = Option1(3).Caption
        ElseIf Option1(4).Value Then
            Label4(36).Caption = Option1(4).Caption
        ElseIf Option1(5).Value Then
            Label4(36).Caption = Option1(5).Caption
        ElseIf Option1(6).Value Then
            Label4(36).Caption = Option1(6).Caption
        End If
        
        
        If Not IsNull(rsrs!AWM) Then
            ctmp = rsrs!AWM
        Else
            ctmp = "0"
        End If
        If ctmp = "" Then ctmp = "0"
        
        Label4(30).ForeColor = vbBlack
        Label4(30).BackColor = glfarbe(ctmp)
        Label4(30).Tag = ctmp
        
        Label3(2).Caption = 5
        Frame3.Visible = True
        If gbBILDTAST = False Then
            Frame2.Visible = False
        Else
            Frame2.Visible = True
        End If

        Frame1.Visible = False
        Frame6.Visible = False 'preiskz
        Command45(0).Visible = True
        Frame3.Refresh
        
        Zeige_Kunden_Bilder cSuch, 0
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDatenWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub

Private Sub LeereDialogWKL13()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text2.Text = ""
    Label3(2).Caption = "0"
    Combo2.Text = "BLZ auswählen"
    Combo1.Text = "bitte wählen"
    Combo3.Text = "Rabatt wählen"
    Text1(37).Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub LeereKundenMaskeWKL13()
    On Error GoTo LOKAL_ERROR
    
    Dim lWert As Long
    
    For lWert = 5 To 27
        Text1(lWert).Text = ""
    Next lWert
    Text1(28).Text = ""
    Text1(29).Text = "N"
    Text1(31).Text = ""
    Text1(33).Text = ""
    Text1(34).Text = ""
    Text1(35).Text = ""
    Text1(32).Text = "0,00"
    cbo1.Visible = True
    Check1.Value = vbUnchecked
    Label4(30).ForeColor = vbBlack
    Label4(30).BackColor = glfarbe(0)
    Label4(30).Tag = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereKundenMaskeWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheKundeWKL13(cKdnr As String)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    BeginTrans
    If Not rsrs.EOF Then
        rsrs.Edit
        rsrs!Status = "D"
        rsrs!SYNStatus = "D"
        rsrs.Update
        MSFlexGrid1.Row = glSelect
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = "gelöscht"
    End If
    CommitTrans
    rsrs.Close: Set rsrs = Nothing
    
    If gcFilNr <> 0 Then
        cSQL = "Select * from KUNDEDEL where KUNDNR = -1"
        FnOpenrecordset rsrs, cSQL, 1, gdBase
        
        rsrs.AddNew
        
        rsrs!ADATE = Format$(Now, "HH:MM:SS")
        rsrs!AZEIT = Fix(Now)
        rsrs!Kundnr = cKdnr
        rsrs!BEDNU = gcBedienerNr
        rsrs!FILIALE = gcFilNr
        rsrs!SENDOK = False
        
        rsrs.Update
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheKundeWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWKL13(bLeereMaske As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKey As String
    Dim dWert As Double
    Dim ctmp As String
    Dim cMaxNr As String
    Dim iDbNr As Integer
    Dim bNeu As Boolean
    Dim cKundnr As String
    Dim lWert As Long
    Dim cDBFeld As String
    Dim iStufe As Integer
    
    iStufe = 0
    bNeu = False
    
    cKey = Text1(5).Text
    cKey = Trim$(cKey)
    
    If cKey = "" Then
        MsgBox "Schreiben nicht möglich, da keine Kundennummer vorhanden!", vbCritical, "Winkiss Hinweis:"
        Text1(5).SetFocus
        Exit Sub
    End If
    
    iStufe = 1
    cSQL = "Select max(KUNDNR) from KUNDEN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs.Fields(0)) Then
            dWert = rsrs.Fields(0)
        Else
            dWert = 0
        End If
    Else
        dWert = 0
    End If
    dWert = dWert + 1
    cMaxNr = Format$(dWert, "#####0")
    
    rsrs.Close: Set rsrs = Nothing
    
    iStufe = 2
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKey & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    iStufe = 3
    If Not rsrs.EOF Then
    
        iStufe = 4
        If giDlgZustand = giNEU Then
            MsgBox "Kundennummer bereits vorhanden!" & vbCrLf & "Bitte andere Kundennummer eingeben (nächste freie Nummer = " & cMaxNr & " ).", vbCritical, "DOPPELTE KUNDENNUMMER"
            Text1(5).SetFocus
            Exit Sub
        End If
            
        rsrs.Edit
        rsrs!Status = "E" 'Status editieren
        rsrs!SYNStatus = "E" 'Status editieren
       
    Else
        iStufe = 5
        If gbFilNr Then
            cKundnr = Text1(5).Text
            If Val(cKundnr) <= Val(gcFilNr & "00000") Or Val(cKundnr) > Val(gcFilNr & "99999") Then
                
                ctmp = "Die Kunden-Nr liegt außerhalb Ihres zugeordneten Nummernkreises!" & vbCrLf & vbCrLf
                ctmp = ctmp & "In der Filial-Version von WinKISS gelten für die Vergabe" & vbCrLf
                ctmp = ctmp & "von Kundennummern folgende Regeln:" & vbCrLf & vbCrLf
                ctmp = ctmp & "1. Die Kundennummer ist immer sechsstellig " & vbCrLf
                ctmp = ctmp & "2. Die erste Ziffer der Kundennummer entspricht der Filialnummer " & vbCrLf & vbCrLf
                lWert = fnHoleMaxKundenNr()
                ctmp = ctmp & "Die nächste freie Kundennummer ist " & Format$(lWert, "#####0")
                MsgBox ctmp, vbCritical, "STOP!"
                Text1(5).SetFocus
                Exit Sub
            End If
        End If
        rsrs.AddNew
        rsrs!Status = "A" 'Status addnew
        rsrs!SYNStatus = "A" 'Status addnew
        
        If gcBedKUNEU = "" Then gcBedKUNEU = gcBedienerNr
        rsrs!Rechnr = gcBedKUNEU
''''        Label4(30).Tag = "0"
        bNeu = True
    End If
    
    iStufe = 6
    ctmp = Trim$(Text1(5).Text)
    cDBFeld = "KUNDNR"
    If ctmp = "" Then
        rsrs!Kundnr = Null
    Else
        rsrs!Kundnr = ctmp
    End If
    
    
    iStufe = 11
    cDBFeld = "NAME"
    
    If Len(Text1(10).Text) > 0 Then
        If rsrs!name <> Trim$(Text1(10).Text) Then
            If Len(Text1(10).Text) >= 5 Then
                Text1(6).Text = UCase(Left(Text1(10).Text, 5))
            Else
                Text1(6).Text = UCase(Left(Text1(10).Text, Len(Text1(10).Text)))
            End If
        End If
    End If
    
    rsrs!name = Trim$(Text1(10).Text)
    
    
    
    iStufe = 7
    cDBFeld = "KUERZEL"
    rsrs!Kuerzel = Trim$(Text1(6).Text)
    iStufe = 8
    cDBFeld = "GESCHLECHT"
    rsrs!geschlecht = Trim$(Text1(7).Text)
    iStufe = 9
    cDBFeld = "FIRMA"
    rsrs!firma = Trim$(Text1(8).Text)
    iStufe = 10
    cDBFeld = "TITEL"
    rsrs!titel = Trim$(Text1(9).Text)
    
    iStufe = 12
    cDBFeld = "VORNAME"
    rsrs!vorname = Trim$(Text1(11).Text)
    iStufe = 13
    cDBFeld = "STRASSE"
    rsrs!strasse = Trim$(Text1(12).Text)
    iStufe = 14
    cDBFeld = "PLZ"
    rsrs!Plz = Trim$(Text1(13).Text)
    iStufe = 15
    cDBFeld = "STADT"
    rsrs!STADT = Trim$(Text1(14).Text)
    iStufe = 16
    cDBFeld = "TEL"
    rsrs!Tel = Trim$(Text1(15).Text)
    
    cDBFeld = "MOBILTEL"
    rsrs!Mobiltel = Trim$(Text1(34).Text)
    iStufe = 17
    cDBFeld = "FAXNR"
    rsrs!FAXNR = Trim$(Text1(16).Text)
    iStufe = 18
    cDBFeld = "ANREDE"
    
    rsrs!anrede = Trim$(cbo1.Text)
    iStufe = 19
    
    cDBFeld = "EMAIL"
    rsrs!Email = Trim$(Text1(33).Text)
    iStufe = 20
    
    cDBFeld = "MERKMAL"
    rsrs!MERKMAL = Trim$(Text1(18).Text)
    
    cDBFeld = "MERKMAL2"
    If Check6.Value = vbChecked Then
        rsrs!MERKMAL2 = "J"
    Else
        rsrs!MERKMAL2 = "N"
    End If
    iStufe = 21
    cDBFeld = "FORMATDAT"
    rsrs!FORMATDAT = Trim$(Text1(19).Text)
    iStufe = 22
    cDBFeld = "KURZTEXT1"
    rsrs!KURZTEXT1 = Trim$(Text1(20).Text)
    iStufe = 23
    cDBFeld = "KURZTEXT2"
    rsrs!KurzTEXT2 = Trim$(Text1(21).Text)
    iStufe = 24
    cDBFeld = "KUNDKART"
    rsrs!KUNDKART = Trim$(Text1(31).Text)
    iStufe = 25
    
    cDBFeld = "DATUM1"
    ctmp = Trim$(Text1(22).Text)
    If Not IsDate(ctmp) Then
        Text1(22).Text = ""
    Else
        Text1(22).Text = ctmp
    End If
    
    iStufe = 26
    ctmp = Text1(22).Text
    If ctmp <> "" Then
        dWert = DateValue(ctmp)
        rsrs!Datum1 = dWert
    Else
        rsrs!Datum1 = Null
    End If
    
    
    
    
    iStufe = 27
    
    cDBFeld = "DATUM2"
    ctmp = Trim$(Text1(23).Text)
    If Not IsDate(ctmp) Then
        Text1(23).Text = ""
    Else
        Text1(23).Text = ctmp
    End If
    ctmp = Text1(23).Text
    
    iStufe = 28

    If ctmp <> "" Then
        dWert = DateValue(ctmp)
        rsrs!DATUM2 = dWert
    Else
        rsrs!DATUM2 = Null
    End If
    
    iStufe = 29
    
    cDBFeld = "UMSLJ"
    ctmp = Trim$(Text1(24).Text)
    If ctmp <> "" Then
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        rsrs!UMSLJ = dWert
    End If
    
    iStufe = 30
    
    cDBFeld = "UMSVJ"
    ctmp = Trim$(Text1(25).Text)
    If ctmp <> "" Then
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        rsrs!UMSVJ = dWert
    End If
    
    iStufe = 31
    
    cDBFeld = "OSUM"
    ctmp = Trim$(Text1(26).Text)
    If ctmp <> "" Then
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        rsrs!OSUM = dWert
    End If
    
    iStufe = 32
    
    cDBFeld = "RABATT"
    ctmp = Trim$(Text1(27).Text)
    If ctmp <> "" Then
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        rsrs!RABATT = dWert
    End If
    
    iStufe = 33

    cDBFeld = "BONUS"
    ctmp = Trim$(Text1(32).Text)
    If ctmp = "" Then
        ctmp = "0,00"
    End If
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    iStufe = 34
    
    If IsNull(rsrs!BONUS) Then
        rsrs!BONUS = 0
    End If
    
    If rsrs!Status <> "A" Then
    
        If Not IsNull(rsrs!BONUS) Then
        
            If CDbl(rsrs!BONUS) <> dWert Then
                rsrs!Status = "E"
                rsrs!SYNStatus = "E"
            End If
        Else
            rsrs!Status = "E"
            rsrs!SYNStatus = "E"
        
        End If
    Else
        rsrs!Status = "A"
        rsrs!SYNStatus = "A"
    End If
    
    iStufe = 35
    
    If Not IsNull(rsrs!BONUS) Then
        If Not IsNull(rsrs!TBONUS) Then
            
            rsrs!TBONUS = CDbl(rsrs!TBONUS) + (dWert - CDbl(rsrs!BONUS))
        Else
            rsrs!TBONUS = dWert
        End If
        
    Else
        If Not IsNull(rsrs!TBONUS) Then
            
            rsrs!TBONUS = CDbl(rsrs!TBONUS) + dWert
        Else
            rsrs!TBONUS = dWert
        End If
    
    End If
    
    iStufe = 36
    
    If giDlgZustand <> giNEU Then
        If Format(CDbl(rsrs!BONUS), "####0.00") <> Format(dWert, "####0.00") Then
            schreibeProtokollUNITXT "Kundnr: " & rsrs!Kundnr & " Bonus wurde manuell geändert von " & Format(rsrs!BONUS, "####0.00") & " auf " & Format(dWert, "####0.00") & " ", "BONUSMANU"
        End If
    End If
    
    iStufe = 37
    rsrs!BONUS = dWert
    iStufe = 38
    cDBFeld = "ECIDENT"
    rsrs!ECIDENT = Trim$(Text1(28).Text)
    
    iStufe = 39
    
    cDBFeld = "GESPERRT"
    rsrs!GESPERRT = Trim$(Text1(29).Text)
    
    iStufe = 40
    cDBFeld = "NOTIZEN"
    rsrs!NOTIZEN = Text2.Text
    
    iStufe = 41
    If Check1.Value = vbChecked Then
        rsrs!BE = True
    Else
        cSQL = "Delete from  BANKKU where kundnr = " & Trim$(Text1(5).Text)
        gdBase.Execute cSQL, dbFailOnError
        
        rsrs!BE = False
    End If
    
     iStufe = 41
    If Check7.Value = vbChecked Then
        rsrs!DS = True
    Else
        rsrs!DS = False
    End If
    
    iStufe = 42
    rsrs!AWM = Label4(30).Tag
    iStufe = 43
    cDBFeld = "PREISKZ"
    If Option1(0).Value Then
        rsrs!PREISKZ = 0
    ElseIf Option1(1).Value Then
        rsrs!PREISKZ = 1
    ElseIf Option1(2).Value Then
        rsrs!PREISKZ = 2
    ElseIf Option1(3).Value Then
        rsrs!PREISKZ = 3
    ElseIf Option1(4).Value Then
        rsrs!PREISKZ = 4
    ElseIf Option1(5).Value Then
        rsrs!PREISKZ = 5
    ElseIf Option1(6).Value Then
        rsrs!PREISKZ = 6
    End If
    
    iStufe = 44
    cDBFeld = "FILIALNR"
    
    ctmp = Trim$(Text1(17).Text)
    If ctmp <> "" Then
        If IsNumeric(ctmp) Then
            rsrs!FILIALNR = ctmp
        Else
            rsrs!FILIALNR = Val(gcFilNr)
        End If
    Else
        rsrs!FILIALNR = Val(gcFilNr)
    End If
    iStufe = 45
    cDBFeld = "ANGELEGT"
    If bNeu Then
        rsrs!angelegt = Fix(Now)
        cDBFeld = "AENDER"
        rsrs!AENDER = "I"
    Else
        cDBFeld = "AENDER"
        rsrs!AENDER = "U"
    End If
    
    iStufe = 46
    
    '//Felder LastDate und LastTime
    rsrs!LASTDATE = DateValue(Now)
    rsrs!LASTTIME = TimeValue(Now)
    
    iStufe = 47
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
        
    If bLeereMaske = True Then
        LeereKundenMaskeWKL13
    End If
    
    Text1(5).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten. " & iStufe
    
    Fehlermeldung1
    
'    Resume Next
End Sub
Private Function SucheKundenWKL13() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim cFeld As String
    Dim cwhere As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim dWert As Double
    Dim iRet As Integer
    Dim lcount As Long
    Dim cKun As String
    Dim counter     As Long
    Dim lrow        As Long
    Dim dOSUM       As Double
    Dim iSterne As Integer
    Dim isterneAlt As Integer
    
    iRet = fnPruefeEingabeWKL13()
    
    SucheKundenWKL13 = False
    
    loeschNEW "TOPK" & srechnertab, gdBase
    
    If iRet <> 0 Then
        anzeige "rot", "Suchkriterium!", Label0(2)
'        MsgBox "Bitte mindestens ein Suchkriterium angeben!", vbCritical, "SUCHEN"
        Text1(0).SetFocus
        Exit Function
    End If
    Frame1.Visible = False
    
    anzeige "normal", "Kunden suchen...", Label0(2)
    
    picprogress.Visible = True
    txtStatus.Text = 20
    
    cSQL = "Select "
    
    cSQL = cSQL & " KUNDNR "
    cSQL = cSQL & " , KUERZEL"
    cSQL = cSQL & " , GESCHLECHT"
    cSQL = cSQL & " , FIRMA"
    cSQL = cSQL & " , TITEL"
    cSQL = cSQL & " , NAME"
    cSQL = cSQL & " , VORNAME"
    cSQL = cSQL & " , STRASSE"
    cSQL = cSQL & " , PLZ"
    cSQL = cSQL & " , STADT"
    cSQL = cSQL & " , TEL"
    cSQL = cSQL & " , FAXNR"
    cSQL = cSQL & " , MOBILTEL"
    cSQL = cSQL & " , EMAIL"
    cSQL = cSQL & " , ANREDE"
    cSQL = cSQL & " , MERKMAL"
    cSQL = cSQL & " , DATUM1"
    cSQL = cSQL & " , angelegt as DATUM2"
    cSQL = cSQL & " , FILIALNR"
    cSQL = cSQL & " , UMSLJ"
    cSQL = cSQL & " , UMSVJ"
    cSQL = cSQL & " , RABATT"
    cSQL = cSQL & " , GESPERRT"
    cSQL = cSQL & " , AWM"
    cSQL = cSQL & " , BONUS"
    cSQL = cSQL & " , BE"
    cSQL = cSQL & " , KUNDKART"
    cSQL = cSQL & " , 0.0 as OSUM"
    
 
    cSQL = cSQL & " , 'N' as OKB"
    cSQL = cSQL & " , 'J' as aktiv"
    cSQL = cSQL & " , 0 as Sterne"
    cSQL = cSQL & " , ECIDENT "
    cSQL = cSQL & " , KURZTEXT1 "
    
    cSQL = cSQL & "  into TOPK" & srechnertab & " from KUNDEN "
            

    cwhere = ""
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDNR = " & cFeld & " "
    End If
    
    cFeld = Text1(36).Text
    cFeld = Trim$(cFeld)
    cFeld = fnMoveComma2Point$(cFeld)
    
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & " and "
        End If
        cwhere = cwhere & " Bonus " & Command1(6).Caption & " " & cFeld & " "
    End If
    
    cFeld = Trim(Text1(37).Text)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & " and "
        End If
        cwhere = cwhere & " right(ECIDENT,10) like '*" & cFeld & "*' "
    End If
    
    cFeld = Text3.Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " angelegt >= " & CLng(DateValue(cFeld)) & " "
    End If
    
    cFeld = Text1(1).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUERZEL like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(2).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "NAME like '" & UCase$(cFeld) & "*' "
    End If
    
    cFeld = Text1(3).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "PLZ like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "STADT like '" & UCase$(cFeld) & "*' "
    End If
    
    cFeld = Label1(5).Tag
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "AWM = '" & cFeld & "' "
    End If
    
    If Combo3.Text <> "Rabatt wählen" Then
        
        cFeld = Combo3.Text
        If cFeld <> "" Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            cwhere = cwhere & " Rabatt = " & cFeld & " "
        End If
    End If
    
    If Combo1.Text <> "bitte wählen" Then
        Select Case Combo1.Text
            
            Case "Listen - EK"
                cFeld = "2"
            Case "Schnitt - EK"
                cFeld = "3"
            Case "Kassen - VK"
                cFeld = "0"
            Case "Listen - VK"
                cFeld = "1"
            Case "spez Kassen - VK"
                cFeld = "4"
            Case "Listen - VK mit Ausn."
                cFeld = "5"
            Case "Netto - KVK"
                cFeld = "6"
        End Select
        
        If cFeld <> "" Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            cwhere = cwhere & " Preiskz = " & cFeld & " "
        End If
    End If
    
    'BLZ
    cFeld = Trim(Left(Combo2.Text, 8))
    
    If cFeld <> "BLZ ausw" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & " and "
        End If
        cwhere = cwhere & " left(ECIDENT,8) like '" & cFeld & "*' "
    End If
    
    If cwhere = "" Then
        cwhere = "where "
    Else
        cwhere = cwhere & "and "
    End If
    cwhere = cwhere & " (STATUS <> 'D' or STATUS is null) "
   
    If Option2(0).Value Then
        cSQL = cSQL & cwhere & "order by NAME, VORNAME, STADT "
    ElseIf Option2(1).Value Then
        cSQL = cSQL & cwhere & "order by DATUM1, NAME, VORNAME, STADT "
    ElseIf Option2(2).Value Then
        cSQL = cSQL & cwhere & "order by DATUM2, NAME, VORNAME, STADT "
    ElseIf Option2(3).Value Then
        cSQL = cSQL & cwhere & "order by MERKMAL, NAME, VORNAME, STADT "
    Else
        cSQL = cSQL & cwhere & "order by KUERZEL "
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    SucheKundenWKL13 = True
    
    
    If gbSQLSERVER Then
    
        Dim lJahr As Long
        lJahr = Year(Now)
        cSQL = "Update TOPK" & srechnertab & " Set TOPK" & srechnertab & ".UMSLJ = "
        cSQL = cSQL & "(Select UMSATZJ from UMSKDJ where TOPK" & srechnertab & ".kundnr = UMSKDJ.kundnr and UMSKDJ.Jahr = " & Trim$(Str$(lJahr)) & ") "
        SQL_Befehl_ausführen cSQL
    
        cSQL = "Update TOPK" & srechnertab & " Set TOPK" & srechnertab & ".UMSVJ = "
        cSQL = cSQL & "(Select UMSATZJ from UMSKDJ where TOPK" & srechnertab & ".kundnr = UMSKDJ.kundnr and UMSKDJ.Jahr = " & Trim$(Str$(lJahr - 1)) & ") "
        SQL_Befehl_ausführen cSQL
        
        'offene KB
        
        'jemals gekauft
        
'        offene Kredite
        
        'Sterne
    
        'mit offenen KB einfärben
    
    Else
    
    
    
        lrow = 0
        counter = 0
        cSQL = "Select * from TOPK" & srechnertab
    
    '    OpenRecordset("Authors",dbOpenDynamic,0,dbOptimistic)
    '    Set rsrs = gdBase.OpenRecordset(cSQL, dbOpenDynamic, 0, dbOptimistic) 'öffnet
    
    
        Set rsrs = gdBase.OpenRecordset(cSQL, dbOpenDynaset)
        If Not rsrs.EOF Then
            rsrs.MoveLast
            counter = rsrs.RecordCount
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!Kundnr) Then
                    lrow = lrow + 1
                    txtStatus.Text = (lrow * 100) / counter
    
    
                    If Check2.Value = vbChecked Then
                        If OffeneKB(rsrs!Kundnr) Then
    
                            rsrs.Edit
                            rsrs!OKB = "J"
                            rsrs.Update
                        Else
                            If Check2.Value = vbChecked Then
                                rsrs.delete
                                GoTo hier
                            Else
                                rsrs.Edit
                                rsrs!OKB = "N"
                                rsrs.Update
    
                            End If
    
                        End If
                    End If
    
                    If Check3.Value = vbChecked Then
                        dOSUM = OffeneKredite(rsrs!Kundnr)
                        If dOSUM > 0 Then
    
                            rsrs.Edit
                            rsrs!OSUM = dOSUM
                            rsrs.Update
                        Else
                            If Check3.Value = vbChecked Then
                                rsrs.delete
                                GoTo hier
                            Else
                                rsrs.Edit
                                rsrs!OSUM = 0
                                rsrs.Update
    
                            End If
    
                        End If
                    End If
    
                    If Check4.Value = vbChecked Then
                        If JEMALSKAUF(rsrs!Kundnr) Then
    
                            rsrs.Edit
                            rsrs!Aktiv = "J"
                            rsrs.Update
                        Else
                            If Check4.Value = vbChecked Then
                                rsrs.delete
                                GoTo hier
                            Else
                                rsrs.Edit
                                rsrs!Aktiv = "N"
                                rsrs.Update
    
                            End If
    
                        End If
                    End If
    
                    If Check5.Value = vbChecked Then
                        If rsrs!BE = False Then
                            rsrs.delete
                            GoTo hier
                        End If
                    End If
    
                    rsrs.Edit
    
                    If gbSterne Then
                        AllesinKUNDAZE rsrs!Kundnr, False
    
                        iSterne = wievieleSterne(rsrs!Kundnr, "KUNDAZE", 180)
                        isterneAlt = wievieleSterneAlt
    
                        If isterneAlt > iSterne Then
    
                            Select Case iSterne
                                Case 1
                                    Select Case isterneAlt
                                        Case 2
                                            rsrs!Sterne = 7
                                        Case 3
                                            rsrs!Sterne = 8
                                        Case 4
                                            rsrs!Sterne = 9
                                        Case 5
                                            rsrs!Sterne = 10
                                    End Select
    
                                Case 2
    
                                    Select Case isterneAlt
    
                                        Case 3
                                            rsrs!Sterne = 12
                                        Case 4
                                            rsrs!Sterne = 13
                                        Case 5
                                            rsrs!Sterne = 14
                                    End Select
    
                                Case 3
                                    Select Case isterneAlt
                                        Case 4
                                            rsrs!Sterne = 16
                                        Case 5
                                            rsrs!Sterne = 17
                                    End Select
    
                                Case 4
                                    Select Case isterneAlt
                                        Case 5
                                            rsrs!Sterne = 19
                                    End Select
    
                                Case 5
    
                            End Select
    
    
                        Else
    
                            Select Case iSterne
                                Case 0
                                    rsrs!Sterne = 0
                                Case 1
                                    rsrs!Sterne = 6
                                Case 2
                                    rsrs!Sterne = 11
                                Case 3
                                    rsrs!Sterne = 15
                                Case 4
                                    rsrs!Sterne = 18
                                Case 5
                                    rsrs!Sterne = 20
                            End Select
    
    
    
                        End If
                    End If
    
                    rsrs!UMSVJ = umsatzVJJahr(rsrs!Kundnr)
                    rsrs!UMSLJ = umsatzLFJahr(rsrs!Kundnr)
                    rsrs.Update
    
                End If
hier:
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
        
    
    
    picprogress.Visible = False
    
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheKundenWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten." & rsrs!Kundnr

    Fehlermeldung1
    
'    Resume Next

End Function
Private Function SucheKundenMitSterne(iSternew As Integer) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cwhere As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim dWert As Double
    Dim iRet As Integer
    Dim lcount As Long
    Dim cKun As String
    Dim counter     As Long
    Dim lrow        As Long
    Dim dOSUM       As Double
    Dim iSterne As Integer
    Dim isterneAlt As Integer
    
    SucheKundenMitSterne = False
    
    loeschNEW "TOPK" & srechnertab, gdBase
    
    Frame1.Visible = False
    
    anzeige "normal", "Kunden suchen...", Label0(2)
    
    picprogress.Visible = True
    txtStatus.Text = 20
    
    cSQL = "Select "
    
    cSQL = cSQL & " KUNDEN.KUNDNR "
    cSQL = cSQL & " , KUNDEN.KUERZEL"
    cSQL = cSQL & " , KUNDEN.GESCHLECHT"
    cSQL = cSQL & " , KUNDEN.FIRMA"
    cSQL = cSQL & " , KUNDEN.TITEL"
    cSQL = cSQL & " , KUNDEN.NAME"
    cSQL = cSQL & " , KUNDEN.VORNAME"
    cSQL = cSQL & " , KUNDEN.STRASSE"
    cSQL = cSQL & " , KUNDEN.PLZ"
    cSQL = cSQL & " , KUNDEN.STADT"
    cSQL = cSQL & " , KUNDEN.TEL"
    cSQL = cSQL & " , KUNDEN.FAXNR"
    cSQL = cSQL & " , KUNDEN.MOBILTEL"
    cSQL = cSQL & " , KUNDEN.EMAIL"
    cSQL = cSQL & " , KUNDEN.ANREDE"
    cSQL = cSQL & " , KUNDEN.MERKMAL"
    cSQL = cSQL & " , KUNDEN.DATUM1"
    cSQL = cSQL & " , KUNDEN.DATUM2"
    cSQL = cSQL & " , KUNDEN.FILIALNR"
    cSQL = cSQL & " , KUNDEN.UMSLJ"
    cSQL = cSQL & " , KUNDEN.UMSVJ"
    cSQL = cSQL & " , KUNDEN.RABATT"
    cSQL = cSQL & " , KUNDEN.GESPERRT"
    cSQL = cSQL & " , KUNDEN.AWM"
    cSQL = cSQL & " , KUNDEN.BONUS"
    cSQL = cSQL & " , KUNDEN.BE"
    cSQL = cSQL & " , KUNDEN.KUNDKART"
    cSQL = cSQL & " , 0.0 as OSUM"
    cSQL = cSQL & " , 'N' as OKB"
    cSQL = cSQL & " , 'J' as aktiv"
    cSQL = cSQL & " , 0 as Sterne"
    cSQL = cSQL & " , ECIDENT "
    cSQL = cSQL & " , KURZTEXT1 "
    
    
    cSQL = cSQL & ", Sterne.UMSDURCHSCHNITT365  "
    cSQL = cSQL & ", Sterne.Verkvorgänge  "
    cSQL = cSQL & ", Sterne.BONDURCHSCHNITT365  "
    cSQL = cSQL & ", Sterne.AllesUMSDURCHSCHNITT365  "
    cSQL = cSQL & ", Sterne.AllesVerkvorgänge  "
    cSQL = cSQL & ", Sterne.AllesBONDURCHSCHNITT365  "
    
    
    
    
    
    cSQL = cSQL & " into TOPK" & srechnertab & " from KUNDEN inner join Sterne on "
    cSQL = cSQL & " KUNDEN.Kundnr = Sterne.Kundnr "
            
    
    
    
    cwhere = " where Sterne.Sterne = " & iSternew & " "
    
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.KUNDNR = " & cFeld & " "
    End If
    
    cFeld = Text3.Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " KUNDEN.angelegt >= " & CLng(DateValue(cFeld)) & " "
    End If
    
    cFeld = Text1(1).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.KUERZEL like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(2).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.NAME like '" & UCase$(cFeld) & "*' "
    End If
    
    cFeld = Text1(3).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.PLZ like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.STADT like '" & UCase$(cFeld) & "*' "
    End If
    
    cFeld = Label1(5).Tag
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUNDEN.AWM = '" & cFeld & "' "
    End If
    
    If cwhere = "" Then
        cwhere = "where "
    Else
        cwhere = cwhere & "and "
    End If
    cwhere = cwhere & " (STATUS <> 'D' or STATUS is null) "
   
    If Option2(0).Value Then
        cSQL = cSQL & cwhere & "order by NAME, VORNAME, STADT "
    ElseIf Option2(1).Value Then
        cSQL = cSQL & cwhere & "order by DATUM1, NAME, VORNAME, STADT "
    ElseIf Option2(2).Value Then
        cSQL = cSQL & cwhere & "order by DATUM2, NAME, VORNAME, STADT "
    ElseIf Option2(3).Value Then
        cSQL = cSQL & cwhere & "order by MERKMAL, NAME, VORNAME, STADT "
    Else
        cSQL = cSQL & cwhere & "order by KUERZEL "
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    SucheKundenMitSterne = True
    
    lrow = 0
    counter = 0
    Set rsrs = gdBase.OpenRecordset("TOPK" & srechnertab)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        counter = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kundnr) Then
                lrow = lrow + 1
                txtStatus.Text = (lrow * 100) / counter
                
                rsrs.Edit
                If iSternew = 5 Then
                    iSterne = 5
                    isterneAlt = 5
                Else
                    iSterne = wievieleSterneQuick(CLng(rsrs!Verkvorgänge), CDbl(rsrs!UMSDURCHSCHNITT365), CDbl(rsrs!BONDURCHSCHNITT365))
                    isterneAlt = wievieleSterneAltQuick(CLng(rsrs!allesVerkvorgänge), CDbl(rsrs!allesUMSDURCHSCHNITT365), CDbl(rsrs!allesBONDURCHSCHNITT365))
                End If
                
                If isterneAlt > iSterne Then
                    Select Case iSterne
                        Case 1
                            Select Case isterneAlt
                                Case 2
                                    rsrs!Sterne = 7
                                Case 3
                                    rsrs!Sterne = 8
                                Case 4
                                    rsrs!Sterne = 9
                                Case 5
                                    rsrs!Sterne = 10
                            End Select
                        Case 2
                            Select Case isterneAlt
                                Case 3
                                    rsrs!Sterne = 12
                                Case 4
                                    rsrs!Sterne = 13
                                Case 5
                                    rsrs!Sterne = 14
                            End Select
                        Case 3
                            Select Case isterneAlt
                                Case 4
                                    rsrs!Sterne = 16
                                Case 5
                                    rsrs!Sterne = 17
                            End Select
                        Case 4
                            Select Case isterneAlt
                                Case 5
                                    rsrs!Sterne = 19
                            End Select
                        Case 5
                    End Select
                Else
                    Select Case iSterne
                        Case 0
                            rsrs!Sterne = 0
                        Case 1
                            rsrs!Sterne = 6
                        Case 2
                            rsrs!Sterne = 11
                        Case 3
                            rsrs!Sterne = 15
                        Case 4
                            rsrs!Sterne = 18
                        Case 5
                            rsrs!Sterne = 20
                    End Select
                End If
                
                rsrs!UMSVJ = umsatzVJJahr(rsrs!Kundnr)
                rsrs!UMSLJ = umsatzLFJahr(rsrs!Kundnr)
                rsrs.Update
                
            End If
hier:
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    picprogress.Visible = False
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheKundenMitSterne"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function fnPruefeEingabeWKL13()
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long

    fnPruefeEingabeWKL13 = 1
    
    If Check2.Value = vbChecked Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Check3.Value = vbChecked Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Check4.Value = vbChecked Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Check5.Value = vbChecked Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If

    If Label1(5).Caption = "Farbauswahl" Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    
    
    If Combo3.Text <> "Rabatt wählen" Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Combo1.Text <> "bitte wählen" Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Combo2.Text <> "BLZ auswählen" Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If
    
    If Trim$(Text1(37).Text) <> "" Then
        fnPruefeEingabeWKL13 = 0
        Exit Function
    End If

    For lcount = 0 To 36
        If lcount = 0 Or lcount = 1 Or lcount = 2 Or lcount = 3 Or lcount = 4 Or lcount = 36 Then
            If Trim$(Text1(lcount).Text) <> "" Then
                fnPruefeEingabeWKL13 = 0
                Exit Function
            End If
        End If
    Next lcount
    
    If Trim$(Text3.Text) <> "" Then
        If IsDate(Text3.Text) Then
            fnPruefeEingabeWKL13 = 0
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Private Sub SwitchUpperLowerCaseWKL13()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If Left(SSCommand2(4).Caption, 1) = "A" Then
        For lcount = 22 To 32
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        SSCommand1(54).Caption = ","
        SSCommand1(55).Caption = "."
        SSCommand1(56).Caption = "-"
        
        SSCommand2(4).Caption = "a -> A"
    Else
        For lcount = 22 To 32
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        
        SSCommand1(54).Caption = ";"
        SSCommand1(55).Caption = ":"
        SSCommand1(56).Caption = "_"
        
        SSCommand2(4).Caption = "A -> a"
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SwitchUpperLowerCaseWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub cbo1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    cbo1.BackColor = glSelBack1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cbo1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cbo1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lWert As Long
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'Suchen
            glSelect = 0
            Text1(1).Text = UCase$(Text1(1).Text)
            If SucheKundenWKL13 Then
                zeige_Grid
                If MSFlexGrid1.Visible = True Then
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Row = 2
                    MSFlexGrid1.SetFocus
                End If
            End If
        Case Is = 1     'Neu
            LeereKundenMaskeWKL13
            lWert = fnHoleMaxKundenNr()
            Text1(5).Text = Format$(lWert, "#####0")
            Check6.Value = vbChecked
            Frame3.Visible = True
            Frame0.Visible = False
            giDlgZustand = giNEU
            
            Dim sZielpfad       As String
    
            sZielpfad = gcDBPfad
            If Right(sZielpfad, 1) <> "\" Then
                sZielpfad = sZielpfad & "\"
            End If
            sZielpfad = sZielpfad & "PICTURE\KUNDEN"
            
            If FileExists(sZielpfad & "\" & lWert & "_1.jpg") = True Then
                Kill sZielpfad & "\" & lWert & "_1.jpg"
            End If
            
            If FileExists(sZielpfad & "\" & lWert & "_2.jpg") = True Then
                Kill sZielpfad & "\" & lWert & "_2.jpg"
            End If
            
            If FileExists(sZielpfad & "\" & lWert & ".jpg") = True Then
                Kill sZielpfad & "\" & lWert & ".jpg"
            End If
            
        Case Is = 2     'Beenden
            gcKundenNr = ""
            Unload frmWKL13
        Case 3
            frmWKL120.Show 1
        Case 4
            Frame5.Visible = True
            fülleRab Combo3
        Case 5
            Frame5.Visible = False
        Case 6
            If Command1(6).Caption = "<" Then
                Command1(6).Caption = ">"
            Else
                Command1(6).Caption = "<"
            End If
        Case 7
        
            loeschNEW "KUBLZ", gdBase
            fülleBLZ Combo2
            
        Case Is = 11     'farbe
            Screen.MousePointer = 0

            gsBackcolor = Label1(5).BackColor
            gsForecolor = Label1(5).ForeColor
            gsKundenFarbe = Label1(5).Tag
            
            frmWKL65.Show 1
            
            Label1(5).BackColor = gsBackcolor
            Label1(5).ForeColor = gsForecolor
            Label1(5).Tag = gsKundenFarbe
            If gsKundenFarbe <> "" Then
                Label1(5).Caption = "Farbauswahl"
            Else
                Label1(5).Caption = "alle Farben"
            End If
            
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim cSuch As String
    Dim cMeld As String
    Dim iRet As Integer
    
    Select Case Index
        Case Is = 0     'Auswählen
        
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
            cSuch = Trim$(cSuch)
        
            giDlgZustand = giUPD
            If IsNumeric(cSuch) Then
                HoleDatenWKL13 cSuch
            End If
            glSelect = 0
        Case Is = 1     'Schließen
            Frame1.Visible = False
            Frame0.Visible = True
            Text1(0).Enabled = True
            Text1(1).Enabled = True
            Text1(2).Enabled = True
            Text1(3).Enabled = True
            Text1(4).Enabled = True
            
            If gbBILDTAST = False Then
                Frame2.Visible = False
            Else
                Frame2.Visible = True
            End If
            LeereDialogWKL13
            Text1(0).SetFocus
            
        Case Is = 2     'Löschen
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
            cSuch = Trim$(cSuch)
            
            If Val(cSuch) > 0 Then
                cMeld = "ACHTUNG!" & vbCrLf & vbCrLf
                cMeld = cMeld & "Kunde Nr. " & cSuch & " soll gelöscht werden." & vbCrLf & vbCrLf
                cMeld = cMeld & "Das Löschen eines Kunden kann zu Unstimmigkeiten" & vbCrLf
                cMeld = cMeld & "in der Datenbank führen, wenn der Kunde bereits" & vbCrLf
                cMeld = cMeld & "Ware gekauft hat!" & vbCrLf & vbCrLf
                cMeld = cMeld & "Wollen Sie den Kunden trotzdem löschen?"
                iRet = MsgBox(cMeld, vbYesNo + vbQuestion, "LÖSCHEN")
                If iRet = vbYes Then
                    LoescheKundeWKL13 cSuch
                End If
            End If
        Case 3
        
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            gckundnr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
            gckundnr = Trim$(gckundnr)
            gsARTNR = ""
            
            If Val(gckundnr) > 0 Then
                frmWKL74.Show 1
            End If
        Case 4
            frmWKL77.Show 1
            Exit Sub
        Case 5
        

            gckundnr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
            gckundnr = Trim$(gckundnr)
            
            If IsNumeric(gckundnr) Then
                    
                If IsNumeric(gckundnr) Then
                    frmWKL92.Show 1
                End If
            End If
            gckundnr = ""
            
        Case 6
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
            cSuch = Trim$(cSuch)
            
            If Val(cSuch) > 0 Then
                DatenschutzblattKundeDrucken cSuch
            End If

    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    Dim ctmp As String
    
    iZielIndex = Label3(2).Caption
    If iZielIndex < 999 Then
        If iZielIndex >= 500 Then
            SendKeys SSCommand1(Index).Caption
        Else
            If iZielIndex = 0 Or iZielIndex = 5 Then
                If Index >= 11 And Index <= 20 Then
                    Text1(iZielIndex).Text = Text1(iZielIndex).Text & SSCommand1(Index).Caption
                End If
            Else
                Text1(iZielIndex).Text = Text1(iZielIndex).Text & SSCommand1(Index).Caption
            End If
            Text1(iZielIndex).SetFocus
        End If
    Else
        Text2.Text = Text2.Text & SSCommand1(Index).Caption
        Text2.SetFocus
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Private Sub SSCommand1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    If Frame0.Visible Then
        Text1(iZielIndex).BackColor = glSelBack1
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub


Private Sub SSCommand2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    Dim lcount As Long
    
    iZielIndex = Label3(2).Caption
    
    If iZielIndex < 999 Then
        If iZielIndex < 500 Then
            Select Case Index
                Case Is = 0     'CLEAR
                    Text1(iZielIndex).Text = ""
                    
                Case Is = 1     'ENTER
                    If Len(Text1(iZielIndex).Text) > 0 Then
                        Text1(iZielIndex).Text = Left(Text1(iZielIndex).Text, Len(Text1(iZielIndex).Text) - 1)
                    End If
                Case Is = 2     'BEFORE
                    If Frame0.Visible Then
                        If iZielIndex > 0 Then
                            iZielIndex = iZielIndex - 1
                            Label3(1).Caption = Label1(iZielIndex).Caption
                            Label3(2).Caption = Trim$(Str$(iZielIndex))
                            Text1(iZielIndex).BackColor = glSelBack1
                        End If
                    End If
                    If Frame3.Visible Then
                        If iZielIndex > 5 Then
                            Select Case iZielIndex
                                Case 6 To 19
                                    iZielIndex = iZielIndex - 1
                                    If iZielIndex = 17 Then
                                        iZielIndex = 34
                                    ElseIf iZielIndex = 6 Then
                                        iZielIndex = 17
                                    ElseIf iZielIndex = 16 Then
                                        iZielIndex = 6
                                    End If
                                Case Is = 28
                                    iZielIndex = 19
                                Case Is = 34
                                    iZielIndex = 16
                                Case Is = 20
                                    iZielIndex = 28
                                Case Is = 21
                                    iZielIndex = iZielIndex - 1
                                Case Is = 22
                                    iZielIndex = iZielIndex - 1
                                Case Is = 23
                                    iZielIndex = iZielIndex - 1
                                Case Is = 24
                                    iZielIndex = iZielIndex - 1
                                Case Is = 500
                                    iZielIndex = 21
                                Case Is = 501
                                    iZielIndex = 500
                                Case 25 To 26
                                    iZielIndex = iZielIndex - 1
                                Case Is = 33
                                    iZielIndex = 26
                                Case Is = 29
                                    iZielIndex = 33
                                Case Is = 27
                                    iZielIndex = 30
                                Case Is = 31
                                    iZielIndex = 27
                                Case 30
                                    iZielIndex = iZielIndex - 1
                                Case 32
                                    iZielIndex = iZielIndex - 1
                            End Select
                            If iZielIndex < 500 Then
                                Label3(1).Caption = Label4(iZielIndex - 5).Caption
                                Label3(2).Caption = Trim$(Str$(iZielIndex))
                                Text1(iZielIndex).BackColor = glSelBack1
                            Else
                                Label3(1).Caption = "DATUM"
                                Label3(2).Caption = Trim$(Str$(iZielIndex))
'                                MaskEdBox1(iZielIndex - 500).BackColor = glSelBack1
                            End If
                        End If
                    End If
                    
                Case Is = 3     'NEXT
                    If Frame0.Visible Then
                        If iZielIndex < 4 Then
                            iZielIndex = iZielIndex + 1
                            Label3(1).Caption = Label1(iZielIndex).Caption
                            Label3(2).Caption = Trim$(Str$(iZielIndex))
                            Text1(iZielIndex).BackColor = glSelBack1
                        End If
                    End If
                    If Frame3.Visible Then
                        If iZielIndex < 35 Then
                            Select Case iZielIndex
                                Case 5 To 18
                                    iZielIndex = iZielIndex + 1
                                    If iZielIndex = 17 Then
                                        iZielIndex = 34
                                    ElseIf iZielIndex = 18 Then
                                        iZielIndex = 7
                                    ElseIf iZielIndex = 7 Then
                                        iZielIndex = 17
                                    End If
                                Case Is = 34
                                    iZielIndex = 18
                                Case Is = 19
                                    iZielIndex = 28
                                Case Is = 28
                                    iZielIndex = 20
                                Case Is = 500
                                    iZielIndex = iZielIndex + 1
                                Case Is = 501
                                    iZielIndex = 24
                                Case 20 To 25
                                    iZielIndex = iZielIndex + 1
                                Case Is = 26
                                    iZielIndex = 33
                                Case Is = 27
                                    iZielIndex = 31
                                Case Is = 33
                                    iZielIndex = 29
                                Case Is = 30
                                    iZielIndex = 27
                                Case 29
                                    iZielIndex = iZielIndex + 1
                                Case 31
                                    iZielIndex = iZielIndex + 1
                                Case Is = 32
                                    iZielIndex = 32
                            End Select
                            If iZielIndex < 500 Then
                                Label3(1).Caption = Label4(iZielIndex - 5).Caption
                                Label3(2).Caption = Trim$(Str$(iZielIndex))
                                Text1(iZielIndex).BackColor = glSelBack1
                            Else
                                Label3(1).Caption = "DATUM"
                                Label3(2).Caption = Trim$(Str$(iZielIndex))
'                                MaskEdBox1(iZielIndex - 500).BackColor = glSelBack1
                            End If
                        End If
                    End If
                    
                Case Is = 4     'switch UPPER CASE / lower case
                    SwitchUpperLowerCaseWKL13
                    
                Case Is = 5     'switch to lower case
                    For lcount = 0 To 25
                        SSCommand1(lcount).Caption = Chr$(97 + lcount)
                    Next lcount
                    SSCommand1(26).Caption = "ä"
                    SSCommand1(27).Caption = "ö"
                    SSCommand1(28).Caption = "ü"
                    SSCommand1(29).Caption = "ß"
                    
                Case Is = 7
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    For lcount = 5 To 30
'                        If lcount = 17 Then
'
'                        Else
                            Text1(lcount).Enabled = True
'                        End If
                    Next lcount
                    Frame2.Enabled = True
                    Frame11.Visible = False
                    Frame4.Visible = False
            End Select
            If iZielIndex < 500 Then
                Text1(iZielIndex).SetFocus
            Else

            End If
        Else
            Select Case Index
                Case 0, 1     'CLEAR = BACKSPACE
        
                
                Case Is = 2     'BEFORE
                    If iZielIndex = 501 Then
                        iZielIndex = iZielIndex - 1
                        Label3(1).Caption = "DATUM"
                        Label3(2).Caption = Trim$(Str$(iZielIndex))

                    Else

                        iZielIndex = 21
                        Label3(1).Caption = Label4(iZielIndex - 5).Caption
                        Label3(2).Caption = Trim$(Str$(iZielIndex))
                        Text1(iZielIndex).BackColor = glSelBack1
                        Text1(iZielIndex).SetFocus
                    End If
                    
                Case Is = 3     'NEXT
                    If iZielIndex = 500 Then
                        iZielIndex = iZielIndex + 1
                        Label3(1).Caption = "DATUM"
                        Label3(2).Caption = Trim$(Str$(iZielIndex))

                    Else
                        iZielIndex = 24
                        Label3(1).Caption = Label4(iZielIndex - 5).Caption
                        Label3(2).Caption = Trim$(Str$(iZielIndex))
                        Text1(iZielIndex).BackColor = glSelBack1
                        Text1(iZielIndex).SetFocus
                    End If
                    
                Case Is = 4     'switch UPPER CASE / lower case
                    SwitchUpperLowerCaseWKL13
                    
                Case Is = 5     'switch to lower case
                    For lcount = 0 To 25
                        SSCommand1(lcount).Caption = Chr$(97 + lcount)
                    Next lcount
                    SSCommand1(26).Caption = "ä"
                    SSCommand1(27).Caption = "ö"
                    SSCommand1(28).Caption = "ü"
                    SSCommand1(29).Caption = "ß"
                    
                Case Is = 7
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    For lcount = 5 To 30
'                        If lcount = 17 Then
'
'                        Else
                            Text1(lcount).Enabled = True
'                        End If
                    Next lcount
                    Frame2.Enabled = True
                    Frame11.Visible = False
                    Frame4.Visible = False
            End Select
            
        
        End If
    Else
        Select Case Index
            Case Is = 0     'CLEAR
                Text2.Text = ""
                
            Case Is = 1     'BACKSPACE
                If Len(Text2.Text) > 0 Then
                    Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
                End If
    
            
            Case Is = 2     'BEFORE
                
            Case Is = 3     'NEXT
                
            Case Is = 4     'switch UPPER CASE / lower case
                SwitchUpperLowerCaseWKL13
                
            Case Is = 7
                Command5(0).Enabled = True
                Command5(1).Enabled = True
                Command5(2).Enabled = True
                Command5(3).Enabled = True
                For lcount = 5 To 30
'                    If lcount = 17 Then
'
'                    Else
                        Text1(lcount).Enabled = True
'                    End If
                Next lcount
                Frame2.Enabled = True
                Frame11.Visible = False
                Frame4.Visible = False
                
        End Select
                
        'Text2.SetFocus
    
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten. " & Index
    
    Fehlermeldung1
End Sub

Private Sub SSCommand2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    If iZielIndex < 999 Then
        If iZielIndex < 0 Then
            iZielIndex = 0
        End If
        If iZielIndex >= 500 Then
'            MaskEdBox1(iZielIndex - 500).BackColor = glSelBack1
        Else
            Text1(iZielIndex).BackColor = glSelBack1
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim i As Integer
    Dim cPos(1 To 5) As Long
    Dim lRet As Long
    Dim lcount As Long
    Dim cKdnr As String
    Dim bUnload As Boolean
    Dim bGespeichert As Boolean
    Dim ctmp As String
    Dim cTxT As String
    Dim cLBSatz As String
    Dim sSQL As String
    
    bGespeichert = False
    
    Select Case Index
        Case Is = 0     'Speichern
            
            '//firma
            Text1(8).Text = SwapStr(Text1(8).Text, "'", " ")
            Text1(8).Text = SwapStr(Text1(8).Text, "*", " ")
            
            '//Name
            Text1(10).Text = SwapStr(Text1(10).Text, "'", " ")
            Text1(10).Text = SwapStr(Text1(10).Text, "*", " ")
            '//Vorname
            Text1(11).Text = SwapStr(Text1(11).Text, "'", " ")
            Text1(11).Text = SwapStr(Text1(11).Text, "*", " ")
            '//Kuerzel
            Text1(6).Text = SwapStr(Text1(6).Text, "'", " ")
            Text1(6).Text = SwapStr(Text1(6).Text, "*", " ")
            
            
            lRet = fnPruefeDialogEingabenWKL13()
            If lRet = 0 Then
                If Trim$(Text1(6).Text) = "" Then
                    If Trim$(Text1(10).Text) <> "" Then
                        Text1(6).Text = UCase$(Left(Text1(10).Text, 5))
                        Command5_Click 0
                    Else
                        MsgBox "Bitte geben Sie einen Namen ein!", vbInformation, "Winkiss Hinweis:"
                        Text1(10).SetFocus
                    End If
                Else
                    Text1(6).Text = UCase$(Text1(6).Text)
                    gcKundenNr = Trim$(Text1(5).Text)
                    SchreibeDatenWKL13 True
                    bGespeichert = True
                    If iKasse = 1 Then 'komme aus Kasse Kunde neu
                        iKasse = 0
                        

                        Unload Me

                    ElseIf iKasse = 2 Then 'komme aus Kasse Daten ansehen
                        iKasse = 0
                        giDlgZustand = giUPD
                        Screen.MousePointer = 0
                        Unload Me
                    ElseIf iKasse = 3 Then 'komme aus Termine Daten ansehen
                        iKasse = 0
                        giDlgZustand = giUPD
                        Screen.MousePointer = 0
                        Unload Me
                    Else
                        Command5_Click 1
                    End If
                End If
            Else
                If lRet = 999 Then
                    ctmp = "Die Kunden-Nr liegt außerhalb Ihres zugeordeneten Nummernkreises!" & vbCrLf & vbCrLf
                    ctmp = ctmp & "In der Filial-Version von WinKISS gelten für die Vergabe" & vbCrLf
                    ctmp = ctmp & "folgende Regeln:" & vbCrLf & vbCrLf
                    ctmp = ctmp & "1. Die Kundennummer ist immer sechsstellig " & vbCrLf
                    ctmp = ctmp & "2. Die erste Ziffer der Kundennummer entspricht der Filialnummer " & vbCrLf & vbCrLf
                    lRet = fnHoleMaxKundenNr()
                    ctmp = ctmp & "Die nächste freie Kundennummer ist " & Format$(lRet, "#####0")
                    MsgBox ctmp, vbCritical, "STOP!"
                    Text1(5).SetFocus
                Else
                    MsgBox "Das Feld enthält ein ungültiges Zeichen!", vbCritical, "STOP!"
                    Text1(lRet).SetFocus
                End If
            End If
            
            
        Case Is = 1     'Schließen
        
            Timer1.Enabled = False
        
            bUnload = True
            If Trim$(Text1(6).Text) <> "" Then
                bUnload = True
                If iKasse = 2 Or iKasse = 1 Or iKasse = 3 Then
                    iKasse = 0
                    gcKundenNr = ""
                    Screen.MousePointer = 0
                    Unload Me
                    Exit Sub
                End If
            Else
                bUnload = True
                If iKasse = 2 Or iKasse = 1 Or iKasse = 3 Then
                    iKasse = 0
                    gcKundenNr = ""
                    Screen.MousePointer = 0
                    Unload Me
                    Exit Sub
                End If
            End If
            
            If bUnload = True Then
                If giDlgZustand = giNEU Then
                    Frame0.Visible = True
                    If gbBILDTAST = False Then
                        Frame2.Visible = False
                    Else
                        Frame2.Visible = True
                    End If
                    
                    Frame3.Visible = False
                    Frame1.Visible = False
                    
                    Text1(0).Enabled = True
                    Text1(1).Enabled = True
                    Text1(2).Enabled = True
                    Text1(3).Enabled = True
                    Text1(4).Enabled = True
                Else
                    Frame1.Visible = True
                    Frame3.Visible = False
                    Frame2.Visible = False
                    
                    If CInt(gcFilNr) > 0 Then
                        Command2(2).Enabled = False
                    Else
                        Command2(2).Enabled = True
                    End If
                    
                    If bGespeichert Then
                        Command1_Click 0
                    End If
                    If MSFlexGrid1.Visible = False Then
                        Frame0.Visible = True
                        If gbBILDTAST = False Then
                            Frame2.Visible = False
                        Else
                            Frame2.Visible = True
                        End If
'                        Frame2.Visible = True
                        Text1(0).Enabled = True
                        Text1(1).Enabled = True
                        Text1(2).Enabled = True
                        Text1(3).Enabled = True
                        Text1(4).Enabled = True
                    End If
                End If
                giDlgZustand = giUPD
                
            End If
        Case Is = 2     'Historie
            gckundnr = Text1(5).Text
            gckundnr = Trim$(gckundnr)
            gsARTNR = ""
            
            If gckundnr <> "" Then
                frmWKL74.Show 1
            End If
            
        Case Is = 3     'Notizen
            Frame4.Visible = True
            Frame3.Enabled = False
            Text2.SetFocus
            
        Case Is = 4     'Drucken
            cKdnr = Text1(5).Text
            cKdnr = Trim$(cKdnr)
            
            If cKdnr = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            SchreibeDatenWKL13 False
            
            If gcListenDrucker <> gcBonDrucker Then
                StammdatenblattKundeDrucken cKdnr, False, Text1(26).Text
            Else
                'auf Bondrucker
                StammdatenblattKundeDrucken_Bondrucker cKdnr
            End If
            
        Case 8 'offene Kundenbestellungen
            KB "INBESTELLUNG", ""
        Case 5 'Bestellte Kundenbestellungen
            KB "BESTELLT", ""
        Case 6 'GELIEFERT Kundenbestellungen
            KB "GELIEFERT", ""
        Case 7 'NICHTGELIEFERT Kundenbestellungen
            KB "NICHTGELIEFERT", ""
        Case 9 'nur nach Sternchen suchen
            Sternchensuche
        Case 11, 12, 13, 14, 15
            If SucheKundenMitSterne(Index - 10) Then
                zeige_Grid
                If MSFlexGrid1.Visible = True Then
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Row = 2
                    MSFlexGrid1.SetFocus
                End If
            End If
        Case 10 'Notizen drucken
        
            cKdnr = Text1(5).Text
            cKdnr = Trim$(cKdnr)
            
            If cKdnr = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        
            
            
            sSQL = "Update kunden set notizen = '" & Text2.Text & "'"
            sSQL = sSQL & " , Status = 'E' "
            sSQL = sSQL & " , SynStatus = 'E' "
            sSQL = sSQL & " where kundnr = " & Trim$(Text1(5).Text)
            gdBase.Execute sSQL, dbFailOnError
            
            
            StammdatenblattKundeDrucken cKdnr, True, Text1(26).Text
        Case 16 'gelöschte Kunden drucken
            If Datendrin("KUNDEDEL", gdBase) = True Then
        
                reportbildschirm "dWKL001b", "aWKL13c"
                
            Else
                MsgBox "Keine Daten enthalten", vbInformation, "Winkiss Hinweis:"
            End If
            
        Case 17
        
            cKdnr = Text1(5).Text
            cKdnr = Trim$(cKdnr)
            
            If cKdnr = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            DatenschutzblattKundeDrucken cKdnr
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Del_Kunden_Bild(sKund As String)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim i               As Integer
    Dim sBildname       As String
    
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim lRet            As String
    Dim lfail            As Long
    
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\KUNDEN"
    
    
    
    Kill sPfad & "\" & sKund
    
    Dim sKUNDNR As String
    sKUNDNR = Left(sKund, Len(sKund) - 4)
    
    File1.Path = sPfad
    File1.Pattern = sKUNDNR & "*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sBildname = File1.list(i)
            
            cQuelle = sPfad & "\"
            cQuelle = ShortPath(cQuelle)
            cQuelle = cQuelle & sBildname
    
            cZiel = sPfad & "\"
            cZiel = ShortPath(cZiel)
            If i > 0 Then
                cZiel = cZiel & "del_" & i & ".jpg"
            Else
                cZiel = cZiel & "del.jpg"
            End If
    
            lRet = CopyFile(cQuelle, cZiel, lfail)
        Next i
    End If
    
    Kill sPfad & "\" & sKUNDNR & "*.jpg"
    
    File1.Path = sPfad
    File1.Pattern = "del*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sBildname = File1.list(i)
            
            cQuelle = sPfad & "\"
            cQuelle = ShortPath(cQuelle)
            cQuelle = cQuelle & sBildname
    
            cZiel = sPfad & "\"
            cZiel = ShortPath(cZiel)
            If i > 0 Then
                cZiel = cZiel & sKUNDNR & "_" & i & ".jpg"
            Else
                cZiel = cZiel & sKUNDNR & ".jpg"
            End If
    
            lRet = CopyFile(cQuelle, cZiel, lfail)
        Next i
    End If
    
    Kill sPfad & "\del*.jpg"
        
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Del_Kunden_Bild"
        Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Kunden_Pic_hinzu(sKUNDNR As String)
On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sBilddatei  As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim sPfad       As String
    
    cZiel = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If

    sPfad = sPfad & "PICTURE\KUNDEN"
    sTitle = "Speichern der Bild - Datei"
    sFilter = "JPEG (*.JPG)| *.JPG|GIF (*.GIF)| *.GIF|PNG (*.PNG)| *.PNG| Bitmapdateien (*.bmp)|*.bmp"
    sBilddatei = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    cQuelle = sBilddatei
    cQuelle = ShortPath(cQuelle)

    cZiel = gcDBPfad
    If Right(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    cZiel = ShortPath(cZiel)
    
    cZiel = cZiel & "PICTURE\KUNDEN"
    
    'multibild
    File1.Path = cZiel
    File1.Pattern = sKUNDNR & "*.jpg"
    File1.Refresh
                
    If File1.ListCount > 0 Then
        cZiel = cZiel & "\" & sKUNDNR & "_" & File1.ListCount & ".jpg"
    Else
        cZiel = cZiel & "\" & sKUNDNR & ".jpg"
    End If
    
    lRet = CopyFile(cQuelle, cZiel, lfail)
    
    Zeige_Kunden_Bilder sKUNDNR, 1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kunden_Pic_hinzu"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Zeige_Kunden_Bilder(sKUNDNR As String, iStatus As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim sSpeicherpfad   As String
    Dim i               As Integer
    Dim sBildname       As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\KUNDEN"
    
    For i = 0 To 2
        Picture5(i).Enabled = False
        Picture5(i).Visible = False
    Next i
    
    If FileExists(sPfad & "\" & sKUNDNR & ".jpg") Then
        Image1.Picture = LoadPicture(sPfad & "\" & sKUNDNR & ".jpg")
        Image2.Picture = LoadPicture(sPfad & "\" & sKUNDNR & ".jpg")
        sSpeicherpfad = sPfad & "\" & sKUNDNR & "_s.jpg"
        Label5(12).Caption = sKUNDNR & ".jpg"
        Label4(38).Visible = True
        
        Picture3.Visible = True
        
        File1.Path = sPfad
        File1.Pattern = sKUNDNR & "*.jpg"
        File1.Refresh
                    
        If File1.ListCount > 1 Then
            'dann zeige weitere
            For i = 0 To File1.ListCount - 1
        
                If i > 2 Then
                    Exit For
                End If
                sBildname = File1.list(i)
                Picture5(i).Tag = sBildname
                Zeige_weitere_Kunden_Bilder sBildname, 0, Picture5(i), Image4, 50
                
                If i > 0 Then
                    Picture5(i).Top = Picture5(0).Top
                    Picture5(i).Left = Picture5(i - 1).Left + 50 + Picture5(i - 1).Width
                End If
            Next i
        End If
    Else
        sSpeicherpfad = ""
        Dim sGeschlecht As String
    
        sGeschlecht = lookingForKundendaten(sKUNDNR).geschlecht
        
        If UCase(sGeschlecht) = "W" Then
            sBildname = "keinBildw.jpg"
        ElseIf UCase(sGeschlecht) = "M" Then
            sBildname = "keinBildm.jpg"
        Else
            Dim sAnrede As String
    
            sAnrede = lookingForKundendaten(sKUNDNR).anrede
            
            If UCase(sAnrede) = "FRAU" Then
                sBildname = "keinBildw.jpg"
            ElseIf UCase(sAnrede) = "HERR" Then
                sBildname = "keinBildm.jpg"
            Else
                sBildname = "keinBildw.jpg"
            End If
        End If
        
        If FileExists(sPfad & "\" & sBildname) Then
            Image1.Picture = LoadPicture(sPfad & "\" & sBildname)
            Image2.Picture = LoadPicture(sPfad & "\" & sBildname)
            Label4(38).Visible = False
            Label5(12).Caption = ""
            Picture3.Visible = True
        Else
            Picture3.Visible = False
            Label5(12).Caption = ""
            
            Exit Sub
        End If
    End If
    
    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 250
        Case Is > 12000
            iDiv = 200
        Case Is > 11000
            iDiv = 150
        Case Is > 10000
            iDiv = 120
        Case Is > 8000
            iDiv = 100
    End Select
    
    Picture3.Tag = sKUNDNR & ".jpg"
    zeigImage_Kunden_In_Picture Image1, Picture3, iDiv, "" 'sSpeicherpfad

    Picture3.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
        Kill sPfad & "\" & sKUNDNR & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Zeige_Kunden_Bilder"
        Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub Zeige_weitere_Kunden_Bilder(sBild As String, iStatus As Integer, PicX As PictureBox, imx As Image, iDiv As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    
    If sBild = "" Then
        Exit Sub
    End If
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\KUNDEN"
    
    imx.Picture = LoadPicture("")
    imx.Refresh
    PicX.Picture = LoadPicture("")
    PicX.Refresh
    
    If FileExists(sPfad & "\" & sBild) Then
        imx.Picture = LoadPicture(sPfad & "\" & sBild)
        
        PicX.Visible = True
        PicX.Enabled = True
    End If
    
    zeigImage_Kunden_In_Picture imx, PicX, iDiv, ""
    PicX.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 481 Then
        If iStatus = 1 Then
            MsgBox "Dieses Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        End If
'        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Zeige_weitere_Kunden_Bilder"
        Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub zeigImage_Kunden_In_Picture(imgx As Image, PicX As PictureBox, iDiv As Integer, sSpeicherpfad As String)
    On Error GoTo LOKAL_ERROR

    Dim höhe As Integer
    Dim Breite As Integer
    Dim iTeiler As Integer

    If imgx.Width >= imgx.Height Then
        iTeiler = imgx.Width / iDiv
    Else
        iTeiler = imgx.Height / iDiv
    End If
    
    höhe = imgx.Height / iTeiler
    Breite = imgx.Width / iTeiler
    
    imgx.Height = höhe * Screen.TwipsPerPixelX
    imgx.Width = Breite * Screen.TwipsPerPixelY
    
    PicX.Picture = LoadPicture("")
    PicX.Refresh
    
    With PicX
        .BorderStyle = 0
        .Width = imgx.Width
        .Height = imgx.Height
        
        ' Wichtig: AutoRedraw = True
        .AutoRedraw = True
        
        ' Bild aus ImageBox übertragen
        .PaintPicture imgx.Picture, 0, 0, _
        imgx.Width, imgx.Height
      
        ' Bild abspeichern
        If sSpeicherpfad <> "" Then
            SavePicture .Image, sSpeicherpfad
        End If
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigImage_Kunden_In_Picture"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Sternchensuche()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim ctmp As String
    Dim rsrs As Recordset
    Dim lBon As Long
    Dim iSterne As Integer
    Dim isterneAlt As Integer
    Dim lcount As Long
    
    picprogress.Visible = True
    txtStatus.Text = "3"
    
    Screen.MousePointer = 11
    
    loeschNEW "KUNDAZEALLE", gdBase
    CreateTableT2 "KUNDAZEALLE", gdBase
    
    sSQL = "Insert into KUNDAZEALLE select artnr,menge,adate,Filiale,preis,bediener as bednr,kundnr from kassjour "
    sSQL = sSQL & " where kundnr > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Insert into KUNDAZEALLE select artnr,menge,adate,Filiale,preis,bednr,kundnr from kundkass "
    sSQL = sSQL & " where kundnr > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    CheckIndex "KUNDAZEALLE", "Kundnr", "", gdBase
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    CheckIndex "KUNDAZEALLE", "preis", "", gdBase
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    CheckIndex "KUNDAZEALLE", "adate", "", gdBase
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    
    loeschNEW "STERNTEMP1", gdBase
    CreateTableT2 "STERNTEMP1", gdBase
    
    loeschNEW "STERNTEMPZ", gdBase
    CreateTableT2 "STERNTEMPZ", gdBase
    
    sSQL = "Insert into STERNTEMPZ select kundnr, 0 as BONDURCHSCHNITT365,0 as UMSDURCHSCHNITT365,0 as verkvorgänge "
    sSQL = sSQL & ", 0 as AllesBONDURCHSCHNITT365,0 as AllesUMSDURCHSCHNITT365,0 as Allesverkvorgänge"
    sSQL = sSQL & " from KUNDAZEALLE where adate > " & CLng(DateValue(Now) - 180)
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2

    sSQL = "Insert into STERNTEMP1 select kundnr, sum(preis) as UMSDURCHSCHNITT365  from KUNDAZEALLE"
    sSQL = sSQL & " where adate > " & CLng(DateValue(Now) - 180)
    sSQL = sSQL & " group by kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMPZ inner join STERNTEMP1 on STERNTEMPZ.kundnr = STERNTEMP1.kundnr"
    sSQL = sSQL & " set STERNTEMPZ.UMSDURCHSCHNITT365 = STERNTEMP1.UMSDURCHSCHNITT365"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMP6", gdBase
    CreateTableT2 "STERNTEMP6", gdBase
    
    sSQL = "Insert into STERNTEMP6 select kundnr, sum(preis) as AllesUMSDURCHSCHNITT365  from KUNDAZEALLE"
    sSQL = sSQL & " group by kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMPZ inner join STERNTEMP6 on STERNTEMPZ.kundnr = STERNTEMP6.kundnr"
    sSQL = sSQL & " set STERNTEMPZ.AllesUMSDURCHSCHNITT365 = STERNTEMP6.AllesUMSDURCHSCHNITT365"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMP2", gdBase
    CreateTableT2 "STERNTEMP2", gdBase

    sSQL = "Insert into STERNTEMP2 select * from STERNTEMP1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMPT", gdBase
    CreateTableT2 "STERNTEMPT", gdBase

    sSQL = "Insert into STERNTEMPT select distinct(kundnr) as kund,adate  from KUNDAZEALLE"
    sSQL = sSQL & " where adate > " & CLng(DateValue(Now) - 180)
    sSQL = sSQL & " group by adate,kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMP3", gdBase
    CreateTableT2 "STERNTEMP3", gdBase

    sSQL = "Insert into STERNTEMP3 select distinct(kund) as kundnr,count(adate) as verkvorgänge  from STERNTEMPT"
    sSQL = sSQL & " where adate > " & CLng(DateValue(Now) - 180)
    sSQL = sSQL & " group by kund "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMPZ inner join STERNTEMP3 on STERNTEMPZ.kundnr = STERNTEMP3.kundnr"
    sSQL = sSQL & " set STERNTEMPZ.verkvorgänge = STERNTEMP3.verkvorgänge"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    
    loeschNEW "STERNTEMPT", gdBase
    CreateTableT2 "STERNTEMPT", gdBase

    sSQL = "Insert into STERNTEMPT select distinct(kundnr) as kund,adate  from KUNDAZEALLE"
    sSQL = sSQL & " group by adate,kundnr "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMP7", gdBase
    CreateTableT2 "STERNTEMP7", gdBase

    sSQL = "Insert into STERNTEMP7 select distinct(kund) as kundnr,count(adate) as Allesverkvorgänge  from STERNTEMPT"
    sSQL = sSQL & " group by kund "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMPZ inner join STERNTEMP7 on STERNTEMPZ.kundnr = STERNTEMP7.kundnr"
    sSQL = sSQL & " set STERNTEMPZ.Allesverkvorgänge = STERNTEMP7.Allesverkvorgänge"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNTEMP4", gdBase
    CreateTableT2 "STERNTEMP4", gdBase
    
    sSQL = "Insert into STERNTEMP4 select * from STERNTEMP1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMP4 inner join STERNTEMP3 on STERNTEMP4.kundnr = STERNTEMP3.kundnr "
    sSQL = sSQL & "  set STERNTEMP4.verkvorgänge = STERNTEMP3.verkvorgänge"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMP4 set BONDURCHSCHNITT365 = UMSDURCHSCHNITT365 / verkvorgänge"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNTEMPZ inner join STERNTEMP4 on STERNTEMPZ.kundnr = STERNTEMP4.kundnr"
    sSQL = sSQL & " set STERNTEMPZ.BONDURCHSCHNITT365 = STERNTEMP4.BONDURCHSCHNITT365"
    gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    
    loeschNEW "STERNTEMP5", gdBase
    CreateTableT2 "STERNTEMP5", gdBase
    
    sSQL = "Insert into STERNTEMP5 select * from STERNTEMP4 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    'Umsatz erreicht bleibt
    
    sSQL = "Delete from  STERNTEMP1 where UMSDURCHSCHNITT365 <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    ctmp = fnMoveComma2Point$(CStr(gdUmsatzMittelproKunde))
    sSQL = "Delete from  STERNTEMP1 where UMSDURCHSCHNITT365 < " & ctmp
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Delete from  STERNTEMP2 where UMSDURCHSCHNITT365 <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2

    ctmp = fnMoveComma2Point$(CStr(2 * gdUmsatzMittelproKunde))

    sSQL = "Delete from  STERNTEMP2 where  UMSDURCHSCHNITT365 < " & ctmp
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    'bonUmsatz erreicht bleibt
    
    sSQL = "Delete from  STERNTEMP4 where BONDURCHSCHNITT365 <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    ctmp = fnMoveComma2Point$(CStr(gdUmsatzproKundeDurchschnitt))
    sSQL = "Delete from  STERNTEMP4 where BONDURCHSCHNITT365 < " & ctmp
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Delete from  STERNTEMP5 where BONDURCHSCHNITT365 <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    ctmp = fnMoveComma2Point$(CStr(2 * gdUmsatzproKundeDurchschnitt))
    sSQL = "Delete from  STERNTEMP5 where BONDURCHSCHNITT365 < " & ctmp
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    'verkäufe erreicht bleibt
    ctmp = fnMoveComma2Point$(CStr(gdKaufvorgänge))
    
    sSQL = "Delete from  STERNTEMP3 where verkvorgänge < " & ctmp
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    loeschNEW "STERNE", gdBase
    CreateTableT2 "STERNE", gdBase
    
    sSQL = "Insert into STERNE select KUNDNR, 0 as STERNE from KUNDEN "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNE inner join STERNTEMP1 on STERNE.kundnr = STERNTEMP1.kundnr "
    sSQL = sSQL & "  set STERNE.STERNE = STERNE.STERNE + 1"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNE inner join STERNTEMP2 on STERNE.kundnr = STERNTEMP2.kundnr "
    sSQL = sSQL & "  set STERNE.STERNE = STERNE.STERNE + 1"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNE inner join STERNTEMP3 on STERNE.kundnr = STERNTEMP3.kundnr "
    sSQL = sSQL & "  set STERNE.STERNE = STERNE.STERNE + 1"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNE inner join STERNTEMP4 on STERNE.kundnr = STERNTEMP4.kundnr "
    sSQL = sSQL & "  set STERNE.STERNE = STERNE.STERNE + 1"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "Update STERNE inner join STERNTEMP5 on STERNE.kundnr = STERNTEMP5.kundnr "
    sSQL = sSQL & "  set STERNE.STERNE = STERNE.STERNE + 1"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    sSQL = "Update STERNTEMPZ set AllesBONDURCHSCHNITT365 = AllesUMSDURCHSCHNITT365 / allesverkvorgänge "
    sSQL = sSQL & " where allesverkvorgänge <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    sSQL = "Update STERNE inner join STERNTEMPZ on STERNE.kundnr = STERNTEMPZ.kundnr "
    sSQL = sSQL & "  set STERNE.AllesBONDURCHSCHNITT365 = STERNTEMPZ.AllesBONDURCHSCHNITT365 "
    sSQL = sSQL & "  , STERNE.AllesUMSDURCHSCHNITT365 = STERNTEMPZ.AllesUMSDURCHSCHNITT365 "
    sSQL = sSQL & "  , STERNE.UMSDURCHSCHNITT365 = STERNTEMPZ.UMSDURCHSCHNITT365 "
    sSQL = sSQL & "  , STERNE.BONDURCHSCHNITT365 = STERNTEMPZ.BONDURCHSCHNITT365 "
    sSQL = sSQL & "  , STERNE.allesverkvorgänge = STERNTEMPZ.allesverkvorgänge "
    sSQL = sSQL & "  , STERNE.verkvorgänge = STERNTEMPZ.verkvorgänge "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = CInt(txtStatus.Text) + 1

    loeschNEW "STERNEZAHL", gdBase
    CreateTableT2 "STERNEZAHL", gdBase
    
    sSQL = "Insert into STERNEZAHL (adate) values ('" & DateValue(Now) & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    Dim i As Integer
    For i = 1 To 5
        txtStatus.Text = CInt(txtStatus.Text) + 2
        sSQL = "Update STERNEZAHL set STERNE" & i & " = " & sternanzahlausSterne(i)
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    txtStatus.Text = CInt(txtStatus.Text) + 2
    
    sSQL = "select top 1 adate from STERNEZAHL "
    sSQL = sSQL & " order by adate desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ADATE) Then
            Label11(0).Caption = "Stand: " & rsrs!ADATE
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    For i = 1 To 5
        txtStatus.Text = CInt(txtStatus.Text) + 2
        Label11(i).Caption = sternanzahl(i)
        If CInt(Label11(i).Caption) > 0 Then
            Command5(i + 10).Visible = True
            If Val(Label11(i).Caption) > 10000 Then
                Command5(i + 10).Enabled = False
            Else
                Command5(i + 10).Enabled = True
            End If
        Else
            Command5(i + 10).Visible = False
        End If
    Next i
    
    loeschNEW "KUNDAZEALLE", gdBase
    loeschNEW "STERNTEMP1", gdBase
    loeschNEW "STERNTEMP2", gdBase
    loeschNEW "STERNTEMP3", gdBase
    loeschNEW "STERNTEMP4", gdBase
    loeschNEW "STERNTEMP5", gdBase
    loeschNEW "STERNTEMP6", gdBase
    loeschNEW "STERNTEMP7", gdBase
    loeschNEW "STERNTEMPZ", gdBase
    loeschNEW "STERNTEMPT", gdBase
    
    picprogress.Visible = False
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sternchensuche"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub StammdatenblattKundeDrucken_Bondrucker(cKnr As String)
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    loeschNEW "kuntt", gdBase
    CreateTableT2 "KUNTT", gdBase
    
    sSQL = "Insert into KUNTT select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", FAXNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", MOBILTEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", STADT "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & ", datum1  "
    sSQL = sSQL & ", GESCHLECHT "
    sSQL = sSQL & ", KUNDKART "
    sSQL = sSQL & ", RABATT "
    sSQL = sSQL & ", BONUS "
    sSQL = sSQL & ", KURZTEXT1  "
    sSQL = sSQL & ", KURZTEXT2  "
    sSQL = sSQL & ", NOTIZEN  "
    
    sSQL = sSQL & ", Merkmal "
    sSQL = sSQL & ", Merkmal2 "
    sSQL = sSQL & ", Anrede "
    sSQL = sSQL & ", Kuerzel  "
    sSQL = sSQL & ", Gesperrt  "
    sSQL = sSQL & ", Angelegt  "

    sSQL = sSQL & " from Kunden where kundnr = " & cKnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "update kuntt set osum = '" & Text1(26).Text & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Dim cDatum      As String
    Dim czeit       As String
    ReDim cZeilen(0 To 40) As String
    Dim rsrs        As DAO.Recordset
    Dim iZeilenZahl As Integer
    Dim ctmp        As String
    Dim cZAnfang    As String
    Dim iÜberlänge  As Integer
    
    iZeilenZahl = 0
    cZeilen(iZeilenZahl) = "     Kundenstammdatenblatt      "
    
    iZeilenZahl = iZeilenZahl + 1
    cZeilen(iZeilenZahl) = "--------------------------------"
    
    sSQL = "select * from kuntt "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Kundnr) Then
            ctmp = rsrs!Kundnr
            cZAnfang = "Kundennummer"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!firma) Then
            ctmp = rsrs!firma
            cZAnfang = "Firma"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!anrede) Then
            ctmp = rsrs!anrede
            cZAnfang = "Anrede"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!titel) Then
            ctmp = rsrs!titel
            cZAnfang = "Titel"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Kuerzel) Then
            ctmp = rsrs!Kuerzel
            cZAnfang = "Kürzel"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!vorname) Then
            ctmp = rsrs!vorname
            cZAnfang = "Vorname"
            
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!name) Then
            ctmp = rsrs!name
            cZAnfang = "Name"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Plz) Then
            ctmp = rsrs!Plz
            cZAnfang = "Plz"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!STADT) Then
            ctmp = rsrs!STADT
            cZAnfang = "Ort"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!strasse) Then
            ctmp = rsrs!strasse
            cZAnfang = "Straße"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Email) Then
            ctmp = rsrs!Email
            cZAnfang = "Mail"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Tel) Then
            ctmp = rsrs!Tel
            cZAnfang = "Telefon"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!FAXNR) Then
            ctmp = rsrs!FAXNR
            cZAnfang = "Fax"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Mobiltel) Then
            ctmp = rsrs!Mobiltel
            cZAnfang = "Mobil"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!RABATT) Then
            ctmp = rsrs!RABATT
            cZAnfang = "Rabatt"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!BONUS) Then
            ctmp = rsrs!BONUS
            cZAnfang = "Bonus"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!geschlecht) Then
            ctmp = rsrs!geschlecht
            cZAnfang = "Geschlecht"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!angelegt) Then
            ctmp = rsrs!angelegt
            cZAnfang = "angelegt am"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!Datum1) Then
            ctmp = rsrs!Datum1
            cZAnfang = "Geburtstag"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!KUNDKART) Then
            ctmp = rsrs!KUNDKART
            cZAnfang = "Kundenkarte"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!GESPERRT) Then
            ctmp = rsrs!GESPERRT
            cZAnfang = "gesperrt"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!MERKMAL) Then
            ctmp = rsrs!MERKMAL
            cZAnfang = "Merkmal"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If
        
        If Not IsNull(rsrs!MERKMAL2) Then
            ctmp = rsrs!MERKMAL2
            cZAnfang = "Merkmal"
        
            iZeilenZahl = iZeilenZahl + 1
            If Len(ctmp) + Len(cZAnfang) + 2 > 32 Then
                iÜberlänge = Len(ctmp) + Len(cZAnfang) + 2
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Left(ctmp, 32 - (Len(cZAnfang) + 2))
            Else
                cZeilen(iZeilenZahl) = cZAnfang & ": " & Space(32 - 2 - Len(ctmp) - Len(cZAnfang)) & ctmp
            End If
        End If

        
    End If
    rsrs.Close: Set rsrs = Nothing

    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    iZeilenZahl = iZeilenZahl + 1
    cZeilen(iZeilenZahl) = Space(32)
    
    iZeilenZahl = iZeilenZahl + 1
    cZeilen(iZeilenZahl) = "Datum/Zeit: " & Format(cDatum, "DD.MM.YY") & "    " & czeit

    'Drucke den Beleg

    

    DruckeArbeitszeitBelegWK20d cZeilen(), iZeilenZahl
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StammdatenblattKundeDrucken_Bondrucker"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function fnPruefeDialogEingabenWKL13() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iCount As Integer
    Dim ctmp As String
    Dim cZeichen As String
    Dim cValid As String
    Dim cKundnr As String
    
    For lcount = 5 To 31
        ctmp = Text1(lcount).Text
        ctmp = Trim$(ctmp)
        Select Case lcount
            Case 5, 28, 31, 17
                cValid = "1234567890"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            Case 26, 32
                cValid = "1234567890,"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            Case 22 To 23
                cValid = "1234567890."
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            Case 29 To 30
                cValid = "JN"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            
            Case Is = 7
                cValid = "MWF"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
                
            Case Is = 27
                cValid = "1234567890-,"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL13 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            
        End Select
    Next lcount
    
    If giDlgZustand = giNEU Then
        If gbFilNr Then
            cKundnr = Text1(5).Text
            If Val(cKundnr) <= Val(gcFilNr & "00000") Or Val(cKundnr) > Val(gcFilNr & "99999") Then
                fnPruefeDialogEingabenWKL13 = 999
                Exit Function
            End If
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogEingabenWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Private Sub HoleKundenHistorieWKL13(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim cLBSatz As String
    Dim dUmsatz As Double
    
    dUmsatz = 0
    List6.Clear
    
    List7.Clear
    List7.AddItem "Datum----- Art.Nr Artikelbezeichnung----------------- Menge----- Umsatz----"
    
    loeschNEW "KUHIST", gdBase
    CreateTable "KUHIST", gdBase
    cSQL = "Insert into KUHIST select "
    cSQL = cSQL & "ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", ADATE "
    cSQL = cSQL & ", AZEIT "
    cSQL = cSQL & ", PREIS "
    cSQL = cSQL & ", MENGE "
    cSQL = cSQL & ", FILIALE "
    cSQL = cSQL & ", KUNDNR "
    
    
    cSQL = cSQL & " from KASSJOUR where KUNDNR = " & cKundnr & " order by ADATE desc"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Kuhist inner join  Kunden on KUHIST.kundnr = kunden.kundnr "
    cSQL = cSQL & " set kuhist.name = kunden.name "
    cSQL = cSQL & ",  kuhist.vorname = kunden.vorname "
    cSQL = cSQL & ",  kuhist.titel = kunden.titel "
    cSQL = cSQL & ",  kuhist.firma = kunden.firma "

    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Select * from KASSJOUR where KUNDNR = " & cKundnr & " order by ADATE desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
                ctmp = rsrs!ADATE
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
            cLBSatz = ctmp & " "
            
            If Not IsNull(rsrs!artnr) Then
                ctmp = rsrs!artnr
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(6 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(35 - Len(ctmp))
            cLBSatz = cLBSatz & ctmp & " "
            
            If Not IsNull(rsrs!Menge) Then
                ctmp = rsrs!Menge
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & " "
            
            If Not IsNull(rsrs!Preis) Then
                ctmp = rsrs!Preis
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dUmsatz = dUmsatz + Val(ctmp)
            ctmp = Format$(Val(ctmp), "###,##0.00")
            ctmp = Space$(10 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp & " "
                
            List6.AddItem cLBSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Label5(1).Caption = Format$(dUmsatz, "###,###,##0.00")
    
    Frame11.Visible = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleKundenHistorieWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame3.Enabled = True
    Frame4.Visible = False
    Command5(3).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click()
    On Error GoTo LOKAL_ERROR

    Dim ctmp As String
    Dim lcount As Long
    Dim cZeichen As String
    
   
    ctmp = Trim(Text1(5).Text)
    
    If ctmp = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Text1(5).SetFocus
        Exit Sub
    End If
    gckundnr = ""
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If cZeichen = " " Then
            Exit For
        End If
        If cZeichen >= "0" And cZeichen <= "9" Then
            gckundnr = gckundnr & cZeichen
        End If
    Next lcount
    
    frmWKL94.Show 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim i As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    
    WKL13Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label0(2)
    
    Command10.BackColorFrom = vbWhite
    Command10.BackColorTo = vbWhite
    Command0(0).BackColorFrom = vbWhite
    Command0(0).BackColorTo = vbWhite
    Command0(1).BackColorFrom = vbWhite
    Command0(1).BackColorTo = vbWhite
    
    
    If gbBILDTAST = False Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
    
    If gbEKMAX = True Then
        Label1(8).Caption = "bei Listen-EK: immer größten EK beziehen"
    Else
        Label1(8).Caption = "bei Listen-EK: immer kleinsten EK beziehen"
    End If
            
   
    If gcKundenNr <> "" Then
        Frame0.Visible = False
        Frame2.Visible = False
        If iKasse = 2 Then

        End If
        HoleDatenWKL13 gcKundenNr
        giDlgZustand = giUPD
        
    Else
        If iKasse = 1 Then
            Command1_Click 1
        Else
            Frame0.Visible = True
'            LeereDialogWKL13
        End If
        
        
        
        bSterneDa = sindDieSterneda

    
        If NewTableSuchenDBKombi("STERNEZAHL", gdBase) Then
        
            sSQL = "select top 1 adate from STERNEZAHL "
            sSQL = sSQL & " order by adate desc "
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!ADATE) Then
                    Label11(0).Caption = "Stand: " & rsrs!ADATE
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
            
            For i = 1 To 5
                Label11(i).Caption = sternanzahl(i)
                If Val(Label11(i).Caption) > 0 Then
                    Command5(i + 10).Visible = True
                    If Val(Label11(i).Caption) > 10000 Then
                        Command5(i + 10).Enabled = False
                    Else
                        Command5(i + 10).Enabled = True
                    End If
                End If
            Next i
        Else
            For i = 1 To 5
                Label11(i).Caption = ""
            Next i
            
            Label11(0).Caption = "Drück auf 'Aktualisieren'!"
        End If
        
        
        
        rechneNeuKunden
        
        Label1(22).Caption = ermNeuKunden(1)
        Label1(22).Refresh
        
        Label1(20).Caption = ermNeuKunden(2)
        Label1(20).Refresh
        
        Label1(23).Caption = ermNeuKunden(3)
        Label1(23).Refresh
        
        Label1(21).Caption = ermNeuKunden(4)
        Label1(21).Refresh
        
        fuellecombo
        fülleBLZ Combo2
        
        LeereDialogWKL13
        
    End If
    
    
    
    Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleRab(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    cbox.Clear
    cbox.AddItem "Rabatt wählen"
    cbox.Text = "Rabatt wählen"


    sSQL = "Select distinct(Rabatt) from Kunden order by Rabatt"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!RABATT) Then
                cFeld = rsrs!RABATT
                cbox.AddItem cFeld
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
    Fehler.gsFunktion = "fülleRab"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleBLZ(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    
    cbox.Clear
    cbox.AddItem "BLZ auswählen"
    cbox.Text = "BLZ auswählen"

    If NewTableSuchenDBKombi("KUBLZ", gdBase) = False Then
    
        Screen.MousePointer = 11
        CreateTableT2 "KUBLZ", gdBase
        
        sSQL = "Insert into KUBLZ Select distinct(left(ECIDENT,8)) as KUBLZ,count(Kundnr) as anzahl from Kunden "
        sSQL = sSQL & " where not ECIDENT is null and ECIDENT <> '' "
        sSQL = sSQL & " Group by left(ECIDENT,8)"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update KUBLZ inner join Banken on KUBLZ.KUBLZ = Banken.BLZ "
        sSQL = sSQL & " set KUBLZ.Bankname = Banken.Bankname "
        gdBase.Execute sSQL, dbFailOnError
        
        Screen.MousePointer = 0
        
    End If
        
    sSQL = "Select KUBLZ,Bankname,Anzahl from KUBLZ "
    
    If Label12.Caption = "BLZ - sortiert nach BLZ" Then
        sSQL = sSQL & " Order by KUBLZ "
    ElseIf Label12.Caption = "BLZ - sortiert nach Bankname" Then
        sSQL = sSQL & " Order by Bankname "
    ElseIf Label12.Caption = "BLZ - sortiert nach Kundenanzahl" Then
        sSQL = sSQL & " Order by anzahl desc"
    End If
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!KUBLZ) Then
                cFeld = rsrs!KUBLZ
                cSatz = cFeld & " "
                
            End If
            
            If Not IsNull(rsrs!BankName) Then
                cFeld = rsrs!BankName
                cSatz = cSatz & cFeld & " "
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                cFeld = rsrs!ANZAHL
                cSatz = cSatz & "(" & cFeld & ")"
            End If
            
            
            cbox.AddItem cSatz
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fülleBLZ"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub fuellecombo()
    On Error GoTo LOKAL_ERROR
    
    Combo1.Clear
    Combo1.AddItem "bitte wählen"
    Combo1.AddItem "Listen - EK"
    Combo1.AddItem "Schnitt - EK"
    Combo1.AddItem "Kassen - VK"
    Combo1.AddItem "Listen - VK"
    Combo1.AddItem "spez Kassen - VK"
    Combo1.AddItem "Listen - VK mit Ausn."
    Combo1.AddItem "Netto - KVK"
    Combo1.Text = "bitte wählen"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleCombo"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function sindDieSterneda() As Boolean
On Error GoTo LOKAL_ERROR

Dim i As Integer

sindDieSterneda = True

For i = 0 To 20
    If Modul6.FindFile(App.Path, "\" & i & "s.bmp") = False Then
        sindDieSterneda = False
    End If
Next i

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "sindDieSterneda"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub zeige_Grid()
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
    
    If Not NewTableSuchenDBKombi("TOPK" & srechnertab, gdBase) Then
        MsgBox "Keine Kunden gefunden!", vbInformation, "Winkiss Hinweis:"
        
        Exit Sub
    End If
    
    Set recAnz = gdBase.OpenRecordset("TOPK" & srechnertab)
    
    If recAnz.EOF Then
        
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        Frame1.Visible = False
        Frame0.Visible = True
        MsgBox "Keine Kunden gefunden!", vbInformation, "Winkiss Hinweis:"
        
        Exit Sub
    Else
        recAnz.MoveLast
        If recAnz.RecordCount > 1000 Then
            iRet = MsgBox("Uppss..." & vbCrLf & "Es wurden " & recAnz.RecordCount & " Datensätze gefunden." & vbCrLf & "Wirklich anzeigen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbNo Then
                Exit Sub
            End If
        End If
    End If
    recAnz.Close: Set recAnz = Nothing
    
    
    Screen.MousePointer = 11

    Tabcheck "BEAKU"
    
    FormatGridOverTablay "BEAKU"

    With MSFlexGrid1
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
        Next j
    
        FuellenMSFlex13
        
        ermittlespalten
        
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
        
        FaerbenGrid MSFlexGrid1, CInt(SpaltennummerAWM), CInt(SpaltennummerKUNDNR)
        
        FaerbeKundemitKB MSFlexGrid1, SpaltennummerKUNDNR, SpaltennummerSTERNE

        
        Frame1.Visible = True
        
'        If CInt(gcFilNr) > 0 Then
'            Command2(2).Enabled = False
'        Else
'            Command2(2).Enabled = True
'        End If
        
        Frame0.Visible = False
        .Visible = True
        .Redraw = True
        .Row = 1
        .SetFocus
    
    End With
   
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Grid"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub FaerbeKundemitKB(gridx As MSFlexGrid, spalteKDNR As Byte, spaltesterne As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim i           As Integer
    Dim lfakt       As Long
    Dim sierg       As Single
    Dim sKDNR       As String
    
    
    With gridx
        
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spalteKDNR
            sKDNR = .Text
            
            
            If IsNumeric(sKDNR) Then

                If ermoffenKUBmitKDNR(sKDNR) Then
    
                    For i = 1 To .Cols - 1
                        If i = spaltesterne Then
                        
                        Else
                            .Col = i
                            .CellBackColor = vbGreen
                        End If
                    
                    Next i
                    
                   
                End If
            End If
        Next j
            

    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeKundemitKB"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex13()
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
    Dim corder      As String
    
    If Option2(0).Value Then
        cSQL = cSQL & " order by NAME, VORNAME, STADT "
    ElseIf Option2(1).Value Then
        cSQL = cSQL & " order by DATUM1, NAME, VORNAME, STADT "
    ElseIf Option2(2).Value Then
        cSQL = cSQL & " order by DATUM2, NAME, VORNAME, STADT "
    ElseIf Option2(3).Value Then
        cSQL = cSQL & " order by MERKMAL, NAME, VORNAME, STADT "
    Else
        cSQL = cSQL & " order by KUERZEL "
    End If
    
    Label11(7).Caption = ""

    picprogress.Visible = True
    
    cSQL = "Select * from TOPK" & srechnertab & corder
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        rsrs.MoveLast
        counter = rsrs.RecordCount
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                txtStatus.Text = (lrow * 100) / counter
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        
                        Select Case sSpaltenname(i)
                            Case Is = "noch offen", "Umsatz VJ", "Umsatz LJ", "Bonus"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                                
                            Case Is = "BE"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    If (rsrs(sSpaltenbez(i))) = True Then
                                        sWert = "J"
                                    Else
                                        sWert = "N"
                                    End If
                                Else
                                    sWert = "N"
                                End If
                                .Row = lrow
                                .Text = sWert
                                
                            Case Is = "Stern Bewertung"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                
                                .Row = lrow
                                
                                
                                If bSterneDa Then
                                    Set .CellPicture = LoadPicture(App.Path & "\" & sWert & "s.bmp")
                                    .CellForeColor = vbWhite
                                    .Text = sWert
                                Else
                                    For j = 0 To CInt(sWert)
                                    .Text = .Text & "*"
                                    Next
                                End If
                                
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
        Label11(7).Caption = "Kunden: " & lrow
    End With
    
    picprogress.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex13"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "KUNDNR"
                SpaltennummerKUNDNR = i
            Case Is = "STERNE"
                SpaltennummerSTERNE = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(8).ForeColor = glS1
    Label4(37).ForeColor = glS1
    Label4(38).ForeColor = glS1
    Label4(39).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

Label12.ForeColor = glS1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame5_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 24 Then
    
        Screen.MousePointer = 11
    
        Label1(22).Caption = "...warten"
        Label1(22).Refresh
        
        Label1(20).Caption = "...warten"
        Label1(20).Refresh
        
        Label1(23).Caption = "...warten"
        Label1(23).Refresh
        
        Label1(21).Caption = "...warten"
        Label1(21).Refresh
        
        
        rechneNeuKunden
        
        Label1(22).Caption = ermNeuKunden(1)
        Label1(22).Refresh
        
        Label1(20).Caption = ermNeuKunden(2)
        Label1(20).Refresh
        
        Label1(23).Caption = ermNeuKunden(3)
        Label1(23).Refresh
        
        Label1(21).Caption = ermNeuKunden(4)
        Label1(21).Refresh
        
        Screen.MousePointer = 0
    
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_DblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Index = 5 Then
        Label1(5).Caption = "alle Farben"
        Label1(5).Tag = ""
        Label1(5).BackColor = Label1(4).BackColor
        Label1(5).ForeColor = Label1(4).ForeColor
    ElseIf Index = 8 Then
        If Label1(8).Caption = "bei Listen-EK: immer kleinsten EK beziehen" Then
            Label1(8).Caption = "bei Listen-EK: immer größten EK beziehen"
            
            sSQL = "Update KASSEIN Set EKMAX = True"
            gdBase.Execute sSQL, dbFailOnError
            gbEKMAX = True
        
        Else
            Label1(8).Caption = "bei Listen-EK: immer kleinsten EK beziehen"
            
            sSQL = "Update KASSEIN Set EKMAX = False"
            gdBase.Execute sSQL, dbFailOnError
            gbEKMAX = False
            
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 24 Then
        Label1(24).ForeColor = glLink
    ElseIf Index = 8 Then
        Label1(8).ForeColor = glLink
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label12_Click()
On Error GoTo LOKAL_ERROR

    If Label12.Caption = "BLZ - sortiert nach BLZ" Then
        Label12.Caption = "BLZ - sortiert nach Bankname"
        fülleBLZ Combo2
    ElseIf Label12.Caption = "BLZ - sortiert nach Bankname" Then
        Label12.Caption = "BLZ - sortiert nach Kundenanzahl"
        fülleBLZ Combo2
    ElseIf Label12.Caption = "BLZ - sortiert nach Kundenanzahl" Then
        Label12.Caption = "BLZ - sortiert nach BLZ"
        fülleBLZ Combo2
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label12_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

Label12.ForeColor = glLink

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label12_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label3_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    If Index = 2 Then
        For lcount = 0 To 30
            Text1(lcount).BackColor = vbWhite
        Next lcount
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Change"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim Task$
    Dim sZielpfad       As String
    Dim sKUNDNR         As String

    Select Case Index
        
        Case 37
            If Trim(Text1(5).Text) = "" Then
                MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
                Text1(5).SetFocus
                Exit Sub
            End If
            
            Kunden_Pic_hinzu Trim(Text1(5).Text)
        
        Case 38 'Bild löschen
        
            Del_Kunden_Bild Label5(12).Caption
            Zeige_Kunden_Bilder Trim(Text1(5).Text), 0
            
        Case 39 'Bild aufnehmen
        
            If Trim(Text1(5).Text) = "" Then
                MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
                Text1(5).SetFocus
                Exit Sub
            End If
    
            sKUNDNR = Text1(5).Text
            
            sZielpfad = gcDBPfad
            If Right(sZielpfad, 1) <> "\" Then
                sZielpfad = sZielpfad & "\"
            End If
            sZielpfad = sZielpfad & "PICTURE\KUNDEN"
            
            File2.Path = sZielpfad & "\"
            File2.Pattern = sKUNDNR & "*.jpg"
            File2.Refresh
            
            If File2.ListCount >= 3 Then
                MsgBox "Bitte löschen Sie erst Fotos, bevor Sie weitere hinzufügen.", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
        
            
        
            'Kamera starten
            If gsWebcamPfad <> "" And gsFotoPfad <> "" Then
                Task = Shell(gsWebcamPfad, 1)  'Display öffnen
            End If
            
            'Alles löschen vorher
            Kill gsFotoPfad & "\*.jpg"
            
            'Killer Timer beenden
            Timer2.Enabled = False
            
            'Timer starten
            Timer1.Enabled = True
            Timer1.Interval = 1000
        
    End Select
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Label4_Click"
        Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Label4_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 30
            Command8_Click
        Case 37
            If Trim(Text1(5).Text) = "" Then
                MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
                Text1(5).SetFocus
                Exit Sub
            End If
            
            Kunden_Pic_hinzu Trim(Text1(5).Text)
        
        Case 38 'Bild löschen
        
            Del_Kunden_Bild Label5(12).Caption
            Zeige_Kunden_Bilder Trim(Text1(5).Text), 0
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 37
            Label4(37).ForeColor = glLink
        Case 38
            Label4(38).ForeColor = glLink
        Case 39
            Label4(39).ForeColor = glLink
    End Select
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    glSelect = MSFlexGrid1.Row
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    If KeyCode = vbKeyF3 Then
    
        lrow = MSFlexGrid1.Row
        gckundnr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerKUNDNR)
        If gckundnr <> "" Then
            
            frmWKL74.Show 1
            gckundnr = ""
            Me.Refresh
            Screen.MousePointer = 11

            MSFlexGrid1.TopRow = lrow
            MSFlexGrid1.Col = SpaltennummerKUNDNR
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
            Screen.MousePointer = 0
        End If
    End If
    
    MSFlexGrid1.Redraw = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row > 1 Then
        Command2_Click 0
    Else
        sortierenGrid MSFlexGrid1
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    schreibeKUPKZProt " für Kunde: " & Text1(5).Text & " " & Text1(10).Text & " " & Option1(Index).Caption & " wird gesetzt"
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Picture2_Click()
On Error GoTo LOKAL_ERROR

    If bSterneDa = False Then
        Exit Sub
    End If

    If Picture2.Tag = "" Then
        Picture2.Picture = LoadPicture(App.Path & "\20s.bmp")
        Picture2.Tag = "20"
    ElseIf Picture2.Tag = "20" Then
        Picture2.Picture = LoadPicture(App.Path & "\18s.bmp")
        Picture2.Tag = "18"
    ElseIf Picture2.Tag = "18" Then
        Picture2.Picture = LoadPicture(App.Path & "\15s.bmp")
        Picture2.Tag = "15"
    ElseIf Picture2.Tag = "15" Then
        Picture2.Picture = LoadPicture(App.Path & "\11s.bmp")
        Picture2.Tag = "11"
    ElseIf Picture2.Tag = "11" Then
        Picture2.Picture = LoadPicture(App.Path & "\6s.bmp")
        Picture2.Tag = "6"
    ElseIf Picture2.Tag = "6" Then
        Picture2.Picture = LoadPicture(App.Path & "\19s.bmp")
        Picture2.Tag = "19"
    ElseIf Picture2.Tag = "19" Then
        Picture2.Picture = LoadPicture(App.Path & "\17s.bmp")
        Picture2.Tag = "17"
    ElseIf Picture2.Tag = "17" Then
        Picture2.Picture = LoadPicture(App.Path & "\16s.bmp")
        Picture2.Tag = "16"
    ElseIf Picture2.Tag = "16" Then
        Picture2.Picture = LoadPicture(App.Path & "\14s.bmp")
        Picture2.Tag = "14"
    ElseIf Picture2.Tag = "14" Then
        Picture2.Picture = LoadPicture(App.Path & "\13s.bmp")
        Picture2.Tag = "13"
    ElseIf Picture2.Tag = "13" Then
        Picture2.Picture = LoadPicture(App.Path & "\12s.bmp")
        Picture2.Tag = "12"
    ElseIf Picture2.Tag = "12" Then
        Picture2.Picture = LoadPicture(App.Path & "\10s.bmp")
        Picture2.Tag = "10"
    ElseIf Picture2.Tag = "10" Then
        Picture2.Picture = LoadPicture(App.Path & "\9s.bmp")
        Picture2.Tag = "9"
    ElseIf Picture2.Tag = "9" Then
        Picture2.Picture = LoadPicture(App.Path & "\8s.bmp")
        Picture2.Tag = "8"
    ElseIf Picture2.Tag = "8" Then
        Picture2.Picture = LoadPicture(App.Path & "\7s.bmp")
        Picture2.Tag = "7"
    ElseIf Picture2.Tag = "7" Then
        Picture2.Picture = LoadPicture(App.Path & "\5s.bmp")
        Picture2.Tag = "5"
    ElseIf Picture2.Tag = "5" Then
        Picture2.Picture = LoadPicture(App.Path & "\4s.bmp")
        Picture2.Tag = "4"
    ElseIf Picture2.Tag = "4" Then
        Picture2.Picture = LoadPicture(App.Path & "\3s.bmp")
        Picture2.Tag = "3"
    ElseIf Picture2.Tag = "3" Then
        Picture2.Picture = LoadPicture(App.Path & "\2s.bmp")
        Picture2.Tag = "2"
    ElseIf Picture2.Tag = "2" Then
        Picture2.Picture = LoadPicture(App.Path & "\1s.bmp")
        Picture2.Tag = "1"
    ElseIf Picture2.Tag = "1" Then
        Picture2.Picture = LoadPicture(App.Path & "\0s.bmp")
        Picture2.Tag = "0"
    ElseIf Picture2.Tag = "0" Then
        Picture2.Picture = LoadPicture(App.Path & "\20s.bmp")
        Picture2.Tag = "20"
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Picture5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim iDiv As Integer
    
    Select Case Screen.Height
        Case Is > 15000
            iDiv = 250
        Case Is > 12000
            iDiv = 200
        Case Is > 11000
            iDiv = 150
        Case Is > 10000
            iDiv = 120
        Case Is > 8000
            iDiv = 100
    End Select
    
    Dim ctemp As String
    ctemp = Picture5(Index).Tag
    
    Zeige_weitere_Kunden_Bilder Picture5(Index).Tag, 0, Picture3, Image4, iDiv
    
    Label5(12).Caption = ctemp
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub Picture5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo LOKAL_ERROR
'
'    Dim iDiv As Integer
'
'    Select Case Screen.Height
'        Case Is > 15000
'            iDiv = 250
'        Case Is > 12000
'            iDiv = 200
'        Case Is > 11000
'            iDiv = 150
'        Case Is > 10000
'            iDiv = 120
'        Case Is > 8000
'            iDiv = 100
'    End Select
'
'    Dim ctemp As String
'    ctemp = Picture5(Index).Tag
'
'    Zeige_weitere_Kunden_Bilder Picture5(Index).Tag, 0, Picture3, Image4, iDiv
'    Label5(12).Caption = ctemp
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Picture5_MouseMove"
'    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim lcount As Long
    
    
    
    For lcount = 0 To 35
            Text1(lcount).BackColor = vbWhite
    Next lcount

    If Index >= 5 Then
        Label3(1).Caption = Label4(Index - 5).Caption
        Label3(2).Caption = Trim$(Str$(Index))
    Else
        Label3(1).Caption = Label1(Index).Caption
        Label3(2).Caption = Trim$(Str$(Index))
    End If
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim bSpringen As Boolean
    Dim cValid As String
    
    bSpringen = False
    
    Select Case Index
    
        Case 2
            cZeichen = Chr$(KeyAscii)
            cZeichen = UCase$(cZeichen)
            KeyAscii = Asc(cZeichen)
            
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 21
            cZeichen = Chr$(KeyAscii)
            KeyAscii = Asc(cZeichen)
            
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 6, 29, 30
            cZeichen = Chr$(KeyAscii)
            cZeichen = UCase$(cZeichen)
            KeyAscii = Asc(cZeichen)
        Case 7
        
            cZeichen = Chr$(KeyAscii)
            cZeichen = UCase$(cZeichen)
            KeyAscii = Asc(cZeichen)
            
            cValid = "WMwmfF" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            Else
                If cZeichen = "M" Then
                    cbo1.Text = "Herr"
                ElseIf cZeichen = "W" Then
                    cbo1.Text = "Frau"
                ElseIf cZeichen = "F" Then
                    cbo1.Text = "Familie"
                End If
            End If
        Case 0
            cZeichen = Chr$(KeyAscii)
            cZeichen = UCase$(cZeichen)
            KeyAscii = Asc(cZeichen)
            
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    
        Case 5
            cZeichen = Chr$(KeyAscii)
            KeyAscii = Asc(cZeichen)
   
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    
    
        Case 22, 23
        
            cZeichen = Chr$(KeyAscii)
            KeyAscii = Asc(cZeichen)
            
            cValid = "1234567890." & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 36
        
            cZeichen = Chr$(KeyAscii)
            KeyAscii = Asc(cZeichen)
            
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
    
    'Hier die doppelten Kommas prüfen, nur bei den Preisen
    Select Case Index
        Case 36 'Bonus
            If cZeichen = "," Then
                If InStr(Text1(Index).Text, ",") > 0 Then
                    KeyAscii = 0
                End If
            End If
        Case Else
    End Select
    'Ende*****Hier die doppelten Kommas prüfen, nur bei den Preisen
    
    
    If Len(Text1(Index).Text) = Text1(Index).MaxLength - 1 Then
        If KeyAscii <> 8 And KeyAscii <> 0 And KeyAscii <> 13 And KeyAscii <> 27 Then
            bSpringen = True
        End If
    End If
    
    If bSpringen Then
        If Index = 4 Then
            Command1(0).SetFocus
        ElseIf Index = 19 Then
            Text1(28).SetFocus
        ElseIf Index = 16 Then
            Text1(34).SetFocus
        ElseIf Index = 28 Then
            Text1(20).SetFocus
        ElseIf Index = 27 Then
            Text1(29).SetFocus
        ElseIf Index = 34 Then
            Text1(18).SetFocus
        ElseIf Index = 32 Then
            Command5(0).SetFocus
        Else
            Text1(Index + 1).SetFocus
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Index < 5 Then
        If KeyCode = vbKeyReturn Then
            Command1_Click 0
        End If
    End If
    If KeyCode = vbKeyEscape Then
        Command1_Click 1
    End If
    
    If Index = 22 Or Index = 23 Then
        If KeyCode = 8 Then
        
        Else
            If Len(Text1(Index).Text) = 2 Then
                Text1(Index).Text = Text1(Index).Text & "."
            ElseIf Len(Text1(Index).Text) = 5 Then
                Text1(Index).Text = Text1(Index).Text & "."
            End If
            Text1(Index).SelStart = Len(Text1(Index).Text)
            
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Label3(2).Caption = 999
    Text2.SelStart = Len(Text2.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    cZeichen = Chr$(KeyAscii)
    
    cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
    cValid = cValid & Chr(38) & Chr(45) & Chr(46) & Chr(10) & Chr(13) '& - .
    cValid = cValid & "+äÄÜüÖöß%!?,:"
        
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
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim sZielpfad       As String
    Dim i               As Integer
    Dim sBildname       As String
    
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim lRet            As String
    Dim lfail           As Long
    
    'kopiere Kundenfoto
    Dim sKUNDNR         As String
    sKUNDNR = Text1(5).Text
    
    sPfad = gsFotoPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    sZielpfad = gcDBPfad
    If Right(sZielpfad, 1) <> "\" Then
        sZielpfad = sZielpfad & "\"
    End If
    sZielpfad = sZielpfad & "PICTURE\KUNDEN"
    
    
    File1.Path = gsFotoPfad
    File1.Pattern = "*.jpg"
    File1.Refresh
    
    File2.Path = sZielpfad & "\"
    File2.Pattern = sKUNDNR & "*.jpg"
    File2.Refresh
    
                
    If File1.ListCount > 0 Then
    
        sBildname = File1.list(0)
        
        cQuelle = sPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & sBildname

        cZiel = sZielpfad & "\"
        cZiel = ShortPath(cZiel)
        
        If File2.ListCount > 0 Then
            If FileExists(sZielpfad & "\" & sKUNDNR & ".jpg") = False Then
                cZiel = sZielpfad & "\"
                cZiel = ShortPath(cZiel)
                cZiel = cZiel & sKUNDNR & ".jpg"
            End If
            
            If FileExists(sZielpfad & "\" & sKUNDNR & "_1.jpg") = False Then
                cZiel = sZielpfad & "\"
                cZiel = ShortPath(cZiel)
                cZiel = cZiel & sKUNDNR & "_1.jpg"
            End If
            
            If FileExists(sZielpfad & "\" & sKUNDNR & "_2.jpg") = False Then
                cZiel = sZielpfad & "\"
                cZiel = ShortPath(cZiel)
                cZiel = cZiel & sKUNDNR & "_2.jpg"
            End If
        Else
            cZiel = sZielpfad & "\"
            cZiel = ShortPath(cZiel)
            cZiel = cZiel & sKUNDNR & ".jpg"
        End If

        lRet = CopyFile(cQuelle, cZiel, lfail)
        
        Zeige_Kunden_Bilder Trim(Text1(5).Text), 0
        
        Timer2.Enabled = True
        Timer2.Interval = 100
            
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Timer2_Timer()
On Error GoTo LOKAL_ERROR

'Kamera.exe suchen finden und schließen
                                        
    Dim hwnd&
    Dim Y As String
    Dim Result&
    Dim Title$
    
    Dim bgefunden As Boolean
    bgefunden = False
                                    
    Y = "Logitech" & Chr(174) & " Webcam-Software"
    hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
    Do
        Result = GetWindowTextLength(hwnd) + 1
        Title = Space(Result)
        Result = GetWindowText(hwnd, Title, Result)
        Title = Left$(Title, Len(Title) - 1)

        If InStr(1, Title, Y) Then
            SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
        End If

        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop Until hwnd = 0
    
    Y = "Logitech" & Chr(174) & " Webcam-Software"
    hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
    Do
        Result = GetWindowTextLength(hwnd) + 1
        Title = Space(Result)
        Result = GetWindowText(hwnd, Title, Result)
        Title = Left$(Title, Len(Title) - 1)

        If InStr(1, Title, Y) Then
            bgefunden = True
        End If

        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop Until hwnd = 0
    
    If bgefunden = False Then
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
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

