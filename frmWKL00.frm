VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL00 
   BackColor       =   &H00E0E0E0&
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
   Icon            =   "frmWKL00.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command15 
      BackColor       =   &H80000000&
      Caption         =   "Test"
      Height          =   375
      Left            =   10200
      TabIndex        =   218
      Tag             =   "oday"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   2520
   End
   Begin VB.Frame Frame23 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   174
      ToolTipText     =   "all kundenliste"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   39
         Left            =   120
         TabIndex        =   182
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "nach Zugang"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   38
         Left            =   120
         TabIndex        =   181
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Größenauswertung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   37
         Left            =   120
         TabIndex        =   176
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Produktgruppenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   36
         Left            =   120
         TabIndex        =   183
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   35
         Left            =   120
         TabIndex        =   175
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikelgruppenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   34
         Left            =   120
         TabIndex        =   177
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Linienliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   33
         Left            =   120
         TabIndex        =   178
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Markenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   32
         Left            =   120
         TabIndex        =   179
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Farbenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   31
         Left            =   120
         TabIndex        =   180
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenbindung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   240
      TabIndex        =   171
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame22 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   164
      ToolTipText     =   "Preislagenstatistiken"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   16
         Left            =   120
         TabIndex        =   167
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Preislagenstatistik / Lieferant"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   15
         Left            =   120
         TabIndex        =   169
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   14
         Left            =   120
         TabIndex        =   166
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Preislagenstatistik / AGN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   13
         Left            =   120
         TabIndex        =   165
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "&Preislagenstatistik "
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   12
         Left            =   120
         TabIndex        =   168
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Preislagenstatistik / PGN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   158
      ToolTipText     =   "Detaildaten"
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   11
         Left            =   120
         TabIndex        =   162
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Marken bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikelgruppen bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   160
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Produktgruppen bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   163
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   161
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Produktlinien bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   150
      ToolTipText     =   "all kundenliste"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   40
         Left            =   120
         TabIndex        =   197
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Mailing Feedback"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   30
         Left            =   120
         TabIndex        =   156
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenstückzahlen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   29
         Left            =   120
         TabIndex        =   155
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenumsätze"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   28
         Left            =   120
         TabIndex        =   154
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenerträge"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   27
         Left            =   120
         TabIndex        =   153
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bonusliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   26
         Left            =   120
         TabIndex        =   151
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "allg. Kundenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   25
         Left            =   120
         TabIndex        =   157
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   24
         Left            =   120
         TabIndex        =   152
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Rabattliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame19 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   145
      ToolTipText     =   "Bestellung"
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   23
         Left            =   120
         TabIndex        =   148
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Manuell"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   22
         Left            =   120
         TabIndex        =   149
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   21
         Left            =   120
         TabIndex        =   147
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Berechnung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   137
      ToolTipText     =   "Stammdaten"
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   19
         Left            =   120
         TabIndex        =   195
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Excel Import"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   139
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "KISS Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   140
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "andere Formate"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   141
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   134
      ToolTipText     =   "kissnet"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   20
         Left            =   120
         TabIndex        =   135
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Mail Box"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   19
         Left            =   120
         TabIndex        =   146
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   18
         Left            =   120
         TabIndex        =   136
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Sammelbestellungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   17
         Left            =   120
         TabIndex        =   138
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "globale Artikelsuche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   16
         Left            =   120
         TabIndex        =   142
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Stammdaten aktuell"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   15
         Left            =   120
         TabIndex        =   143
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferanten aktuell"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   14
         Left            =   120
         TabIndex        =   144
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Email schreiben"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.CheckBox ChkLM 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00C0C000&
      Caption         =   "In den lokalen Modus umschalten."
      Height          =   210
      Left            =   4440
      TabIndex        =   132
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2520
   End
   Begin Crystal.CrystalReport CrystalReport3 
      Left            =   120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   450
      WindowHeight    =   680
      WindowTitle     =   "Winkiss Fehlermeldung"
      WindowBorderStyle=   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowZoomCtl=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   127
      ToolTipText     =   "Zentrale"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   11
         Left            =   120
         TabIndex        =   129
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kassendateien ausgeben"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   133
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   131
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kassendateien einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   123
      ToolTipText     =   "Lieferanten"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   126
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Rechnungsübersicht"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   124
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferanten bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   125
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferantenzusammenstellung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   128
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   115
      ToolTipText     =   "Datenbank"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   13
         Left            =   120
         TabIndex        =   121
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenbankbefehl"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   12
         Left            =   120
         TabIndex        =   120
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenbank bereinigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   119
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenbank Monitor"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   118
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Fremddaten importieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   117
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenbank optimieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   122
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   116
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Pfad zur Datenbank"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   107
      ToolTipText     =   "Artikel"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   11
         Left            =   120
         TabIndex        =   196
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Pennerartikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   113
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikel löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   112
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikel retournieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   111
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Terminpreise"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   114
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   109
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bestandskorrektur"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikel bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command14 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   110
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kalkulation der Preise"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
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
      Height          =   270
      Left            =   120
      TabIndex        =   106
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.FileListBox File3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Pattern         =   "*.DBF"
      TabIndex        =   105
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   104
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   69
      ToolTipText     =   "Artikellisten"
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   77
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Diverse Listen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   76
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Mindestbestand ermitteln"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   72
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "unterschrittene Mindestmenge"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Gutscheinverwaltung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "reduzierte VK-Preise"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   73
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "negativer Bestand"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "nach Artikelgruppen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "nach Lieferanten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   21
      ToolTipText     =   "Listen"
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Verkaufslisten..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenliste..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bestand nach Verkauf"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Warenzugang / Einkauf"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bestandsliste / Inventur"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Favoritenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferantenliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikelliste..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   84
      ToolTipText     =   "Kreditverkäufe"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   87
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   86
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Sammelrechnung..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kreditverwaltung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Etiketten"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   18
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Spezialetiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   17
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Plakate"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Etiketten aus Lieferschein"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Etiketten selbst wählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Etiketten drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   615
         Index           =   20
         Left            =   120
         TabIndex        =   213
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Rabatt - Aufkleber"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   51
      ToolTipText     =   "Protokolle"
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   194
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "GDPdU/DATEV"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   192
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "allg. Kassenvorgänge"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   191
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kassenprotokolle"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   190
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikelliste aus MDE/Scanner"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   189
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bestandsveränderungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   193
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   188
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Arbeitszeitauswertung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   187
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Rabattverkäufe"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   186
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Verkaufsprotokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   93
      ToolTipText     =   "Statistiken"
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   101
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Preislagenstatistiken..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   100
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Zeitenstatistik"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Sortimentsanalyse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   102
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   98
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Geschäftsanalyse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   97
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Umsatzstatistik"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   96
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenanalyse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   95
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferantenstatistik"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command13 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bedienerstatistik"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Kasse"
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   10
         Left            =   120
         TabIndex        =   38
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundenbestellung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Zusammenfassung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kassenbuch"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kassieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bargeld zählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Protokolle..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kreditverkäufe..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   4455
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
         Caption         =   "Monatsabschluß"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   4455
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
         Caption         =   "Bedienernamen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Tagesbericht"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   88
      ToolTipText     =   "Wareneingang"
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "aus Bestellung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   92
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   90
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "aus Filialumverteilung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command12 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "aus Einzellieferung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Etiketten..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lieferanten..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Wareneingang..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bestellvorschläge..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kundendaten bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Detaildaten bearbeiten..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Stammdaten aktualisieren..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikel..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Pattern         =   "MASTER!.*"
      TabIndex        =   58
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   52
      ToolTipText     =   "Termine"
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command8 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Reparaturverwaltung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Notizen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Termin-Kalender"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Vorgaben"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   59
      ToolTipText     =   "Einstellungen"
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   67
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "MWSt-Sätze"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   66
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bonus auf Bon"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Bedienerverwaltung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   68
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   65
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Warengruppen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Kartenleser konfigurieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Texte Kassenbon"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Drucker definieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Unternehmens-Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   40
      ToolTipText     =   "Service"
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenaustausch..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   6000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "zurück zum Hauptmenü"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   47
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "DTA-Ausgabe"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Programmeinstellungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   3120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Zugriffs-Rechte"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "DsFinVK Expo."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
            Italic          =   -1  'True
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
         Caption         =   "KISSNET..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Datenbank..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Einstellungen..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrinterStartPage=   1
      PrinterCollation=   1
      WindowState     =   1
      PrintFileUseRptNumberFmt=   -1  'True
      PrintFileUseRptDateFmt=   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
   End
   Begin VB.PictureBox picprogress 
      Height          =   300
      Left            =   4680
      ScaleHeight     =   240
      ScaleWidth      =   2595
      TabIndex        =   170
      Top             =   7680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   7
      Left            =   10080
      TabIndex        =   203
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Abmelden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   5
      Left            =   10080
      TabIndex        =   204
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Ende"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   4
      Left            =   8400
      TabIndex        =   205
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Service"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   8
      Left            =   6840
      TabIndex        =   206
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Termine"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   3
      Left            =   5280
      TabIndex        =   207
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   2
      Left            =   3600
      TabIndex        =   208
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Statistiken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   1
      Left            =   1920
      TabIndex        =   209
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Kasse"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   210
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Stammdaten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   211
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Anmelden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   9
      Left            =   1920
      TabIndex        =   212
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      Caption         =   ""
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label lbl_TSE 
      BackColor       =   &H8000000C&
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
      Left            =   240
      TabIndex        =   217
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "PC-V"
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
      Index           =   10
      Left            =   11280
      MouseIcon       =   "frmWKL00.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   216
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "F 0"
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
      Index           =   13
      Left            =   9360
      TabIndex        =   215
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "TSE"
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
      Index           =   6
      Left            =   10680
      MouseIcon       =   "frmWKL00.frx":074C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   214
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hotline: 0511 / 9 559 10"
      Height          =   255
      Index           =   12
      Left            =   8280
      MouseIcon       =   "frmWKL00.frx":0A56
      TabIndex        =   202
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   0
      Left            =   2280
      TabIndex        =   201
      Top             =   1440
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Karten - Terminal"
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
      Index           =   9
      Left            =   1440
      MouseIcon       =   "frmWKL00.frx":0D60
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   200
      Top             =   7900
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonusbriefe erstellen"
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
      Index           =   8
      Left            =   120
      MouseIcon       =   "frmWKL00.frx":106A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   199
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Grundkurs"
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
      Index           =   5
      Left            =   120
      MouseIcon       =   "frmWKL00.frx":1374
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   198
      Top             =   7900
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "2021  TSE integriert"
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
      Index           =   4
      Left            =   3000
      MouseIcon       =   "frmWKL00.frx":167E
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   184
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   8040
      Width           =   6375
   End
   Begin VB.Label lbl6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   28
      Left            =   5520
      TabIndex        =   173
      Top             =   3600
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lbl6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   53
      Left            =   5520
      TabIndex        =   172
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   2520
      TabIndex        =   130
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3840
      MouseIcon       =   "frmWKL00.frx":1988
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   103
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   83
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "von"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   82
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   81
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stand: 14.03.2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   80
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   -120
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   12015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "nicht registrierte Programmversion !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   50
      Top             =   6840
      Width           =   11655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Anwender nicht aktiv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6000
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "www.kisslive.de"
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
      Index           =   7
      Left            =   9600
      MouseIcon       =   "frmWKL00.frx":1C92
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   185
      Top             =   7995
      Width           =   1935
   End
End
Attribute VB_Name = "frmWKL00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim giSprache           As Byte
Dim bStammdaten         As Boolean
Dim bKasse              As Boolean
Dim bStatistiken        As Boolean
Dim bTermine            As Boolean
Dim bListen             As Boolean
Dim bService            As Boolean
Dim bNeuheit            As Boolean
Dim byteZGNr            As Byte

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long
Private Sub CheckProgrammVersion()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr         As Integer
    Dim lRet            As Long
    Dim ctemp           As String
    Dim ctmp            As String
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim lfail           As Long
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim cPfad           As String
    Dim cPfad1          As String
    Dim lVersold        As Long
    Dim k               As Integer
    Dim lDBVersionapp   As Integer
    
    gbDBMod = False
    bNeuheit = False
    k = 0
    
    cPfad1 = gcDBPfad               'Datenbankpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    cPfad = gcPfad                  'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    '//Anwendung
    'aktuelle Winkiss Version
    If Not tableSuchenDBKombi("WKEINSTE", 2) Then
        CreateWKEINSTE
        WKVersion = 300
    Else
        sSQL = "Select PVERSION from WKEINSTE"
        Set rsrs = gdApp.OpenRecordset(sSQL)
        
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!PVersion) Then
                WKVersion = rsrs!PVersion
            End If
        Else
            WKVersion = 300
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    gbSchreibRechnerProto = False

    If WKVersion < 601 Then AppCheckV601
    If WKVersion < 604 Then AppCheckV604
    If WKVersion < 627 Then AppCheckV627
    If WKVersion < 636 Then AppCheckV636
    If WKVersion < 640 Then appcheckV640
    If WKVersion < 668 Then AppCheckV668
    If WKVersion < 673 Then appcheckV673
    If WKVersion < 674 Then AppCheckV674
    If WKVersion < 706 Then AppCheckV706
    If WKVersion < 715 Then AppCheckV715
    If WKVersion < 716 Then AppCheckV716
    If WKVersion < 722 Then AppCheckV722
    If WKVersion < 724 Then AppCheckV724
    If WKVersion < 738 Then AppCheckV738
    If WKVersion < 772 Then AppCheckV772
    If WKVersion < 778 Then AppCheckV778
    If WKVersion < 780 Then AppCheckV780
    If WKVersion < 781 Then AppCheckV781
    If WKVersion < 783 Then AppCheckV783
    If WKVersion < 786 Then AppCheckV786
    If WKVersion < 813 Then AppCheckV813
    If WKVersion < 833 Then AppCheckV833
    If WKVersion < 846 Then AppCheckV846
    If WKVersion < 849 Then AppCheckV849
    If WKVersion < 853 Then AppCheckV853
    If WKVersion < 863 Then AppCheckV863
    If WKVersion < 865 Then AppCheckV865
    If WKVersion < 866 Then AppCheckV866
    If WKVersion < 869 Then AppCheckV869
    If WKVersion < 870 Then AppCheckV870
    If WKVersion < 872 Then AppCheckV872
    If WKVersion < 886 Then AppCheckV886
    If WKVersion < 888 Then AppCheckV888
    If WKVersion < 889 Then AppCheckV889
    If WKVersion < 890 Then AppCheckV890
    If WKVersion < 893 Then AppCheckV893
    If WKVersion < 895 Then AppCheckV895
    If WKVersion < 896 Then AppCheckV896
    If WKVersion < 898 Then AppCheckV898
    If WKVersion < 913 Then AppCheckV913
    If WKVersion < 929 Then AppCheckV929
    If WKVersion < 944 Then AppCheckV944
    If WKVersion < 946 Then AppCheckV946
    If WKVersion < 961 Then AppCheckV961
    If WKVersion < 968 Then AppCheckV968
    If WKVersion < 978 Then AppCheckV978
    If WKVersion < 1011 Then AppCheckV1011
    If WKVersion < 1015 Then AppCheckV1015
    If WKVersion < 1041 Then AppCheckV1041
    If WKVersion < 1060 Then AppCheckV1060
    If WKVersion < 1094 Then AppCheckV1094
    If WKVersion < 1106 Then AppCheckV1106
    If WKVersion < 1107 Then AppCheckV1107
    If WKVersion < 1121 Then AppCheckV1121
    If WKVersion < 1156 Then AppCheckV1156
    If WKVersion < 1157 Then AppCheckV1157
    If WKVersion < 1163 Then AppCheckV1163
    If WKVersion < 1234 Then AppCheckV1234
    If WKVersion < 1302 Then AppCheckV1302
    If WKVersion < 1310 Then AppCheckV1310
    If WKVersion < 1325 Then AppCheckV1325
    If WKVersion < 1339 Then AppCheckV1339
    If WKVersion < 1376 Then AppCheckV1376
    If WKVersion < 1451 Then AppCheckV1451
    If WKVersion < 1499 Then AppCheckV1499
    If WKVersion < 1509 Then AppCheckV1509
    If WKVersion < 1534 Then AppCheckV1534
    If WKVersion < 1571 Then AppCheckV1571
    If WKVersion < 1596 Then AppCheckV1596
    If WKVersion < 1609 Then AppCheckV1609
    If WKVersion < 1616 Then AppCheckV1616
    If WKVersion < 1632 Then AppCheckV1632
    If WKVersion < 1671 Then AppCheckV1671
    If WKVersion < 1695 Then AppCheckV1695
    If WKVersion < 1696 Then AppCheckV1696
    If WKVersion < 1703 Then AppCheckV1703
    If WKVersion < 1709 Then AppCheckV1709
    If WKVersion < 1734 Then AppCheckV1734
    If WKVersion < 1736 Then AppCheckV1736
    If WKVersion < 1750 Then AppCheckV1750
    If WKVersion < 1751 Then AppCheckV1751
    If WKVersion < 1752 Then AppCheckV1752
    If WKVersion < 1753 Then AppCheckV1753
    If WKVersion < 1770 Then AppCheckV1770
    If WKVersion < 1777 Then AppCheckV1777
    If WKVersion < 1808 Then AppCheckV1808
    If WKVersion < 1829 Then AppCheckV1829
    If WKVersion < 1839 Then AppCheckV1839
    If WKVersion < 1859 Then AppCheckV1859
    If WKVersion < 1860 Then AppCheckV1860
    If WKVersion < 1864 Then AppCheckV1864
    If WKVersion < 1876 Then AppCheckV1876
    If WKVersion < 1939 Then AppCheckV1939
    If WKVersion < 1951 Then AppCheckV1951
    If WKVersion < 1954 Then AppCheckV1954
    If WKVersion < 1955 Then AppCheckV1955
    If WKVersion < 1963 Then AppCheckV1963
    If WKVersion < 1965 Then AppCheckV1965
    If WKVersion < 2005 Then AppCheckV2005
    If WKVersion < 2013 Then AppCheckV2013
    If WKVersion < 2071 Then AppCheckV2071
    If WKVersion < 2074 Then AppCheckV2074
    If WKVersion < 2095 Then AppCheckV2095
    If WKVersion < 2096 Then AppCheckV2096
    If WKVersion < 2103 Then AppCheckV2103
    If WKVersion < 2137 Then AppCheckV2137
    If WKVersion < 2162 Then AppCheckV2162
    If WKVersion < 2254 Then AppCheckV2254
    If WKVersion < 2270 Then AppCheckV2270
    If WKVersion < 2290 Then AppCheckV2290
    If WKVersion < 2352 Then AppCheckV2352
    If WKVersion < 2408 Then AppCheckV2408
    If WKVersion < 2439 Then AppCheckV2439
    If WKVersion < 2453 Then AppCheckV2453
    If WKVersion < 2490 Then AppCheckV2490
    If WKVersion < 2518 Then AppCheckV2518
    If WKVersion < 2537 Then AppCheckV2537
    If WKVersion < 2578 Then AppCheckV2578
    If WKVersion < 2589 Then AppCheckV2589
    If WKVersion < 2651 Then AppCheckV2651
    If WKVersion < 2683 Then AppCheckV2683
    If WKVersion < 2684 Then AppCheckV2684
    If WKVersion < 2687 Then AppCheckV2687
    If WKVersion < 2751 Then AppCheckV2751
    If WKVersion < 2761 Then AppCheckV2761
    If WKVersion < 2781 Then AppCheckV2781
    If WKVersion < 2818 Then AppCheckV2818
    If WKVersion < 2825 Then AppCheckV2825
    If WKVersion < 2856 Then AppCheckV2856
    If WKVersion < 2912 Then AppCheckV2912
    If WKVersion < 2925 Then AppCheckV2925
    If WKVersion < 2940 Then AppCheckV2940
    If WKVersion < 2946 Then AppCheckV2946
    If WKVersion < 2955 Then AppCheckV2955
    If WKVersion < 2958 Then AppCheckV2958
    If WKVersion < 2965 Then AppCheckV2965
    If WKVersion < 2966 Then AppCheckV2966
    
    
    If WKVersion < glpVers Then
        'weil ein Programmupdate durchgelaufen ist muss
        'nach dem Lesen der Programmeinstellungen ein neues PCINFO-Protokoll
        'geschrieben werden
        gbSchreibRechnerProto = True
        
        WKVersion = glpVers
    Else
        gbSchreibRechnerProto = False
    End If
    
    sSQL = "Select DBVERSION from WKEINSTE"
    Set rsrs = gdApp.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!DBVersion) Then
            lDBVersionapp = rsrs!DBVersion
        Else
            lDBVersionapp = 0
        End If
    Else
        lDBVersionapp = 0
    End If
    
    rsrs.Close: Set rsrs = Nothing

    '****************************************End Access
    
    '//Datenbank
    If Not NewTableSuchenDBKombi("DBEINSTE", gdBase) Then
        
        giUmleitgrund = 6 'Datenbank abgerissen

        gcUmleittxt = "Wichtige Datenbankinformationen sind verloren." & vbCrLf
        gcUmleittxt = gcUmleittxt & "Drücken Sie 'Weiter', um diese eventuell wieder herzustellen!" & vbCrLf
        
        frmWKL60.Show 1
    Else
        checktab "DBEINSTE", gdBase
        
    End If
    
    sSQL = "Select DBVERSION from DBEINSTE"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!DBVersion) Then
            lDBVersion = rsrs!DBVersion
        Else
            If lDBVersionapp = 0 Then
                lDBVersion = 355
            Else
                lDBVersion = lDBVersionapp
            End If
        End If
    Else
        If lDBVersionapp = 0 Then
            lDBVersion = 355
        Else
            lDBVersion = lDBVersionapp
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    lVersold = lDBVersion
   
    Label1(1).Caption = "Programm " & Trim$(Str$(WKVersion)) & " zu Datenbank " & Trim$(Str$(lDBVersion))
    Label1(1).Refresh
    
    If WKVersion > lDBVersion Then
'nochmal:
'        If BistDualleineinderDatenbank Then
        
            If SmallUpdate("602", lDBVersion) Then DbCheckV602
            
            If SmallUpdate("611", lDBVersion) Then lDBVersion = 611
            
            If BigUpdate("612", lDBVersion) Then DbCheckV612: Kill gcDBPfad & "\upd612.CFG"
            If BigUpdate("629", lDBVersion) Then DbCheckV629: Kill gcDBPfad & "\upd629.CFG"
            
            If BigUpdate("637", lDBVersion) Then DbCheckV637: Kill gcDBPfad & "\upd637.CFG"
            If BigUpdate("640", lDBVersion) Then DbCheckV640: Kill gcDBPfad & "\upd640.CFG"

            If SmallUpdate("641", lDBVersion) Then DbCheckV641
            If BigUpdate("642", lDBVersion) Then DbCheckV642: Kill gcDBPfad & "\upd642.CFG"
            If SmallUpdate("646", lDBVersion) Then DbCheckV646
            
            If SmallUpdate("654", lDBVersion) Then DbCheckV654
            If SmallUpdate("661", lDBVersion) Then DbCheckV661
            If SmallUpdate("666", lDBVersion) Then DbCheckV666
            
            If SmallUpdate("669", lDBVersion) Then DbCheckV669
            If BigUpdate("670", lDBVersion) Then DbCheckV670: Kill gcDBPfad & "\upd670.CFG"
            If BigUpdate("671", lDBVersion) Then DbCheckV671: Kill gcDBPfad & "\upd671.CFG"
            If BigUpdate("673", lDBVersion) Then DbCheckV673: Kill gcDBPfad & "\upd673.CFG"
            If BigUpdate("701", lDBVersion) Then DbCheckV701: Kill gcDBPfad & "\upd701.CFG"
            If SmallUpdate("702", lDBVersion) Then DbCheckV702
            If SmallUpdate("703", lDBVersion) Then DbCheckV703
            If SmallUpdate("712", lDBVersion) Then DbCheckV712
            If SmallUpdate("721", lDBVersion) Then DbCheckV721
            If SmallUpdate("725", lDBVersion) Then DbCheckV725
            If SmallUpdate("728", lDBVersion) Then DbCheckV728
            If BigUpdate("739", lDBVersion) Then DbCheckV739: Kill gcDBPfad & "\upd739.CFG"
            If SmallUpdate("761", lDBVersion) Then DbCheckV761
            If SmallUpdate("762", lDBVersion) Then DbCheckV762
            If SmallUpdate("764", lDBVersion) Then DbCheckV764
            If SmallUpdate("770", lDBVersion) Then DbCheckV770
            
            If SmallUpdate("792", lDBVersion) Then DbCheckV792
            If SmallUpdate("830", lDBVersion) Then DbCheckV830
            If SmallUpdate("839", lDBVersion) Then DbCheckV839
            If BigUpdate("848", lDBVersion) Then DbCheckV848: Kill gcDBPfad & "\upd848.CFG"
            If SmallUpdate("867", lDBVersion) Then DbCheckV867
            If BigUpdate("876", lDBVersion) Then DbCheckV876: Kill gcDBPfad & "\upd876.CFG"
            If BigUpdate("877", lDBVersion) Then DbCheckV877: Kill gcDBPfad & "\upd877.CFG"
            If BigUpdate("878", lDBVersion) Then DbCheckV878: Kill gcDBPfad & "\upd878.CFG"
            
            If SmallUpdate("881", lDBVersion) Then DbCheckV881
            If SmallUpdate("882", lDBVersion) Then DbCheckV882
            If SmallUpdate("887", lDBVersion) Then DbCheckV887
            If SmallUpdate("899", lDBVersion) Then DbCheckV899
            If SmallUpdate("900", lDBVersion) Then DbCheckV900
            If SmallUpdate("901", lDBVersion) Then DbCheckV901
            If SmallUpdate("910", lDBVersion) Then DbCheckV910
            If SmallUpdate("915", lDBVersion) Then DbCheckV915
            If SmallUpdate("919", lDBVersion) Then DbCheckV919
            If SmallUpdate("930", lDBVersion) Then DbCheckV930
            If SmallUpdate("931", lDBVersion) Then DbCheckV931
            If SmallUpdate("934", lDBVersion) Then DbCheckV934
            If SmallUpdate("937", lDBVersion) Then DbCheckV937
            If SmallUpdate("939", lDBVersion) Then DbCheckV939
            If SmallUpdate("944", lDBVersion) Then DbCheckV944
            If SmallUpdate("946", lDBVersion) Then DbCheckV946
            If BigUpdate("950", lDBVersion) Then DbCheckV950: Kill gcDBPfad & "\upd950.CFG"
            If BigUpdate("953", lDBVersion) Then DbCheckV953: Kill gcDBPfad & "\upd953.CFG"
            If SmallUpdate("957", lDBVersion) Then DbCheckV957
            If SmallUpdate("959", lDBVersion) Then DbCheckV959
            If SmallUpdate("970", lDBVersion) Then DbCheckV970
            If SmallUpdate("975", lDBVersion) Then DbCheckV975
            If SmallUpdate("977", lDBVersion) Then DbCheckV977
            If SmallUpdate("979", lDBVersion) Then DbCheckV979
            If SmallUpdate("983", lDBVersion) Then DbCheckV983
            If SmallUpdate("984", lDBVersion) Then DbCheckV984
            If SmallUpdate("989", lDBVersion) Then DbCheckV989
            If SmallUpdate("1006", lDBVersion) Then DbCheckV1006
            If SmallUpdate("1008", lDBVersion) Then DbCheckV1008
            If SmallUpdate("1021", lDBVersion) Then DbCheckV1021
            If SmallUpdate("1029", lDBVersion) Then DbCheckV1029
            If SmallUpdate("1030", lDBVersion) Then DbCheckV1030
            If SmallUpdate("1046", lDBVersion) Then DbCheckV1046
            If SmallUpdate("1052", lDBVersion) Then DbCheckV1052
            If SmallUpdate("1053", lDBVersion) Then DbCheckV1053
            If SmallUpdate("1057", lDBVersion) Then DbCheckV1057
            If SmallUpdate("1061", lDBVersion) Then DbCheckV1061
            If SmallUpdate("1066", lDBVersion) Then DbCheckV1066
            If SmallUpdate("1084", lDBVersion) Then DbCheckV1084
            If SmallUpdate("1086", lDBVersion) Then DbCheckV1086
            If SmallUpdate("1090", lDBVersion) Then DbCheckV1090
            If SmallUpdate("1091", lDBVersion) Then DbCheckV1091
            If SmallUpdate("1103", lDBVersion) Then DbCheckV1103
            If SmallUpdate("1104", lDBVersion) Then DbCheckV1104
            If SmallUpdate("1105", lDBVersion) Then DbCheckV1105
            If SmallUpdate("1113", lDBVersion) Then DbCheckV1113
            If SmallUpdate("1118", lDBVersion) Then DbCheckV1118
            If SmallUpdate("1119", lDBVersion) Then DbCheckV1119
            If SmallUpdate("1133", lDBVersion) Then DbCheckV1133
            If SmallUpdate("1142", lDBVersion) Then DbCheckV1142
            If SmallUpdate("1144", lDBVersion) Then DbCheckV1144
            If SmallUpdate("1180", lDBVersion) Then DbCheckV1180
            If SmallUpdate("1199", lDBVersion) Then DbCheckV1199
            If SmallUpdate("1200", lDBVersion) Then DbCheckV1200
            If SmallUpdate("1201", lDBVersion) Then DbCheckV1201
            If SmallUpdate("1204", lDBVersion) Then DbCheckV1204
            If SmallUpdate("1205", lDBVersion) Then DbCheckV1205
            If SmallUpdate("1207", lDBVersion) Then DbCheckV1207
            If SmallUpdate("1220", lDBVersion) Then DbCheckV1220
            If SmallUpdate("1221", lDBVersion) Then DbCheckV1221
            If SmallUpdate("1223", lDBVersion) Then DbCheckV1223
            If SmallUpdate("1230", lDBVersion) Then DbCheckV1230
            If SmallUpdate("1239", lDBVersion) Then DbCheckV1239
            If SmallUpdate("1241", lDBVersion) Then DbCheckV1241
            If SmallUpdate("1243", lDBVersion) Then DbCheckV1243
            If SmallUpdate("1260", lDBVersion) Then DbCheckV1260
            If SmallUpdate("1277", lDBVersion) Then DbCheckV1277
            If SmallUpdate("1285", lDBVersion) Then DbCheckV1285
            If SmallUpdate("1290", lDBVersion) Then DbCheckV1290
            If SmallUpdate("1292", lDBVersion) Then DbCheckV1292
            If SmallUpdate("1297", lDBVersion) Then DbCheckV1297
            If SmallUpdate("1315", lDBVersion) Then DbCheckV1315
            If SmallUpdate("1319", lDBVersion) Then DbCheckV1319
            If SmallUpdate("1334", lDBVersion) Then DbCheckV1334
            If SmallUpdate("1336", lDBVersion) Then DbCheckV1336
            If SmallUpdate("1337", lDBVersion) Then DbCheckV1337
            If SmallUpdate("1338", lDBVersion) Then DbCheckV1338
            If SmallUpdate("1350", lDBVersion) Then DbCheckV1350
            If SmallUpdate("1357", lDBVersion) Then DbCheckV1357
            If SmallUpdate("1367", lDBVersion) Then DbCheckV1367
            If SmallUpdate("1377", lDBVersion) Then DbCheckV1377
            If SmallUpdate("1381", lDBVersion) Then DbCheckV1381
            If SmallUpdate("1393", lDBVersion) Then DbCheckV1393
            If SmallUpdate("1415", lDBVersion) Then DbCheckV1415
            If SmallUpdate("1420", lDBVersion) Then DbCheckV1420
            If SmallUpdate("1434", lDBVersion) Then DbCheckV1434
            If SmallUpdate("1447", lDBVersion) Then DbCheckV1447
            If SmallUpdate("1455", lDBVersion) Then DbCheckV1455
            If SmallUpdate("1456", lDBVersion) Then DbCheckV1456
            If SmallUpdate("1457", lDBVersion) Then DbCheckV1457
            If SmallUpdate("1459", lDBVersion) Then DbCheckV1459
            If SmallUpdate("1463", lDBVersion) Then DbCheckV1463
            If SmallUpdate("1464", lDBVersion) Then DbCheckV1464
            If SmallUpdate("1467", lDBVersion) Then DbCheckV1467
            If SmallUpdate("1472", lDBVersion) Then DbCheckV1472
            If SmallUpdate("1476", lDBVersion) Then DbCheckV1476
            If SmallUpdate("1480", lDBVersion) Then DbCheckV1480
            If SmallUpdate("1496", lDBVersion) Then DbCheckV1496
            If SmallUpdate("1505", lDBVersion) Then DbCheckV1505
            If SmallUpdate("1514", lDBVersion) Then DbCheckV1514
            If SmallUpdate("1521", lDBVersion) Then DbCheckV1521
            If SmallUpdate("1537", lDBVersion) Then DbCheckV1537
            If SmallUpdate("1550", lDBVersion) Then DbCheckV1550
            If SmallUpdate("1559", lDBVersion) Then DbCheckV1559
            If SmallUpdate("1574", lDBVersion) Then DbCheckV1574
            If SmallUpdate("1584", lDBVersion) Then DbCheckV1584
            If SmallUpdate("1588", lDBVersion) Then DbCheckV1588
            If SmallUpdate("1590", lDBVersion) Then DbCheckV1590
            If SmallUpdate("1601", lDBVersion) Then DbCheckV1601
            If SmallUpdate("1635", lDBVersion) Then DbCheckV1635
            If SmallUpdate("1636", lDBVersion) Then DbCheckV1636
            If SmallUpdate("1642", lDBVersion) Then DbCheckV1642
            If SmallUpdate("1643", lDBVersion) Then DbCheckV1643
            If SmallUpdate("1652", lDBVersion) Then DbCheckV1652
            If SmallUpdate("1656", lDBVersion) Then DbCheckV1656
            If SmallUpdate("1659", lDBVersion) Then DbCheckV1659
            If SmallUpdate("1670", lDBVersion) Then DbCheckV1670
            If SmallUpdate("1697", lDBVersion) Then DbCheckV1697
            If SmallUpdate("1701", lDBVersion) Then DbCheckV1701
            If SmallUpdate("1702", lDBVersion) Then DbCheckV1702
            If SmallUpdate("1705", lDBVersion) Then DbCheckV1705
            If SmallUpdate("1718", lDBVersion) Then DbCheckV1718
            If SmallUpdate("1728", lDBVersion) Then DbCheckV1728
            If SmallUpdate("1745", lDBVersion) Then DbCheckV1745
            If SmallUpdate("1759", lDBVersion) Then DbCheckV1759
            If SmallUpdate("1781", lDBVersion) Then DbCheckV1781
            If SmallUpdate("1788", lDBVersion) Then DbCheckV1788
            If SmallUpdate("1803", lDBVersion) Then DbCheckV1803
            If SmallUpdate("1804", lDBVersion) Then DbCheckV1804
            If SmallUpdate("1813", lDBVersion) Then DbCheckV1813
            If SmallUpdate("1814", lDBVersion) Then DbCheckV1814
            If SmallUpdate("1822", lDBVersion) Then DbCheckV1822
            If SmallUpdate("1832", lDBVersion) Then DbCheckV1832
            If SmallUpdate("1833", lDBVersion) Then DbCheckV1833
            If SmallUpdate("1837", lDBVersion) Then DbCheckV1837
            If SmallUpdate("1855", lDBVersion) Then DbCheckV1855
            If SmallUpdate("1858", lDBVersion) Then DbCheckV1858
            If SmallUpdate("1860", lDBVersion) Then DbCheckV1860
            If SmallUpdate("1868", lDBVersion) Then DbCheckV1868
            If SmallUpdate("1869", lDBVersion) Then DbCheckV1869
            If SmallUpdate("1889", lDBVersion) Then DbCheckV1889
            If SmallUpdate("1909", lDBVersion) Then DbCheckV1909
            If SmallUpdate("1921", lDBVersion) Then DbCheckV1921
            If SmallUpdate("1923", lDBVersion) Then DbCheckV1923
            If SmallUpdate("1934", lDBVersion) Then DbCheckV1934
            If SmallUpdate("1938", lDBVersion) Then DbCheckV1938
            If SmallUpdate("1940", lDBVersion) Then DbCheckV1940
            If SmallUpdate("1944", lDBVersion) Then DbCheckV1944
            If SmallUpdate("1949", lDBVersion) Then DbCheckV1949
            If SmallUpdate("1952", lDBVersion) Then DbCheckV1952
            If SmallUpdate("1956", lDBVersion) Then DbCheckV1956
            If SmallUpdate("1957", lDBVersion) Then DbCheckV1957
            If SmallUpdate("1958", lDBVersion) Then DbCheckV1958
            If SmallUpdate("1959", lDBVersion) Then DbCheckV1959
            If SmallUpdate("1961", lDBVersion) Then DbCheckV1961
            If SmallUpdate("1962", lDBVersion) Then DbCheckV1962
            If SmallUpdate("1968", lDBVersion) Then DbCheckV1968
            If SmallUpdate("1969", lDBVersion) Then DbCheckV1969
            If SmallUpdate("1970", lDBVersion) Then DbCheckV1970
            If SmallUpdate("1979", lDBVersion) Then DbCheckV1979
            
            If SmallUpdate("1997", lDBVersion) Then DbCheckV1997
            If SmallUpdate("2006", lDBVersion) Then DbCheckV2006
            If SmallUpdate("2009", lDBVersion) Then DbCheckV2009
            If SmallUpdate("2026", lDBVersion) Then DbCheckV2026
            If SmallUpdate("2032", lDBVersion) Then DbCheckV2032

            If SmallUpdate("2042", lDBVersion) Then DbCheckV2042
            If SmallUpdate("2043", lDBVersion) Then DbCheckV2043
            If SmallUpdate("2048", lDBVersion) Then DbCheckV2048
            
            
'            If BigUpdate("2049", lDBVersion) Then DbCheckV2049: Kill gcDBPfad & "\upd2049.CFG"
            If SmallUpdate("2055", lDBVersion) Then DbCheckV2055
            If SmallUpdate("2073", lDBVersion) Then DbCheckV2073
            If SmallUpdate("2089", lDBVersion) Then DbCheckV2089
            If SmallUpdate("2091", lDBVersion) Then DbCheckV2091
'            If BigUpdate("2092", lDBVersion) Then DbCheckV2092: Kill gcDBPfad & "\upd2092.CFG"
            If SmallUpdate("2096", lDBVersion) Then DbCheckV2096
            If SmallUpdate("2111", lDBVersion) Then DbCheckV2111
            If SmallUpdate("2142", lDBVersion) Then DbCheckV2142
            If SmallUpdate("2153", lDBVersion) Then DbCheckV2153
            If SmallUpdate("2171", lDBVersion) Then DbCheckV2171
            If SmallUpdate("2173", lDBVersion) Then DbCheckV2173
            If SmallUpdate("2206", lDBVersion) Then DbCheckV2206
            If BigUpdate("2208", lDBVersion) Then DbCheckV2208: Kill gcDBPfad & "\upd2208.CFG"
            
            If SmallUpdate("2225", lDBVersion) Then DbCheckV2225
            If SmallUpdate("2234", lDBVersion) Then DbCheckV2234
            If SmallUpdate("2243", lDBVersion) Then DbCheckV2243
            If SmallUpdate("2248", lDBVersion) Then DbCheckV2248
            If SmallUpdate("2259", lDBVersion) Then DbCheckV2259
            If SmallUpdate("2293", lDBVersion) Then DbCheckV2293
            If SmallUpdate("2305", lDBVersion) Then DbCheckV2305
            If SmallUpdate("2307", lDBVersion) Then DbCheckV2307
            If SmallUpdate("2311", lDBVersion) Then DbCheckV2311
            If SmallUpdate("2318", lDBVersion) Then DbCheckV2318
            If SmallUpdate("2329", lDBVersion) Then DbCheckV2329
            If SmallUpdate("2335", lDBVersion) Then DbCheckV2335
            If SmallUpdate("2344", lDBVersion) Then DbCheckV2344
            If SmallUpdate("2364", lDBVersion) Then DbCheckV2364
            If SmallUpdate("2371", lDBVersion) Then DbCheckV2371
            If SmallUpdate("2405", lDBVersion) Then DbCheckV2405
            If SmallUpdate("2440", lDBVersion) Then DbCheckV2440
            If SmallUpdate("2449", lDBVersion) Then DbCheckV2449
            If SmallUpdate("2455", lDBVersion) Then DbCheckV2455
            If SmallUpdate("2463", lDBVersion) Then DbCheckV2463
            If SmallUpdate("2465", lDBVersion) Then DbCheckV2465
            If SmallUpdate("2467", lDBVersion) Then DbCheckV2467
            If SmallUpdate("2470", lDBVersion) Then DbCheckV2470
            If SmallUpdate("2474", lDBVersion) Then DbCheckV2474
            If SmallUpdate("2484", lDBVersion) Then DbCheckV2484
            If SmallUpdate("2494", lDBVersion) Then DbCheckV2494
            If SmallUpdate("2499", lDBVersion) Then DbCheckV2499
            If SmallUpdate("2505", lDBVersion) Then DbCheckV2505
            If SmallUpdate("2524", lDBVersion) Then DbCheckV2524
            If SmallUpdate("2528", lDBVersion) Then DbCheckV2528
            If SmallUpdate("2535", lDBVersion) Then DbCheckV2535
            If SmallUpdate("2539", lDBVersion) Then DbCheckV2539
            If SmallUpdate("2554", lDBVersion) Then DbCheckV2554
            If SmallUpdate("2555", lDBVersion) Then DbCheckV2555
            If SmallUpdate("2557", lDBVersion) Then DbCheckV2557
            If SmallUpdate("2564", lDBVersion) Then DbCheckV2564
            If SmallUpdate("2573", lDBVersion) Then DbCheckV2573
            If SmallUpdate("2579", lDBVersion) Then DbCheckV2579
            If SmallUpdate("2583", lDBVersion) Then DbCheckV2583
            If SmallUpdate("2593", lDBVersion) Then DbCheckV2593
            If SmallUpdate("2605", lDBVersion) Then DbCheckV2605
            If SmallUpdate("2607", lDBVersion) Then DbCheckV2607
            If SmallUpdate("2608", lDBVersion) Then DbCheckV2608
            If SmallUpdate("2631", lDBVersion) Then DbCheckV2631
            If SmallUpdate("2642", lDBVersion) Then DbCheckV2642
            If SmallUpdate("2647", lDBVersion) Then DbCheckV2647
            If SmallUpdate("2657", lDBVersion) Then DbCheckV2657
            If SmallUpdate("2658", lDBVersion) Then DbCheckV2658
            If SmallUpdate("2670", lDBVersion) Then DbCheckV2670
            If SmallUpdate("2688", lDBVersion) Then DbCheckV2688
            If SmallUpdate("2701", lDBVersion) Then DbCheckV2701
            If SmallUpdate("2709", lDBVersion) Then DbCheckV2709
            If SmallUpdate("2713", lDBVersion) Then DbCheckV2713
            If SmallUpdate("2715", lDBVersion) Then DbCheckV2715
            If SmallUpdate("2723", lDBVersion) Then DbCheckV2723
            If SmallUpdate("2724", lDBVersion) Then DbCheckV2724
            If SmallUpdate("2746", lDBVersion) Then DbCheckV2746
            If SmallUpdate("2747", lDBVersion) Then DbCheckV2747
            If SmallUpdate("2748", lDBVersion) Then DbCheckV2748
            If SmallUpdate("2758", lDBVersion) Then DbCheckV2758
            If SmallUpdate("2767", lDBVersion) Then DbCheckV2767
            If SmallUpdate("2769", lDBVersion) Then DbCheckV2769
            If SmallUpdate("2783", lDBVersion) Then DbCheckV2783
            If SmallUpdate("2784", lDBVersion) Then DbCheckV2784
            If SmallUpdate("2785", lDBVersion) Then DbCheckV2785
            If SmallUpdate("2787", lDBVersion) Then DbCheckV2787
            If SmallUpdate("2791", lDBVersion) Then DbCheckV2791
            If SmallUpdate("2795", lDBVersion) Then DbCheckV2795
            If SmallUpdate("2802", lDBVersion) Then DbCheckV2802
            If SmallUpdate("2807", lDBVersion) Then DbCheckV2807
            If SmallUpdate("2814", lDBVersion) Then DbCheckV2814
            If SmallUpdate("2820", lDBVersion) Then DbCheckV2820
            If SmallUpdate("2832", lDBVersion) Then DbCheckV2832
            If SmallUpdate("2833", lDBVersion) Then DbCheckV2833
            If SmallUpdate("2842", lDBVersion) Then DbCheckV2842
            If SmallUpdate("2881", lDBVersion) Then DbCheckV2881
            If SmallUpdate("2885", lDBVersion) Then DbCheckV2885
            If SmallUpdate("2892", lDBVersion) Then DbCheckV2892
            If SmallUpdate("2896", lDBVersion) Then DbCheckV2896
            If SmallUpdate("2909", lDBVersion) Then DbCheckV2909
            If SmallUpdate("2913", lDBVersion) Then DbCheckV2913
            If SmallUpdate("2922", lDBVersion) Then DbCheckV2922
            If SmallUpdate("2927", lDBVersion) Then DbCheckV2927
            If SmallUpdate("2931", lDBVersion) Then DbCheckV2931
            If SmallUpdate("2933", lDBVersion) Then DbCheckV2933
            If SmallUpdate("2935", lDBVersion) Then DbCheckV2935
            If SmallUpdate("2941", lDBVersion) Then DbCheckV2941
            If SmallUpdate("2968", lDBVersion) Then DbCheckV2968
            If SmallUpdate("2970", lDBVersion) Then DbCheckV2970
            If SmallUpdate("2977", lDBVersion) Then DbCheckV2977
            If SmallUpdate("2991", lDBVersion) Then DbCheckV2991
            If SmallUpdate("3001", lDBVersion) Then DbCheckV3001
            If SmallUpdate("3002", lDBVersion) Then DbCheckV3002
            If SmallUpdate("3004", lDBVersion) Then DbCheckV3004
            If SmallUpdate("3015", lDBVersion) Then DbCheckV3015
            If SmallUpdate("3018", lDBVersion) Then DbCheckV3018
            If SmallUpdate("3019", lDBVersion) Then DbCheckV3019
           
            If SmallUpdate(glpVers, lDBVersion) Then lDBVersion = glpVers

            If gbLokalModus Then
                Label2.Caption = "lokaler Modus"
                Label2.Refresh
            Else
                Label2.Caption = "Anwender aktiv"
                Label2.Refresh
            End If
        
            If gbDBMod Then
                URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/neuigkeiten.html"
                
'                zeigeHilfe "Update", "aneuheit" & ".doc", App.Path
            End If
        
            If lVersold < lDBVersion Then
                sSQL = "Update WKEINSTE Set DBVERSION = " & lDBVersion
                gdApp.Execute sSQL, dbFailOnError
    
                sSQL = "Update DBEINSTE set DBversion = " & lDBVersion
                gdBase.Execute sSQL, dbFailOnError
            End If
    End If
    'Versionszahlen in die jeweilige Tabelle schreiben*********
    
    If Not tableSuchenDBKombi("WKEINSTE", 2) Then
        CreateWKEINSTE
    Else
        sSQL = "Update WKEINSTE set Pversion = " & WKVersion
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    
    
    '********Versionszahlen schreiben Ende ********************
    
    If Not gbDEMO And Not gbKostenlos Then
        Label1(0).Caption = gsPname & " Version " & Left(WKVersion, 2) & "." & Right(WKVersion, 2)
        Label1(0).Refresh
        
        Label3.Caption = "Version " & Right(WKVersion, 4)
        Label3.Visible = True
        Label3.Refresh

    Else
        Label1(0).Caption = gsPname & " " & Left(WKVersion, 2) & "." & Right(WKVersion, 2)
        Label1(0).Refresh
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "CheckProgrammVersion"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub Deaktiviere_alleSchaltflächen()
On Error GoTo LOKAL_ERROR

    Dim i As Integer

    For i = 0 To Me.Controls.Count - 1
        If TypeOf Me.Controls(i) Is Command Then 'alle Commands
            Me.Controls(i).Enabled = False
        End If
    Next i
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Deaktiviere_alleSchaltflächen"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Positionieren()
    On Error GoTo LOKAL_ERROR
    
    With Frame1 'Stammdaten
        .Top = 1440
        .Left = 0
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame2 'Kasse
        .Top = 1440
        .Left = 720
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame3 'Etiketten
        .Top = 2150
        .Left = 3840
        .Height = 5295
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame4 'Listen
        .Top = 1440
        .Left = 3840
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame5 'Service
        .Top = 1440
        .Left = 7320
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame6 'Protokolle
        .Top = 1440
        .Left = 4560
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame7 'Termine
        .Top = 1440
        .Left = 5640
        .Height = 3855
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame8 'Einstellungen
        .Top = 1440
        .Left = 3480
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame9 'Artikellisten
         .Top = 1440
        .Left = 7680
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame10 'Kreditverkäufe
        .Top = 2140
        .Left = 4560
        .Height = 2415
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame11 'Bestellung
        .Top = 5050
        .Left = 3840
        .Height = 3135
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame12 'Statistiken
        .Top = 1440
        .Left = 2400
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame13 'Artikel
        .Top = 1440
        .Left = 3840
        .Height = 6015
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame14 'Datenbank
        .Top = 1440
        .Left = 3480
        .Height = 5295
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame15 'Lieferanten
        .Top = 5040
        .Left = 3840
        .Height = 3135
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame16 'Datenaustausch
        .Top = 5770
        .Left = 3480
        .Height = 2415
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame17 'KISSNET...
        .Top = 1440
        .Left = 3480
        .Height = 5295
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame18 'Stammdaten
        .Top = 2150
        .Left = 3840
        .Height = 3135
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame19 'Bestellung
        .Top = 4300
        .Left = 3840
        .Height = 2415
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame20 'Kundenliste
        .Top = 1440
        .Left = 7680
        .Height = 6015
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame21 'Detaildaten
        .Top = 2150
        .Left = 3840
        .Height = 3855
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame22 'Preislagenstatistiken
        .Top = 4300
        .Left = 6240
        .Height = 3855
        .Width = 3855
        .BorderStyle = 0
    End With
    
    With Frame23 'Verkaufsliste
        .Top = 1440
        .Left = 7680
        .Height = 6735
        .Width = 3855
        .BorderStyle = 0
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Drucke_Terminpreise()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim lHeute  As Long
    
    Screen.MousePointer = 11
    
    lHeute = Fix(Now)
    
    loeschNEW "TERMINPREISPRINT", gdBase
    CreateTableT2 "TERMINPREISPRINT", gdBase

    cSQL = "Insert into TERMINPREISPRINT select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", KVKPR1ALT "
    cSQL = cSQL & ", KVKPR1NEU "
    cSQL = cSQL & ", 'Start' as ART "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & " from PRSTERM where DAT_VON <= " & Trim$(Str$(lHeute)) & " and STATUS = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMINPREISPRINT t inner join ARTIKEL a on t.artnr = a.artnr "
    cSQL = cSQL & " set t.bestand = a.bestand "
    cSQL = cSQL & ", t.BEZEICH = a.BEZEICH "
    cSQL = cSQL & ", t.ean = a.ean "
    gdBase.Execute cSQL, dbFailOnError

    reportbildschirm "", "aWKL61d"

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_Terminpreise"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Drucke_Terminpreise_DEA()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim lHeute  As Long
    
    Screen.MousePointer = 11
    
    lHeute = Fix(Now)
    
    loeschNEW "TERMINPREISPRINT", gdBase
    CreateTableT2 "TERMINPREISPRINT", gdBase

    cSQL = "Insert into TERMINPREISPRINT select "
    cSQL = cSQL & " Artnr "
    cSQL = cSQL & ", KVKPR1ALT "
    cSQL = cSQL & ", KVKPR1NEU "
    cSQL = cSQL & ", 'Start' as ART "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & " from PRSTERM where DAT_BIS < " & Trim$(Str$(lHeute)) & " and STATUS = 1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMINPREISPRINT t inner join ARTIKEL a on t.artnr = a.artnr "
    cSQL = cSQL & " set t.bestand = a.bestand "
    cSQL = cSQL & ", t.BEZEICH = a.BEZEICH "
    cSQL = cSQL & ", t.ean = a.ean "
    gdBase.Execute cSQL, dbFailOnError

    reportbildschirm "", "aWKL61f"

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_Terminpreise_DEA"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AktiviereTerminPreiseWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim lHeute As Long
    Dim lartnr As Long
    Dim dKVkPr1Alt As Double
    Dim dKVkPr1Neu As Double
    
    Dim cBezeich As String
    Dim lBestand As Long
    Dim lAnzahl As Long
    Dim cLiBesNr As String
    Dim cEAN As String
    Dim lLinr As Long
    Dim lLpz As Long
    Dim dVkPr As Double
    
    Dim cRabattOk As String
    Dim cBonusOk As String
    Dim cPreisSchu As String
    
    ReDim caenderung(0 To 0) As String
    Dim cSatz As String
    Dim cFeld As String
    Dim lWert As Long
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    rsArt.index = "ARTNR"
    
    lHeute = Fix(Now)
    
    '************************************************************
    '* Setze Artikel auf terminierte Preis, wenn Datum erreicht
    '************************************************************
    
    lWert = MsgBox("Möchten Sie ein Liste der Preisänderungen ausdrucken?", vbYesNo + vbQuestion, "Winkiss Frage:")
    If lWert = vbYes Then
        Drucke_Terminpreise
    End If
   
    cSQL = "Select * from PRSTERM where DAT_VON <= " & Trim$(Str$(lHeute)) & " and STATUS = 0 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = -1
            End If
            If Not IsNull(rsrs!KVKPR1ALT) Then
                dKVkPr1Alt = rsrs!KVKPR1ALT
            Else
                dKVkPr1Alt = -1
            End If
            If Not IsNull(rsrs!KVKPR1NEU) Then
                dKVkPr1Neu = rsrs!KVKPR1NEU
            Else
                dKVkPr1Neu = -1
            End If
            
            If lartnr > -1 Then

                rsArt.Seek "=", lartnr
                If Not rsArt.NoMatch Then
                    If Not IsNull(rsArt!BEZEICH) Then
                        cBezeich = rsArt!BEZEICH
                    Else
                        cBezeich = ""
                    End If
                    
                    dVkPr = dKVkPr1Neu
                    
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                    
                    lAnzahl = lBestand
                    
                    If Not IsNull(rsArt!LIBESNR) Then
                        cLiBesNr = rsArt!LIBESNR
                    Else
                        cLiBesNr = ""
                    End If
                    
                    If Not IsNull(rsArt!EAN) Then
                        cEAN = rsArt!EAN
                    Else
                        cEAN = ""
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
                    
                    rsrs.Edit
                    rsrs!Status = 1
 
                    rsArt.Edit
                    rsArt!AWM = ermMerkFarbe(rsArt!artnr, "94")
                    DELMerkFarbe rsArt!artnr
                    rsArt!KVKPR1 = dKVkPr1Neu
                    rsArt!PREISSCHU = "J"
                    rsArt!RABATT_OK = "N"
                    
                    'wenn terminpreise bonusfähig dann bleibt die bonusfähigkeit erhalten
                    
                    If gbTPbf = False Then
                        'hier wird die bonusfähigkeit auf NEIN gesetzt
                        rsArt!BONUS_OK = "N"
                    End If

                        
                    rsrs.Update
                    rsArt.Update
                        
                    setzeFarbeinWK lartnr, "93"
                    
                    If lAnzahl <= 0 Then lAnzahl = 1
                    schreibeWKEtidru CStr(lartnr), lAnzahl, Val(gcFilNr)
                Else
                    rsrs.Edit
                    rsrs!Status = 99
                    rsrs.Update
                End If
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsArt.Close: Set rsArt = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktiviereTerminPreiseWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DEAktiviereTerminPreiseWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim lHeute As Long
    Dim lartnr As Long
    Dim dKVkPr1Alt As Double
    Dim dKVkPr1Neu As Double
    
    Dim cBezeich As String
    Dim lBestand As Long
    Dim lAnzahl As Long
    Dim cLiBesNr As String
    Dim cEAN As String
    Dim lLinr As Long
    Dim lLpz As Long
    Dim dVkPr As Double
    Dim cRabattOk As String
    Dim cBonusOk As String
    Dim cPreisSchu As String
    ReDim caenderung(0 To 0) As String
    Dim cSatz As String
    Dim cFeld As String
    Dim lWert As Long
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    rsArt.index = "ARTNR"
    
    lHeute = Fix(Now)
    
    lWert = MsgBox("Möchten Sie ein Liste der Preisänderungen ausdrucken?", vbYesNo + vbQuestion, "Winkiss Frage:")
    If lWert = vbYes Then
        Drucke_Terminpreise_DEA
    End If
    
    
    
    '****************************************************************
    '* Setze Artikel auf alten Preis, wenn Ende-Datum überschritten
    '****************************************************************
    
    cSQL = "Select * from PRSTERM where DAT_BIS < " & Trim$(Str$(lHeute)) & " and STATUS = 1 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = -1
            End If
            
            If Not IsNull(rsrs!KVKPR1ALT) Then
                dKVkPr1Alt = rsrs!KVKPR1ALT
            Else
                dKVkPr1Alt = -1
            End If
            
            If Not IsNull(rsrs!KVKPR1NEU) Then
                dKVkPr1Neu = rsrs!KVKPR1NEU
            Else
                dKVkPr1Neu = -1
            End If
            
            If Not IsNull(rsrs!RABATT_OK) Then
                cRabattOk = rsrs!RABATT_OK
            Else
                cRabattOk = "J"
            End If
            
            If Not IsNull(rsrs!BONUS_OK) Then
                cBonusOk = rsrs!BONUS_OK
            Else
                cBonusOk = "J"
            End If
            
            If Not IsNull(rsrs!PREISSCHU) Then
                cPreisSchu = rsrs!PREISSCHU
            Else
                cPreisSchu = "N"
            End If
            
            If lartnr > -1 Then

                rsArt.Seek "=", lartnr
            
                If Not rsArt.NoMatch Then
                    If Not IsNull(rsArt!BEZEICH) Then
                        cBezeich = rsArt!BEZEICH
                    Else
                        cBezeich = ""
                    End If
                    
                    dVkPr = dKVkPr1Alt
                    
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                    
                    lAnzahl = lBestand
                    
                    If Not IsNull(rsArt!LIBESNR) Then
                        cLiBesNr = rsArt!LIBESNR
                    Else
                        cLiBesNr = ""
                    End If
                    
                    If Not IsNull(rsArt!EAN) Then
                        cEAN = rsArt!EAN
                    Else
                        cEAN = ""
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
                    
                    rsrs.Edit
                    rsrs!Status = 99
                    
                    rsArt.Edit
                    rsArt!AWM = ermMerkFarbe(rsArt!artnr, "93")
                    DELMerkFarbe rsArt!artnr
                    rsArt!KVKPR1 = dKVkPr1Alt
                    rsArt!RABATT_OK = cRabattOk
                    rsArt!BONUS_OK = cBonusOk
                    rsArt!PREISSCHU = cPreisSchu
                    
                    rsrs.Update
                    
                    rsArt.Update
                    
                    schreibeWKEtidru CStr(lartnr), lAnzahl, Val(gcFilNr)
                Else
                    'wird in der Artikel nicht mehr gefunden - wie tragisch
                    rsrs.Edit
                    rsrs!Status = 99
                    rsrs.Update
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsArt.Close: Set rsArt = Nothing
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DEAktiviereTerminPreiseWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function BigUpdate(newVers As Integer, lVersion As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim ctmp       As String
    Dim k           As Integer
    
    
    BigUpdate = False
    
    If lVersion < newVers Then
nochmal:
        If BistDualleineinderDatenbank Then
        
            Label2.Caption = "Aktualisiere auf..."
            Label2.Refresh
            Label3.Caption = "Version " & newVers
            Label3.Visible = True
            Label3.Refresh
            
            BigUpdate = True
            gbDBMod = True
            bNeuheit = True
        Else
    
            If k > 9 Then
    
                SpielSound "update1"
                
                ctmp = "Sie haben ein neues Programmupdate erhalten. " & vbCrLf
                ctmp = ctmp & "Bitte schließen Sie an allen Computern alle Programme!" & vbCrLf
                ctmp = ctmp & "Starten Sie dann erneut Winkiss auf diesem Computer!" & vbCrLf & vbCrLf
    
                MsgBox ctmp, vbCritical, "Winkiss Programmupdate:"
                
                
                
    
                End 'Ende
            Else
                Pause 2
                k = k + 1
                GoTo nochmal
            End If
        
        End If
            
    End If
    
    DoEvents
        
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BigUpdate"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SmallUpdate(newVers As Integer, lVersion As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    SmallUpdate = False
    
    If lVersion < newVers Then
        Label2.Caption = "Aktualisiere auf... "
        Label2.Refresh
        Label3.Caption = "Version " & newVers
        Label3.Visible = True
        Label3.Refresh
        SmallUpdate = True
        bNeuheit = True
        gbDBMod = True
    End If
    DoEvents
       
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SmallUpdate"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub verarbeite_Sales_Daten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sDatei As String
    Dim iFileNr As Integer
    Dim cSatz1 As String
    Dim cEinzelsatz As String
    Dim lAktSatz As Long
    Dim lPosEnde As Long
    Dim lPos As Long
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lAnzSatz As Long
    Dim lAbzugsmenge As Long
    Dim lOldBestand As Long
    Dim cArtNr As String
    Dim ctemp As String
    Dim cWert As String
    Dim lposSemiEnde As Long
    
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "Sales*.txt"
    File1.Refresh
    
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sDatei = File1.list(i)
            
            lPos = 0
            
            iFileNr = FreeFile
            Open gsKinPfad & "\" & sDatei For Binary As #iFileNr
            If LOF(iFileNr) > 0 Then
            
                cSatz1 = Space$(LOF(iFileNr))
                Get #iFileNr, 1, cSatz1

                lLenfil = Len(cSatz1)
        
                lPos = 1
                lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                
                Do
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
        
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
                    lPos = lPos + lPosEnde - lPos + 2
        
                    lposSemi = 1
        
                   
                    cArtNr = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
                    cArtNr = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    lAbzugsmenge = Val(Right(cEinzelsatz, 5))
                    If lAbzugsmenge <> 0 Then
                        lOldBestand = ermBESTAND(cArtNr)
                        Bestandsveraenderung cArtNr, CLng(lOldBestand - lAbzugsmenge), "Shopverkauf"
                    End If
        
                Loop While lLenfil >= lPos
            
            

            Else
                
            End If
            
            Close iFileNr
            
            Kill gsKinPfad & "\" & sDatei
            
        Next i
        
    Else
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "verarbeite_Sales_Daten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub verarbeite_Sales_Daten_Rose()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sDatei As String
    Dim iFileNr As Integer
    Dim cSatz1 As String
    Dim cEinzelsatz As String
    Dim lAktSatz As Long
    Dim lPosEnde As Long
    Dim lPos As Long
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lAnzSatz As Long
    Dim lAbzugsmenge As Long
    Dim lOldBestand As Long
    Dim cArtNr As String
    Dim ctemp As String
    Dim cWert As String
    Dim lposSemiEnde As Long
    
    Dim cDatum As String
    Dim cUhrZeit As String
    Dim cEinzelPreis As String
    Dim cMenge As String
    
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "order_*.csv"
    File1.Refresh
    
    Dim lfail As Long
    Dim lRet As Long
    Dim cQuelle As String
    Dim cZiel As String
    Dim cZielSales As String
    
    cZielSales = ""
    
    cZiel = ""
    cZiel = gcDBPfad
    If Right(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    cZiel = cZiel & "SALES\"
    
    'existiert die Datei schon in SALES dann nicht verarbeiten und löschen
    
    If File1.ListCount > 0 Then
        For i = 0 To File1.ListCount - 1
            sDatei = File1.list(i)
            
            cQuelle = gsKinPfad & "\" & sDatei
            cZielSales = cZiel & sDatei
            
            If FileExists(cZielSales) Then
                Kill cQuelle
            End If
        Next i
    Else
        Exit Sub
    End If
    
    
    
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "order_*.csv"
    File1.Refresh

    cZielSales = ""
       
    
    If File1.ListCount > 0 Then
    
        For i = 0 To File1.ListCount - 1
            sDatei = File1.list(i)
            
            cQuelle = gsKinPfad & "\" & sDatei
            cZielSales = cZiel & sDatei
            Kill cZielSales
            lRet = CopyFile(cQuelle, cZielSales, lfail)
            
            lPos = 0
            
            iFileNr = FreeFile
            Open gsKinPfad & "\" & sDatei For Binary As #iFileNr
            If LOF(iFileNr) > 0 Then
            
                cSatz1 = Space$(LOF(iFileNr))
                Get #iFileNr, 1, cSatz1

                lLenfil = Len(cSatz1)
        
                lPos = 1
                lPosEnde = InStr(lPos, cSatz1, vbLf)
                
                lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile überspringen
                
                
                Do
                    lPosEnde = InStr(lPos, cSatz1, vbLf)
        
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
                    lPos = lPos + lPosEnde - lPos + 2
        
                    lposSemi = 1
        
                   
                    '"Order Id",DATUM,Time,Sku,Einzelpreis,StÃ¼ck
                    '61,2017-11-07,06:14:13,374.778,16.7200,1.0000
                    
                    '   item_id , "Create At", Sku, price, Qty, "Order id"
                    '   187,"2018-11-08 12:23:55",247202,32.0000,1.0000,99


                    cWert = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ",")
                    cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    cWert = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ",")
                    cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    cDatum = Mid(cWert, 10, 2) & "." & Mid(cWert, 7, 2) & "." & Mid(cWert, 2, 4)
                
                    cUhrZeit = Mid(cWert, 13, 8)
                    
                    cWert = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ",")
                    cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    cArtNr = SwapStr(cWert, ".", "")
                    
                    cWert = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ",")
                    cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    cWert = SwapStr(cWert, ".", ",")
                    cEinzelPreis = CStr(CDbl(cWert))
                    
                    cWert = ""
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ",")
                    cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                    lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    
                    
'                    cWert = Mid(cEinzelsatz, lposSemiEnde + 1, Len(cEinzelsatz) - lposSemiEnde + 1)
                    lAbzugsmenge = Val(cWert)
                    
                    If lAbzugsmenge <> 0 Then
                        lOldBestand = ermBESTAND(cArtNr)
                        Bestandsveraenderung cArtNr, CLng(lOldBestand - lAbzugsmenge), "Shopverkauf"
                        
                        Insert_Kassjour cArtNr, lAbzugsmenge, cUhrZeit, cDatum, cEinzelPreis
                        
                    End If
        
                Loop While lLenfil >= lPos
            
            Else
                
            End If
            
            Close iFileNr
            
            Kill gsKinPfad & "\" & sDatei
            
        Next i
        
    Else
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "verarbeite_Sales_Daten_Rose"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function newfnCheck4JaekelSalesDaten() As Long
    On Error GoTo LOKAL_ERROR
    
    newfnCheck4JaekelSalesDaten = 0
    
    
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "Sales*.txt"
    File1.Refresh
    
    If File1.ListCount > 0 Then
        'Dann verarbeiten
        verarbeite_Sales_Daten
    Else
        Exit Function
    End If
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4JaekelSalesDaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function newfnCheck4RoseSalesDaten() As Long
    On Error GoTo LOKAL_ERROR
    
    newfnCheck4RoseSalesDaten = 0
    
    File1.Path = gsKinPfad 'Standard In Pfad
    File1.Pattern = "order_*.csv"
    File1.Refresh
    
    If File1.ListCount > 0 Then
        'Dann verarbeiten
        verarbeite_Sales_Daten_Rose
    Else
        Exit Function
    End If
      
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4RoseSalesDaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function newfnCheck4KassenDateiWKL00() As Long
    On Error GoTo LOKAL_ERROR
    Dim ctmp As String
    
    newfnCheck4KassenDateiWKL00 = 0
    
    If gbFilNr And Val(gcFilNr) > 0 Then
        File1.Path = gsKinPfad
        File1.Pattern = "Y*.lzh"
        File1.Refresh
        
        If File1.ListCount > 0 Then
            gcKassenDatei = File1.list(1)
            
            If gbFTPautomatic Then
                gbfrm27 = True
                frmWKL27.Show 1
                
                newfnCheck4KassenDateiWKL00 = 1
            Else
                ctmp = "Es liegen neue Kassendateien vor. Möchten Sie diese jetzt einlesen?"
            
                dlgAbfrage.BCaptioneins = "Einlesen"
                dlgAbfrage.BCaptionzwei = "Abbrechen"
                dlgAbfrage.Überschrift = "Winkiss Hinweis:"
                dlgAbfrage.Beschriftung = ctmp
                dlgAbfrage.Show vbModal
                
                If dlgAbfrage.Back = 1 Then
                    gbfrm27 = True
                    frmWKL27.Show 1
                Else
                    gbfrm27 = False
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    If err.Number = 68 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "newfnCheck4KassenDateiWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Private Function newfnCheck4ReweStammdaten() As Long
    On Error GoTo LOKAL_ERROR
    Dim ctmp As String
    
    newfnCheck4ReweStammdaten = 0
    
    If Val(gcFilNr) = 0 Then
        File1.Path = gsKinPfad 'Standard In Pfad
        File1.Pattern = "REWE*Delta.csv"
        File1.Refresh
        
        If File1.ListCount > 0 Then
            If NewTableSuchenDBKombi("RORDER", gdBase) = False Then
                frmWKL185.Show 1
            End If
        Else
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4ReweStammdaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function newfnCheck4BudniStammdaten() As Long
    On Error GoTo LOKAL_ERROR
    Dim ctmp As String
    
    newfnCheck4BudniStammdaten = 0
    
    If Val(gcFilNr) = 0 Then
        File1.Path = gsKinPfad 'Standard In Pfad
        File1.Pattern = "BUDNI*.DRO"
        File1.Refresh
        
        If File1.ListCount > 0 Then
            If NewTableSuchenDBKombi("BUORDER", gdBase) = False Then
                frmWKL196.Show 1
            End If
        Else
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4BudniStammdaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function newfnCheck4CouponDaten() As Long
    On Error GoTo LOKAL_ERROR
    Dim ctmp As String

    newfnCheck4CouponDaten = 0

'    If Val(gcFilNr) = 0 Then
        File1.Path = gsKinPfad 'Standard In Pfad
        File1.Pattern = "*.xml"
        File1.Refresh

        If File1.ListCount > 0 Then
            If NewTableSuchenDBKombi("CORDER", gdBase) = False Then
                frmWKL194.Show 1
            End If
        Else
            Exit Function
        End If
'    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4CouponDaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function newfnCheck4LueningStammdaten() As Long
    On Error GoTo LOKAL_ERROR
    Dim ctmp As String
    
    newfnCheck4LueningStammdaten = 0
    
    If Val(gcFilNr) = 0 Then
        File1.Path = gsKinPfad 'Standard In Pfad
        File1.Pattern = "A00*.dat"
        File1.Refresh
        
        If File1.ListCount > 0 Then
            If NewTableSuchenDBKombi("LORDER", gdBase) = False Then
                frmWKL195.Show 1
            End If
        Else
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4LueningStammdaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function newfnCheck4StreckenStammdaten() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim i As Integer
    Dim sDateidatum As String
    Dim lHeute As Long
    Dim bFound As Boolean
    
    bFound = False
    lHeute = Fix(Now)
    
    newfnCheck4StreckenStammdaten = 0
    
    If Val(gcFilNr) = 0 Then
        File1.Path = gsKinPfad 'Standard In Pfad
        File1.Pattern = "Strecke*.csv"
        File1.Refresh
        
        If File1.ListCount > 0 Then
        
            For i = 0 To File1.ListCount - 1
                sDateidatum = Right(File1.list(i), 14)
                sDateidatum = Left(sDateidatum, 10)
                sDateidatum = SwapStr(sDateidatum, "-", ".")
                If IsDate(sDateidatum) = True Then
                    If lHeute >= CLng(DateValue(sDateidatum)) Then
                        bFound = True
                        Exit For
                    End If
                Else
                    bFound = True
                    Exit For
                End If
                
            Next i
        
        
            If bFound = True Then
                If NewTableSuchenDBKombi("SORDER", gdBase) = False Then
                    frmWKL186.Show 1
                End If
            End If
            
        Else
            Exit Function
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "newfnCheck4StreckenStammdaten"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'   Resume Next
End Function
Private Sub fnCheck_Filialtausch(sPattern As String)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "In"
    
    File1.Path = cPfad
    File1.Pattern = sPattern
    File1.Refresh
    
    If File1.ListCount > 0 Then
        gcKassenDatei = File1.list(1)
        
'        ctmp = "Es liegen neue Filialtäusche vor. Möchten Sie diese jetzt einlesen"
        
        ctmp = "Es liegen neue Filialtäusche vor. Möchten Sie den Programmteil zur Weiterverarbeitung dieser Dateien öffnen?"
    
        dlgAbfrage.BCaptioneins = "Öffnen"
        dlgAbfrage.BCaptionzwei = "Abbrechen"
        dlgAbfrage.Überschrift = "Winkiss Hinweis:"
        dlgAbfrage.Beschriftung = ctmp
        dlgAbfrage.Show vbModal
        
        If dlgAbfrage.Back = 1 Then
            frmWKL23.Show 1
        End If
    Else
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnCheck_Filialtausch"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function newfnCheck4LagerHauptgDateien() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim cLagerHauptdat  As String
    
    newfnCheck4LagerHauptgDateien = 0
    
    File1.Path = gsKinPfad
    File1.Pattern = "F*.lzh"
    File1.Refresh
    
    If File1.ListCount > 0 Then
    
        'Datei/en stehen an
        
        For i = 0 To File1.ListCount - 1
        
            cLagerHauptdat = File1.list(i)
            
            If ifThisDatinSteuerki(cLagerHauptdat) = False Then
            
                If verarbeiteLagerHauptdat(cLagerHauptdat) Then
                    lfnrSchreiben 0, Left(cLagerHauptdat, 8), DateValue(Now) & " " & TimeValue(Now)
                    Kill gsKinPfad & "\" & cLagerHauptdat
                    newfnCheck4LagerHauptgDateien = 1
                Else
                    Exit Function
                
                End If
            Else
                Kill gsKinPfad & "\" & cLagerHauptdat
            End If
        Next i
        
        'nur wenn die Verarbeitung erfolgreich war, dann wird eine externe Datenbank bereitgestellt
        schreibeProtokollNachtAblauf "externe Sicherung gestartet"
        picprogress.Visible = True
        Label3.Visible = True
        ExternSichern txtStatus, Label3
        picprogress.Visible = False
        schreibeProtokollNachtAblauf "externe Sicherung beendet"
        'Ende
        'nur wenn die Verarbeitung erfolgreich war, dann wird eine externe Datenbank bereitgestellt
    
    Else
        Exit Function
    End If
    
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "newfnCheck4LagerHauptgDateien"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Private Function verarbeiteLagerHauptdat(cdat As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad   As String
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim rsQ     As Recordset
    Dim rsZ     As Recordset
    Dim lNr     As Long
    
    verarbeiteLagerHauptdat = False
    
    Label3.Visible = True
    Label3.Caption = cdat & " wird entpackt..."
    Label3.Refresh
    
    picprogress.Visible = True
    ShowProgress picprogress, 0, 0, 0
    
'''''    Zip_Folder "", gsKinPfad, gsKinPfad & "\F0001717.lzh", txtStatus
    Zip_Unzip "", gsKinPfad, gsKinPfad & "\" & cdat, txtStatus
    
    If FileExists(gsKinPfad & "\FZ.mdb") Then
    
        loeschNEW "BESTA_in", gdBase
        
        cPfad = gsKinPfad & "\FZ.mdb"
        
        cSQL = "Select * into BESTA_in from BESTAKTO IN '" & cPfad & "'  "
        gdBase.Execute cSQL, dbFailOnError
        
        Set rsrs = gdBase.OpenRecordset("BESTA_IN", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    If Not IsNull(rsrs!Menge) Then
                        If Not IsNull(rsrs!AENART) Then
                            Bestandsveraenderung rsrs!artnr, ermBESTAND(rsrs!artnr) + (1 * CLng(rsrs!Menge)), rsrs!AENART
                        End If
                    End If
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        loeschNEW "afcbL", gdBase
        
        cPfad = gsKinPfad & "\FZ.mdb"
        cSQL = "Select * into afcbL from AFCBUCH IN '" & cPfad & "'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        Label3.Caption = "Kassenjournal wird geschrieben..."
        Label3.Refresh
        
        loeschNEW "KassL", gdBase
        
        cSQL = "select "
        cSQL = cSQL & " aartnr as artnr"
        cSQL = cSQL & " ,aBEZEICH as Bezeich"
        cSQL = cSQL & " ,aMenge as Menge "
        cSQL = cSQL & " ,aPreis as Preis"
        cSQL = cSQL & " ,adate"
        cSQL = cSQL & " ,azeit"
        cSQL = cSQL & " ,aMWSK as MWST"
        cSQL = cSQL & " ,abednu as BEDIENER"
        cSQL = cSQL & " ,akunum as kundnr"
        cSQL = cSQL & " ,filialnr as Filiale"
        cSQL = cSQL & " ,KASNUM"
        cSQL = cSQL & " ,linr"

        cSQL = cSQL & " ,0 as LPZ"
        cSQL = cSQL & " ,0 as AGN"
        cSQL = cSQL & " ,'' as EAN"

        cSQL = cSQL & " ,0 as ekpr"
        cSQL = cSQL & " ,UMS_OK"
        cSQL = cSQL & " ,0 as vkpr"
        cSQL = cSQL & " ,BELEGNR"
        cSQL = cSQL & " ,bestand as best1"
        cSQL = cSQL & " ,KK_ART"
        cSQL = cSQL & " into kassL from afcbL "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update  kassL inner join artikel on kassl.artnr = artikel.artnr "
        cSQL = cSQL & " set kassl.lpz = artikel.lpz "
        cSQL = cSQL & " , kassl.AGN = artikel.AGN "
        cSQL = cSQL & " , kassl.EAN = artikel.EAN "
        cSQL = cSQL & " , kassl.ekpr = artikel.ekpr "
        cSQL = cSQL & " , kassl.vkpr = artikel.vkpr "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Insert into kassjour select * from kassl "
        gdBase.Execute cSQL, dbFailOnError
        
        Label3.Caption = "Kreditverkäufe werden geschrieben..."
        Label3.Refresh
        
        '** Schritt 7: KREDIT füllen **
   
        cSQL = "Select * from afcbL"
        cSQL = cSQL & " where KK_art = 'KR' "
        Set rsQ = gdBase.OpenRecordset(cSQL)
        If Not rsQ.EOF Then
            cSQL = "Select * from KREDIT where ARTNR = -1 "
            Set rsZ = gdBase.OpenRecordset(cSQL)
            
            rsQ.MoveFirst
            Do While Not rsQ.EOF
                rsZ.AddNew
                If Not IsNull(rsQ!aartnr) Then
                    lNr = rsQ!aartnr
                End If
                rsZ!artnr = lNr
                rsZ!Menge = rsQ!aMenge
                rsZ!ekpr = 0
                
                If Not IsNull(rsQ!aMenge) Then
                    If rsQ!aMenge <> 0 Then
                        rsZ!vkpr = rsQ!APREIS / rsQ!aMenge
                    Else
                        rsZ!vkpr = 0
                    End If
                Else
                    rsZ!vkpr = 0
                End If
                
                rsZ!Kundnr = rsQ!AKUNUM
                rsZ!MWST = rsQ!AMWSK
                rsZ!GVKPR = rsQ!APREIS
                rsZ!ADATE = rsQ!ADATE
                rsZ!AVKPR = rsQ!AVKPR
                rsZ!PREISKZ = 0
                rsZ!FLAG = 0
                rsZ!BEZEICH = rsQ!ABEZEICH
                
                rsZ.Update
                
                rsQ.MoveNext
            Loop
            rsZ.Close: Set rsZ = Nothing
        End If
        
        rsQ.Close: Set rsQ = Nothing
        '*************
        cSQL = "Insert into Zugang Select * from ZUKas IN '" & cPfad & "'  "
        gdBase.Execute cSQL, dbFailOnError
        
        
        'Artikelmerkmal
        loeschNEW "ARTMERKi", gdBase
        
        cPfad = gsKinPfad & "\FZ.mdb"
        cSQL = "Select * into ARTMERKi from ARTMERKO IN '" & cPfad & "'  "
        gdBase.Execute cSQL, dbFailOnError
        
        Set rsrs = gdBase.OpenRecordset("ARTMERKi", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    If Not IsNull(rsrs!merk) Then
                        speichernMerkmal rsrs!artnr, rsrs!merk
                    End If
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        cSQL = "Delete from ARTMERK where Merk = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from ARTMERK where Merk is null "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Artmerk set sendok = true"
        gdBase.Execute cSQL, dbFailOnError
    End If

    Label3.Caption = "erfolgreich"
    Label3.Refresh
    
    Pause 1
    picprogress.Visible = False
    Label3.Visible = False
    Label3.Caption = ""
    Label3.Refresh
    
    verarbeiteLagerHauptdat = True
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "verarbeiteLagerHauptdat"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnCheck4MasterDateiWKL00() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim bFehler As Boolean
    Dim ctmp As String
        
    bFehler = False
    fnCheck4MasterDateiWKL00 = 0
    
    ctmp = gcDBPfad & "\IN"
    
    File1.Path = ctmp
    If bFehler Then
        ChDrive Left(gcDBPfad, 2)
        ChDir gcDBPfad
        MkDir "IN"
        ChDrive Left(gcPfad, 2)
        ChDir gcPfad
        File1.Path = gcDBPfad & "\IN"
    End If
    
    File1.Pattern = "MASTER!.*"
    File1.Refresh
    
    If File1.ListCount > 0 Then
        
        If Not Modul6.FindFile(gcPfad, "\REMEBER.TXT") Then
            
            iFileNr = FreeFile
            Open gcPfad & "Remeber.TXT" For Binary As iFileNr
            Close iFileNr
            fnCheck4MasterDateiWKL00 = 1
        Else
            fnCheck4MasterDateiWKL00 = 0
        End If
    Else
        iFileNr = FreeFile
        Open gcPfad & "Remeber.TXT" For Binary As iFileNr
        Close iFileNr
        Kill gcPfad & "Remeber.TXT"
    End If
    
Exit Function
LOKAL_ERROR:
    If err.Number = 76 Then
        bFehler = True
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "fnCheck4MasterDateiWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
    End If
End Function
Private Sub Check4PrsTerminWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lHeute As Long
    Dim lAnz As Long
    Dim iRet As Integer
    Dim ctmp As String
    
    Dim lAnzAKTIVIEREN As Long
    Dim lAnzDEAKTIVIEREN As Long
    Dim sNamePaAKTIVIEREN As String
    Dim sNamePaDEAKTIVIEREN As String
    
    lHeute = Fix(Now)
    
    cSQL = "Select * from PRSTERM where "
    cSQL = cSQL & "(DAT_VON <= " & Trim$(Str$(lHeute)) & " and STATUS = 0) "
    cSQL = cSQL & "or "
    cSQL = cSQL & "(DAT_BIS < " & Trim$(Str$(lHeute)) & " and STATUS = 1) "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
    Else
        lAnz = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lAnz > 0 Then
    
        ' hier machen wir alles neu
        ' Zuerst wird gefragt, welche Art (Aktivieren oder Deaktivieren)
        
        
        lAnzDEAKTIVIEREN = Check_Terminpreis_Art("DEAKTIVIEREN")
        If lAnzDEAKTIVIEREN > 0 Then
        
            sNamePaDEAKTIVIEREN = ermPaName("DEAKTIVIEREN")
            ctmp = "folgende Preisaktionen enden heute:" & vbCrLf & vbCrLf
            ctmp = ctmp & sNamePaDEAKTIVIEREN & vbCrLf
            ctmp = ctmp & "Die " & lAnzDEAKTIVIEREN & " Aktionspreise jetzt zurücksetzen?"
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Preisaktion/en Enden")
            If iRet = vbYes Then
                DEAktiviereTerminPreiseWKL00
            End If
        
        End If
        
        
        lAnzAKTIVIEREN = Check_Terminpreis_Art("AKTIVIEREN")
        If lAnzAKTIVIEREN > 0 Then
        
            sNamePaAKTIVIEREN = ermPaName("AKTIVIEREN")
            ctmp = "folgende Preisaktionen starten heute:" & vbCrLf & vbCrLf
            ctmp = ctmp & sNamePaAKTIVIEREN & vbCrLf
            ctmp = ctmp & "Die " & lAnzAKTIVIEREN & " Aktionspreise jetzt aktivieren?"
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Preisaktion/en Starten?")
            If iRet = vbYes Then
                AktiviereTerminPreiseWKL00
            End If
        
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4PrsTerminWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Check_Terminpreis_Art(sArt As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lHeute As Long
    Dim lAnz As Long
    Dim iRet As Integer
    Dim ctmp As String
    
    Check_Terminpreis_Art = 0
    lHeute = Fix(Now)
    
    cSQL = "Select count(*) as anz from PRSTERM  "
    
    If sArt = "AKTIVIEREN" Then
        cSQL = cSQL & " where (DAT_VON <= " & Trim$(Str$(lHeute)) & " and STATUS = 0) "
    ElseIf sArt = "DEAKTIVIEREN" Then
        cSQL = cSQL & " where (DAT_BIS < " & Trim$(Str$(lHeute)) & " and STATUS = 1) "
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!anz) Then
            Check_Terminpreis_Art = rsrs!anz
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check_Terminpreis_Art"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermPaName(sArt As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim lHeute As Long
    Dim lPNR As Long
    Dim lAnz As Long
    Dim iRet As Integer
    Dim ctmp As String
    Dim rsPA As DAO.Recordset
    
    ermPaName = ""
    lHeute = Fix(Now)
    
    cSQL = "Select distinct(preisnr) as PNR, count(artnr) as anz from PRSTERM "
    
    If sArt = "AKTIVIEREN" Then
        cSQL = cSQL & " where (DAT_VON <= " & Trim$(Str$(lHeute)) & " and STATUS = 0) "
    ElseIf sArt = "DEAKTIVIEREN" Then
        cSQL = cSQL & " where (DAT_BIS < " & Trim$(Str$(lHeute)) & " and STATUS = 1) "
    End If
    
    cSQL = cSQL & " group by preisnr "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!PNR) Then
                lPNR = rsrs!PNR
                
                lAnz = 0
                If Not IsNull(rsrs!anz) Then
                    lAnz = rsrs!anz
                End If
                
                cSQL = "Select * from PREISTERM where preisnr = " & lPNR
                Set rsPA = gdBase.OpenRecordset(cSQL)
                If Not rsPA.EOF Then
                
                    If Not IsNull(rsPA!preisname) Then
                        ermPaName = ermPaName & rsPA!preisname & " (" & lAnz & " Artikel)" & vbCrLf
                    End If
                    
                    If Not IsNull(rsPA!Von) Then
                        ermPaName = ermPaName & "Aktionszeitraum: " & rsPA!Von
                    End If
                    
                    If Not IsNull(rsPA!Bis) Then
                        ermPaName = ermPaName & " - " & rsPA!Bis & vbCrLf
                        
                    End If
                    ermPaName = ermPaName & "__________________________________________" & vbCrLf
                        
                End If
                rsPA.Close: Set rsPA = Nothing
                
            End If
            
            rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermPaName"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Tages_Bonus_zurückstellen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim iRet As Integer
    Dim ctmp As String
    
    ctmp = "Zur Zeit werden keine Kundenboni erhöht." & vbCrLf & vbCrLf
    ctmp = ctmp & "Möchten Sie diese Einstellung wieder löschen?"
    iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        cSQL = "Update Kassein set KUBONUS = True "
        gdBase.Execute cSQL, dbFailOnError
        
        gbKUBONUS = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tages_Bonus_zurückstellen"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tages_Bonus_zurückstellen_wenn_AG_RAB()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim iRet As Integer
    Dim ctmp As String
    
    ctmp = "Zur Zeit werden keine Kundenboni erhöht, wenn Artikel- bzw. Gesamtrabatt gewährt wird." & vbCrLf & vbCrLf
    ctmp = ctmp & "Möchten Sie diese Einstellung wieder löschen?"
    iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        cSQL = "Update Kassein Set NOKUBONUS_AGRAB = false "
        gdBase.Execute cSQL, dbFailOnError
        
        gbNoKUBONUS_wenn_Art_and_Ges_rab = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tages_Bonus_zurückstellen_wenn_AG_RAB"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tages_Bonus_Schwelle_zurückstellen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim iRet As Integer
    Dim ctmp As String
    
    ctmp = "Zur Zeit werden keine Kundenboni erhöht, wenn der Kassenverkaufspreis eines Artikels um " & gsiKUBONUS_SCHWELLE & "% und mehr kleiner ist als der Listenverkaufspreis." & vbCrLf & vbCrLf
    ctmp = ctmp & "Möchten Sie diese Einstellung wieder löschen?"
    iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        cSQL = "Update Kassein set KUBONUS_WENN = True "
        gdBase.Execute cSQL, dbFailOnError
        gbKUBONUS_WENN = True
        
        cSQL = "Update Kassein set KUBONUS_SCHWELLE = 0 "
        gdBase.Execute cSQL, dbFailOnError
        gsiKUBONUS_SCHWELLE = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tages_Bonus_Schwelle_zurückstellen"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnLeseIniDateiWKL00() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    Dim cPfad As String
    
    cdatei = "KISSLITE.INI"
    DabaPfadNew83

    cdatei = "KASNUM.CFG"
    
    If gcDBPfad <> "" Then
        fnLeseIniDateiWKL00 = 0
    Else
        fnLeseIniDateiWKL00 = 1
    End If
    
    iFileNr = FreeFile
    Open gcPfad & "KASNUM.CFG" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        gcKasNum = ctmp
        Close iFileNr
    Else
        Close iFileNr
        Kill gcPfad & "KASNUM.CFG"
        Do
            frmWKLab.Show 1
            iFileNr = FreeFile
            Open gcPfad & "KASNUM.CFG" For Binary As #iFileNr
            If LOF(iFileNr) > 0 Then
                ctmp = Space$(LOF(iFileNr))
                Get #iFileNr, 1, ctmp
                gcKasNum = ctmp
                Close iFileNr
                Exit Do
            End If
        Loop While LOF(iFileNr) = 0
    End If
    
    Do While Right(gcKasNum, 1) = vbCr Or Right(gcKasNum, 1) = vbLf
        gcKasNum = Left(gcKasNum, Len(gcKasNum) - 1)
    Loop
    
    

    Screen.MousePointer = 11
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 76 Or err.Number = 52 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "fnLeseIniDateiWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function


Private Function fnPruefeRegisterMitFirma() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cName As String
    Dim cPlz As String
    Dim cOrt As String
    
    fnPruefeRegisterMitFirma = 0
    
    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!name) Then
            cName = rsrs!name
        Else
            cName = ""
        End If
        cName = cName & Space$(35 - Len(cName))
        
        If Not IsNull(rsrs!Plz) Then
            cPlz = rsrs!Plz
        Else
            cPlz = ""
        End If
        cPlz = cPlz & Space$(7 - Len(cPlz))
        
        If Not IsNull(rsrs!Ort) Then
            cOrt = rsrs!Ort
        Else
            cOrt = ""
        End If
        cOrt = cOrt & Space$(30 - Len(cOrt))
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If cName <> gRegister.firma Then
        fnPruefeRegisterMitFirma = 1
    End If

    If cPlz <> gRegister.Plz Then
        fnPruefeRegisterMitFirma = 2
    End If

    If cOrt <> gRegister.Ort Then
        fnPruefeRegisterMitFirma = 3
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeRegisterMitFirma"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Sub WerHatheuteGeburtstag(iAnzTage As Integer)
    On Error GoTo LOKAL_ERROR
    
    If gcBonDrucker <> gcListenDrucker Then
        druckegebTageBondrucker iAnzTage, gbGebAdresse
        setzedrucker gcListenDrucker
    Else
        druckegebTageListendrucker iAnzTage, gbGebAdresse
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WerHatheuteGeburtstag"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub druckegebTageBondrucker(iTage As Integer, bMitAdresse As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sDate(99)   As String
    Dim sDateM      As String
    Dim cLBSatz     As String
    Dim iRet        As Integer
    Dim cKundnr     As String
    Dim cTel        As String
    Dim cMobil      As String
    Dim cDatum      As String
    Dim cVname      As String
    Dim cNName      As String
    Dim cTitel      As String
    Dim cPlz        As String
    Dim cStadt      As String
    Dim cStrasse    As String
    Dim cEmail      As String
    Dim i           As Integer
    Dim j           As Integer
    Dim iMax        As Integer
    
    ReDim cZeilen(0 To 8) As String
    
    For j = 0 To iTage
        sDate(j) = Left(DateValue(Now) + j, 5)
    Next j
    
    cSQL = "Select * from Kunden where "
    
    For j = 0 To iTage
        If j > 0 Then
            cSQL = cSQL & " or"
        End If
        cSQL = cSQL & " left(datum1,5) =  '" & sDate(j) & "'"
    Next j
    
    cSQL = cSQL & " order by month(datum1), day(datum1)"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        iRet = MsgBox("Wollen Sie wissen wer heute und morgen Geburtstag hat?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbNo Then
            rsrs.Close: Set rsrs = Nothing
            Exit Sub
        End If
        
        setzedrucker gcBonDrucker
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!name) Then
                cNName = rsrs!name
            Else
                cNName = ""
            End If
            If Not IsNull(rsrs!vorname) Then
                cVname = rsrs!vorname
            Else
                cVname = ""
            End If
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            Else
                cKundnr = "0"
            End If
            
            If Not IsNull(rsrs!Tel) Then
                cTel = rsrs!Tel
            Else
                cTel = ""
            End If
            
            If Not IsNull(rsrs!Mobiltel) Then
                cMobil = rsrs!Mobiltel
            Else
                cMobil = ""
            End If
            
            If Not IsNull(rsrs!titel) Then
                cTitel = rsrs!titel
            Else
                cTitel = ""
            End If
            
            If Not IsNull(rsrs!Datum1) Then
                cDatum = rsrs!Datum1
            End If
            
            
            If Not IsNull(rsrs!Plz) Then
                cPlz = rsrs!Plz
            Else
                cPlz = ""
            End If
            
            If Not IsNull(rsrs!STADT) Then
                cStadt = rsrs!STADT
            Else
                cStadt = ""
            End If
        
            If Not IsNull(rsrs!strasse) Then
                cStrasse = rsrs!strasse
            Else
                cStrasse = ""
            End If
            
            If Not IsNull(rsrs!Email) Then
                cEmail = rsrs!Email
            Else
                cEmail = ""
            End If
            
            
            
            
            
            
            If bMitAdresse Then
            
                For i = 0 To 8
                    cZeilen(i) = ""
                Next i
            
                cZeilen(0) = "KundNr: " & cKundnr & " Datum: " & cDatum
                cZeilen(1) = "Name: " & cTitel & " " & cVname & " " & cNName
                cZeilen(2) = "Telefon: " & cTel
                cZeilen(3) = "Mobil: " & cMobil
                cZeilen(4) = "Adresse: "
                cZeilen(5) = cPlz & " " & cStadt
                cZeilen(6) = cStrasse
                cZeilen(7) = cEmail
                cZeilen(8) = " "
                iMax = 8
            
            Else
            
                For i = 0 To 3
                    cZeilen(i) = ""
                Next i
            
                cZeilen(0) = "KundNr: " & cKundnr & " Datum: " & cDatum
                cZeilen(1) = "Name: " & cTitel & " " & cVname & " " & cNName
                cZeilen(2) = "Telefon: " & cTel
                cZeilen(3) = "Mobil: " & cMobil
                cZeilen(4) = " "
                iMax = 4
            
            End If
           
            rsrs.MoveNext
            DruckeEndlosBeleg cZeilen(), iMax, rsrs.EOF
            
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "druckegebTageBondrucker"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub druckegebTageListendrucker(iTage As Integer, bMitAdresse As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sDate(99)   As String
    Dim iRet        As Integer
    Dim j           As Integer
    
    For j = 0 To iTage
        sDate(j) = Left(DateValue(Now) + j, 5)
    Next j

'    sDate = Left(DateValue(Now), 5)
    
    loeschNEW "KUGEB", gdBase
    CreateTableT2 "KUGEB", gdBase
    
    cSQL = "Insert into KUGEB Select "
    cSQL = cSQL & " Kundnr"
    cSQL = cSQL & ", titel "
    cSQL = cSQL & ", name "
    cSQL = cSQL & ", vorname"
    cSQL = cSQL & ", Tel"
    cSQL = cSQL & ", Mobiltel"
    cSQL = cSQL & ", Datum1"
    If bMitAdresse Then
        cSQL = cSQL & ", PLZ"
        cSQL = cSQL & ", STADT"
        cSQL = cSQL & ", STRASSE"
    Else
        cSQL = cSQL & ", '' as PLZ"
        cSQL = cSQL & ", '' as STADT"
        cSQL = cSQL & ", '' as STRASSE"
    End If
    cSQL = cSQL & " from Kunden where "
    
    For j = 0 To iTage
        If j > 0 Then
            cSQL = cSQL & " or"
        End If
        cSQL = cSQL & " left(datum1,5) =  '" & sDate(j) & "'"
    Next j
    
    cSQL = cSQL & " order by month(datum1), day(datum1)"
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("KUGEB")
    If Not rsrs.EOF Then
        iRet = MsgBox("Wollen Sie wissen wer heute und die nächsten " & iTage & " Tage Geburtstag hat?", vbQuestion + vbYesNo, "Winkiss Frage:")
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If iRet = vbYes Then
    
        If bMitAdresse Then
            reportbildschirm "", "aWKL00bb"
        Else
            reportbildschirm "", "aWKL00ba"
        End If
        Pause 2
    End If
    loeschNEW "KUGEB", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "druckegebTageListendrucker"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub LeseMWStSaetzeWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim UstId As Integer
    
    'check, ob die Tabelle MWSTSATZ zum zweiten mal schon erweitert wurde
     If SpalteInTabellegefundenNEW("MWSTSATZ", "vonD", gdBase) = False Then
     
         'check, ob die Tabelle MWSTSATZ zum ersten mal schon erweitert wurde
         If SpalteInTabellegefundenNEW("MWSTSATZ", "id", gdBase) Then
            'die erste Erweiterung abbrechen
            cSQL = "DELETE FROM MWSTSATZ"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Alter table MWSTSATZ drop column id"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Alter table MWSTSATZ drop column FurJahr"
            gdBase.Execute cSQL, dbFailOnError
         
         Else
            'MWSTSATZ wurde vorher nie erweitert, deswegen alle Datensätze davon in Vorbereitung auf der Erweiterung entfernen
            cSQL = "DELETE FROM MWSTSATZ"
            gdBase.Execute cSQL, dbFailOnError
            
         End If
         
         'MWSTSATZ erweitern
         cSQL = "Alter table MWSTSATZ add column id NUMBER"
         gdBase.Execute cSQL, dbFailOnError
         
         cSQL = "Alter table MWSTSATZ add column vonD DATE"
         gdBase.Execute cSQL, dbFailOnError
         
         cSQL = "Alter table MWSTSATZ add column bisD DATE"
         gdBase.Execute cSQL, dbFailOnError
         
         'MWST-Werte der vorherigen Jahren hinzufügen
         VorherigeMwstWerteHinzufugen
         
     End If
    
    '01.01.2100 muss auf Null gesetzt werden, wenn dieses Datum von ExportFormular(DsFinvK) für Zwischenrechnen-Zwecks benutzt wurde
    gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
    
    cSQL = "Select * from MWSTSATZ WHERE vonD>= CDate('" & Date & "') AND bisD<= CDate('" & Date & "') AND bisD <> NULL"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!VOLL) Then
            gdMWStV = rsrs!VOLL
        Else
            gdMWStV = 0
        End If
        
        If Not IsNull(rsrs!ERM) Then
            gdMWStE = rsrs!ERM
        Else
            gdMWStE = 0
        End If
        
        If Not IsNull(rsrs!OHNE) Then
            gdMWStO = rsrs!OHNE
        Else
            gdMWStO = 0
        End If
        
        If Not IsNull(rsrs!id) Then
            UstId = rsrs!id
        Else
            UstId = 0
        End If
        
    Else
      gdMWStV = 0
      gdMWStE = 0
      gdMWStO = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    cSQL = "Select * from MWSTSATZ WHERE CDate('" & Date & "')>= vonD AND bisD is NULL"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!VOLL) Then
            gdMWStV = rsrs!VOLL
        Else
            gdMWStV = 0
        End If
        
        If Not IsNull(rsrs!ERM) Then
            gdMWStE = rsrs!ERM
        Else
            gdMWStE = 0
        End If
        
        If Not IsNull(rsrs!OHNE) Then
            gdMWStO = rsrs!OHNE
        Else
            gdMWStO = 0
        End If
        
        If Not IsNull(rsrs!id) Then
            UstId = rsrs!id
        Else
            UstId = 0
        End If
    Else
      gdMWStV = 0
      gdMWStE = 0
      gdMWStO = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
     If gdMWStV = 0 Then
      MsgBox "Es konnten keine MWST-Sätze gelesen werden!", vbCritical, "Winkiss Hinweis:"
     End If
     
    
    'Tabelle [MWSTSATZDiesesJahrs] erstellen und insertiere darin die Mwst.-Satz dieses Jahrs (dieser Schritt ist nötig für manche Reports wie z.b aWKL25ab.rpt, aWKL25ai.rpt)
    If Not NewTableSuchenDB("MWSTSATZDiesesJahrs", gdBase) Then

        cSQL = "SELECT " & gdMWStV & " as VOLL," & gdMWStE & " as ERM," & gdMWStO & " as OHNE," & UstId & " as id into MWSTSATZDiesesJahrs"
        gdBase.Execute cSQL, dbFailOnError

    Else

        cSQL = "DROP TABLE MWSTSATZDiesesJahrs"
        gdBase.Execute cSQL, dbFailOnError

        cSQL = "SELECT " & gdMWStV & " as VOLL," & gdMWStE & " as ERM," & gdMWStO & " as OHNE," & UstId & " as id into MWSTSATZDiesesJahrs"
        gdBase.Execute cSQL, dbFailOnError

    End If
    
    '01.01.2100 muss auf Null gesetzt werden, weil dieses Datum nur in ExportFormular(DsFinvK) benutzt wird
    gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseMWStSaetzeWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub

Private Sub VorherigeMwstWerteHinzufugen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    
    cSQL = "INSERT INTO MWSTSATZ (VOLL,ERM,OHNE,id,vonD,bisD)VALUES('16','7','0','1','01.04.1998','31.12.2006')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "INSERT INTO MWSTSATZ (VOLL,ERM,OHNE,id,vonD,bisD)VALUES('19','7','0','2','01.01.2007','30.06.2020')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "INSERT INTO MWSTSATZ (VOLL,ERM,OHNE,id,vonD,bisD)VALUES('16','5','0','3','01.07.2020','31.12.2020')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "INSERT INTO MWSTSATZ (VOLL,ERM,OHNE,id,vonD,bisD)VALUES('19','7','0','4','01.01.2021',NULL)"
    gdBase.Execute cSQL, dbFailOnError
    
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VorherigeMwstWerteHinzufugen"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub

Private Sub LeseTexteKassenBonWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iZeilenNr As Integer
    Dim cZeilenText As String
    
    gcBonText(0) = ""
    gcBonText(1) = ""
    gcBonText(2) = ""
    gcBonText(3) = ""
    gcBonText(4) = ""
    gcBonText(5) = ""
    gcBonText(6) = ""
    gcBonText(7) = ""
    
    cSQL = "Select * from BONTEXT order by ZEILENNR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ZEILENNR) Then
                iZeilenNr = rsrs!ZEILENNR
            Else
                iZeilenNr = 0
            End If
            If Not IsNull(rsrs!ZEILENTEXT) Then
                cZeilenText = rsrs!ZEILENTEXT
            Else
                cZeilenText = ""
            End If
            cZeilenText = Trim$(cZeilenText)
            gcBonText(iZeilenNr) = cZeilenText
        
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseTexteKassenBonWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeRegistrierungWKL00() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim lSize As Long
    Dim lWert As Long
    Dim cSysPfad As String
    Dim cDaten As String
    Dim iRet As Integer
    
    fnPruefeRegistrierungWKL00 = 0
    
    cSysPfad = Space$(255)
    lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
    cSysPfad = Left(cSysPfad, lWert)
    If Right(cSysPfad, 1) <> "\" Then
        cSysPfad = cSysPfad & "\"
    End If


    
    gcSysPfad = cSysPfad
    
    iFileNr = FreeFile
    If gbDebug Then
        MsgBox cSysPfad
    End If
'    MsgBox cSysPfad & gcRegDatei

    


    Open cSysPfad & gcRegDatei For Binary As #iFileNr
    lSize = LOF(iFileNr)
    
    If lSize = 0 Then
        Close iFileNr
        Kill cSysPfad & gcRegDatei
        frmWKL01.Show 1
        
        iFileNr = FreeFile
        Open cSysPfad & gcRegDatei For Binary As #iFileNr
        lSize = LOF(iFileNr)
        If lSize = 0 Then
            Close iFileNr
            Kill cSysPfad & gcRegDatei
            fnPruefeRegistrierungWKL00 = 1
        Else
            cDaten = Space$(112)
            Get #iFileNr, 212, cDaten
            Close iFileNr
            'MsgBox cDaten
            
            cDaten = fnDecrypt(cDaten)
            ZerlegeRegisterDatenWKL00 cDaten
            cDaten = gRegister.KdWert3 & gRegister.KdWert4 & gRegister.Confirm3 & gRegister.Confirm4
            iRet = fnPruefeRegisterMitFirma()
            
            fnPruefeRegistrierungWKL00 = iRet
        End If
    Else
        cDaten = Space$(112)
        Get #iFileNr, 212, cDaten
        Close iFileNr
        cDaten = fnDecrypt(cDaten)
        ZerlegeRegisterDatenWKL00 cDaten
        
        cDaten = gRegister.KdWert3 & gRegister.KdWert4 & gRegister.Confirm3 & gRegister.Confirm4
        iRet = fnPruefeRegisterMitFirma()
        fnPruefeRegistrierungWKL00 = iRet
        
    End If
    
Exit Function
LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "PruefeRegistrierungWKL00"
'    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
    
    
'    If FileExists(cSysPfad & gcRegDatei) = False Then
        MsgBox "Bitte das Programm 'Als Administrator ausführen' und die Registrierung vornehmen!", vbInformation + vbOKOnly, "Winkiss Hinweis:"
        
        End
'    End If
    
    
    
End Function
Private Sub StammdatenZu()
On Error GoTo LOKAL_ERROR

    Command2_Click 7
    Command4_Click 3
    Command4_Click 5
    Command4_Click 8
    Command12_Click 2
    Command14_Click 4
    Command14_Click 0
    Command12_Click 22
    
    Command1(0).ForeColor = vbBlack
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "stammdatenzu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KasseZu()
On Error GoTo LOKAL_ERROR

    Command3_Click 6
    Command7_Click 7
    Command11_Click 2
    
    Command1(1).ForeColor = vbBlack
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KasseZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub StatistikenZu()
On Error GoTo LOKAL_ERROR

    Command13_Click 5
    Command4_Click 15
    
    Command1(2).ForeColor = vbBlack
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StatistikenZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ListenZu()
On Error GoTo LOKAL_ERROR

    Command5_Click 9
    Command10_Click 5
    Command12_Click 25
    Command12_Click 36
    
    Command1(3).ForeColor = vbBlack

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ListenZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ServiceZu()
On Error GoTo LOKAL_ERROR

    Command6_Click 8
    Command9_Click 5
    Command12_Click 5
    Command12_Click 10
    Command12_Click 19
    
    Command1(4).ForeColor = vbBlack
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ServiceZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub allesZu()
On Error GoTo LOKAL_ERROR



KasseZu
StatistikenZu
ServiceZu
ListenZu
StammdatenZu
TermineZu
                
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "allesZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TermineZu()
On Error GoTo LOKAL_ERROR

    Command8_Click 3
    
    Command1(8).ForeColor = vbBlack
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TermineZu"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub StarteSubMenue(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    Select Case index
        Case Is = 0 'Stammdaten
        
            If bStammdaten Then
            
                StammdatenZu
                bStammdaten = False
            Else
                bStammdaten = True
            
            
                KasseZu
                StatistikenZu
                ServiceZu
                ListenZu
'                StammdatenZu
                TermineZu

                Frame1.Visible = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
            
                    Command1(0).Enabled = True
                    Command1(0).ForeColor = &HC00000
                    
                    Command2(0).SetFocus
                End If
                
            End If
        Case Is = 1 'Kasse
            If bKasse Then
                KasseZu
                
                bKasse = False
            Else
                bKasse = True
                If gbBargeldEingabe = True Then
                    Command3(7).Enabled = False
                Else
                    Command3(7).Enabled = True
                End If
                
                If gbAABSCHL = True Then
                    Command3(1).Enabled = False
                Else
                    Command3(1).Enabled = True
                End If
                
'                KasseZu
                StatistikenZu
                ServiceZu
                ListenZu
                StammdatenZu
                TermineZu
                Frame2.Visible = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                
                    Command1(1).Enabled = True
                    Command1(1).ForeColor = &HC00000
                    
                    Command3(0).SetFocus
                End If
            End If
            
        Case Is = 2 'Statistiken
        
            If bStatistiken Then
                StatistikenZu
                bStatistiken = False
            Else
                bStatistiken = True
                
                KasseZu
'                StatistikenZu
                ServiceZu
                ListenZu
                StammdatenZu
                TermineZu
                
                Frame12.Visible = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                
                    Command1(2).Enabled = True
                    Command1(2).ForeColor = &HC00000
                    Command13(0).SetFocus
                End If
            End If
            
        Case Is = 3 'Listen
            If bListen Then
                ListenZu
                
                bListen = False
            Else
                bListen = True
                        
                KasseZu
                StatistikenZu
                ServiceZu
'                ListenZu
                StammdatenZu
                TermineZu
                
                Frame4.Visible = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    Command1(3).Enabled = True
                    Command1(3).ForeColor = &HC00000
                    Command5(0).SetFocus
                End If
            End If

        Case Is = 4 'Service
            If bService Then
                ServiceZu
                
                bService = False
            Else
                bService = True
                        
                If gcFilNr = 0 Then
                    Command6(2).Enabled = False
                Else
                    Command6(2).Enabled = True
                End If
                
                KasseZu
                StatistikenZu
'                ServiceZu
                ListenZu
                StammdatenZu
                TermineZu
                
                Frame5.Visible = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    Command1(4).Enabled = True
                    Command1(4).ForeColor = &HC00000
                    Command6(0).SetFocus
                End If
            End If
            
        Case Is = 5
            'Ende der Anwendung
            
        Case Is = 6
            'Anmelden
            If gbBEDKARTE Then
                fAnmeldung
            Else
                frmWKL99.Show 1
            End If
            
        Case Is = 7
            'Abmelden
            If gbBEDKARTE Then
                fAbmeldung
            End If
            
            gcUserName = ""
            gcPass = ""
            glLevel = -1
            
            If gbLokalModus Then
                frmWKL00!Label2.Visible = True
                frmWKL00!Label2.ForeColor = vbRed
                frmWKL00!Label2.Caption = "lokaler Modus (Datenbank nicht erreichbar) - Anwender nicht aktiv"
                frmWKL00!Label2.Refresh
            Else
                frmWKL00!Label2.Visible = True
                frmWKL00!Label2.Caption = "Anwender nicht aktiv"
                frmWKL00!Label2.Refresh
            End If

        Case Is = 8 'Termine
            If bTermine Then
                TermineZu
                bTermine = False
            Else
                bTermine = True
                
                KasseZu
                StatistikenZu
                ServiceZu
                ListenZu
                StammdatenZu
'                TermineZu
                Frame7.Visible = True
                
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    Command1(8).Enabled = True
                    Command1(8).ForeColor = &HC00000
                    Command8(0).SetFocus
                End If
            End If
        Case 9  'Kassennummer ändern
            frmWKLab.Show 1
            
            If Val(gcKasNum) > 0 And Val(gcKasNum) < 10 Then
                Command1(9).BackColorTo = glfarbe(gcKasNum)
                Command1(9).BackColorFrom = glfarbe(gcKasNum)
            End If
            
            Command1(9).Caption = gcKasNum
            
            If ermaktUmsatz(False) > 0 Then
                frmWKL00.Command3(1).BackColor = vbRed
            Else
                frmWKL00.Command3(1).BackColor = &H8000000F
            End If
    
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StarteSubMenue"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub ZerlegeRegisterDatenWKL00(cDaten As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lPosMerk As Long
    Dim lPos As Long
    Dim ctmp As String
    Dim iStufe As Integer
    
    gRegister.firma = ""
    gRegister.Plz = ""
    gRegister.Ort = ""
    gRegister.KdWert1 = ""
    gRegister.KdWert2 = ""
    gRegister.KdWert3 = ""
    gRegister.KdWert4 = ""
    gRegister.Confirm1 = ""
    gRegister.Confirm2 = ""
    gRegister.Confirm3 = ""
    gRegister.Confirm4 = ""
    
    iStufe = 0
    
    lPosMerk = 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 1
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    ctmp = ctmp & Space$(35 - Len(ctmp))
    gRegister.firma = ctmp
    
    iStufe = 2
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 3
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    ctmp = ctmp & Space$(7 - Len(ctmp))
    gRegister.Plz = ctmp
    
    iStufe = 4
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 5
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    ctmp = ctmp & Space$(30 - Len(ctmp))
    gRegister.Ort = ctmp
    
    iStufe = 6
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 7
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.KdWert1 = ctmp
    
    iStufe = 8
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 9
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.KdWert2 = ctmp
    
    iStufe = 10
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 11
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.KdWert3 = ctmp
    
    iStufe = 12
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 13
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.KdWert4 = ctmp
    
    iStufe = 14
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 15
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.Confirm1 = ctmp
    
    iStufe = 16
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 17
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.Confirm2 = ctmp
    
    iStufe = 18
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 19
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.Confirm3 = ctmp
    
    iStufe = 20
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    iStufe = 21
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.Confirm4 = ctmp
    
    iStufe = 22
    lPosMerk = lPos + 1
    lPos = InStr(lPosMerk, cDaten, Chr$(27))
    
    iStufe = 23
    
    ctmp = Mid(cDaten, lPosMerk, lPos - lPosMerk)
    gRegister.Datum = ctmp
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 5 And iStufe = 23 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZerlegeRegisterDatenWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten. " & iStufe
        
        Fehlermeldung1
    End If
    
End Sub
Private Sub ChkLM_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim ctemp As String
    
    picprogress.Visible = True
    txtStatus.Text = "5"
    Label2.Caption = "Moment bitte ... Vorbereitung..."
    Label2.Refresh
    
    
    If ChkLM.Caption = "In den lokalen Modus umschalten." And ChkLM.value = vbChecked Then
    
        If Not Modul6.FindFile("C:\aLeer", "kissdata.mdb") Then
            ChkLM.value = vbUnchecked
            Exit Sub
        
        End If
        
        'globale Variable setzen
        gbLokalModus = True
        
        'lokal.cfg schreiben
        If gbLokalModus Then
            gcDBPfad = "C:\aLeer"
            PruefeSubDir
            iFileNr = FreeFile
            Open gcPfad & "Lokal.CFG" For Binary As #iFileNr
        
            If LOF(iFileNr) > 0 Then
                Close iFileNr
            Else
                ctemp = "Rechner befindet sich zur Zeit im lokalen Modus"
                Put #iFileNr, 1, ctemp
                Close iFileNr
            End If
        End If
        
        'aussehen verändern
        
        If gbLokalModus Then
            
            
            Command1(0).Enabled = False
            
            Command1(1).Enabled = True
            Command1(2).Enabled = False
            Command1(3).Enabled = False
            Command1(4).Enabled = False
            Command1(8).Enabled = False
            
            Command3(10).Enabled = False
            Command3(9).Enabled = False
            Command3(8).Enabled = False
            Command3(7).Enabled = False
            Command3(4).Enabled = False
            Command3(5).Enabled = False
            
            If gbLocalSec Then
                If gbAutoLokalModus Then
                    Command3(1).Caption = "Z Bon"
                Else
                    Command3(1).Caption = "Tagesbericht"
                End If
            End If

        Else
        
            Command1(0).Enabled = True
            Command1(1).Enabled = True
            Command1(2).Enabled = True
            Command1(3).Enabled = True
            Command1(4).Enabled = True
            Command1(8).Enabled = True
            
            Command3(1).Caption = "Tagesbericht"
            
        End If
        
        'Datenbank auf lokale Datenbank setzen
        
''        AbmeldungDabaNew
        gdBase.Close
        Set gdBase = Nothing
        gcDBPfad = "C:\aLeer"
        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
        
        
        anzeige "normal", "Offline - Betrieb", frmWKL00.Label2
        
        
        If gbLokalModus Then
            'Datenbank Synchronisieren einstellen
            ChkLM.Caption = "Datenbank synchronisieren"
            ChkLM.value = vbUnchecked
            Command1(1).SetFocus
        End If
        

        
    ElseIf ChkLM.Caption = "Datenbank synchronisieren" And ChkLM.value = vbChecked Then
    
    
    
        'globale Variable zurücksetzen
        gbLokalModus = False
        
        'Datenbank setzen
        gdBase.Close
        DabaPfadNew83
        
        If gcDBPfad <> "" Then
            If Not Modul6.FindFile(gcDBPfad, "kissdata.mdb") Then
            
                
                
                gbLokalModus = True '***Hier bleiben wir im lokalen Modus weil
                                    '***Haupdatenbank nicht erreichbar
                gcDBPfad = "C:\aLeer"
                Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
                frmWKL00!Label2.ForeColor = vbRed
                frmWKL00!Label2.Caption = "lokaler Modus(Datenbank nicht erreichbar) - Anwender aktiv"
                frmWKL00!Label2.Refresh
                ChkLM.value = vbUnchecked
                Exit Sub
                
            Else
            
                
            End If
        Else
            gbLokalModus = True '***Hier bleiben wir im lokalen Modus weil
                                '***Haupdatenbank nicht erreichbar
            gcDBPfad = "C:\aLeer"
            Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
            frmWKL00!Label2.ForeColor = vbRed
            frmWKL00!Label2.Caption = "lokaler Modus (Datenbank nicht erreichbar) - Anwender aktiv"
            frmWKL00!Label2.Refresh
            ChkLM.value = vbUnchecked
            Exit Sub
        End If
        
        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
'        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)

        
        If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        
            gbLokalModus = True '***Hier bleiben wir im lokalen Modus weil
                            '***Haupdatenbank nicht erreichbar
            gcDBPfad = "C:\aLeer"
            Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
            frmWKL00!Label2.ForeColor = vbRed
            frmWKL00!Label2.Caption = "lokaler Modus(Datenbank nicht erreichbar) - Anwender aktiv"
            frmWKL00!Label2.Refresh
            ChkLM.value = vbUnchecked
            Exit Sub
               
        End If
        
        
        If gbLocalSec Then
            If gbAutoLokalModus Then
                If gbBONWG Then
                    synchronisiereDB
                    HoleLokalDB
                Else
'                    ABSCHIEBENDB
                    
                    synchronisiereDB
                    HoleLokalDB

                End If

                
            Else
            
                synchronisiereDB
                HoleLokalDB
                
            End If
        End If

'        synchronisiereDB 'Datenbank synchronisieren
        
        
        
        'lokal.cfg löschen
        Kill gcPfad & "Lokal.CFG"
        
        'aussehen verändern
        
        If gbLokalModus Then
            
            Command1(0).Enabled = False
            Command1(1).Enabled = True
            Command1(2).Enabled = False
            Command1(3).Enabled = False
            Command1(4).Enabled = False
            Command1(8).Enabled = False
            
            Command3(10).Enabled = False
            Command3(9).Enabled = False
            Command3(8).Enabled = False
            Command3(7).Enabled = False
            Command3(4).Enabled = False
            Command3(5).Enabled = False
            
            If gbLocalSec Then
                If gbAutoLokalModus Then
                    Command3(1).Caption = "Z Bon"
                Else
                    Command3(1).Caption = "Tagesbericht"
                End If
            End If
           
            
        Else
        
            Command1(0).Enabled = True
            Command1(1).Enabled = True
            Command1(2).Enabled = True
            Command1(3).Enabled = True
            Command1(4).Enabled = True
            Command1(8).Enabled = True
            
            Command3(10).Enabled = True
            Command3(9).Enabled = True
            Command3(8).Enabled = True
            Command3(7).Enabled = True
            Command3(4).Enabled = True
            Command3(5).Enabled = True
            
            Command3(1).Caption = "Tagesbericht"
        End If
        
        If gbLokalModus = False Then
            'Datenbank für den lokalen modus einstellen
            ChkLM.Caption = "In den lokalen Modus umschalten."
            ChkLM.value = vbUnchecked
            ChkLM.Visible = False
'            Command1(0).SetFocus
        End If
        
        frmWKL00!Label2.Caption = "Anwender aktiv"
        frmWKL00!Label2.Refresh
        
    End If
    
    picprogress.Visible = False
    
    Label2.Caption = ""
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3011 Then
    
        gbLokalModus = True '***Hier bleiben wir im lokalen Modus weil
                            '***Haupdatenbank nicht erreichbar
        gcDBPfad = "C:\aLeer"
        Set gdBase = OpenDatabase(gcDBPfad & "\kissdata.mdb", False)
        frmWKL00!Label2.ForeColor = vbRed
        frmWKL00!Label2.Caption = "lokaler Modus(Datenbank nicht erreichbar) - Anwender aktiv"
        frmWKL00!Label2.Refresh
        ChkLM.value = vbUnchecked
        Exit Sub
    
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ChkLM_Click"
        Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If index <> 5 Then
        Label2.Visible = False
        Label3.Visible = False
        Label1(1).Visible = False
        StarteSubMenue index
    Else
        Unload frmWKL00
        End 'Ende
    End If
    
    If index = 1 Then
        If gbLokalModus Then

            Command1(0).Enabled = False
            Command1(1).Enabled = True
            Command1(2).Enabled = False
            Command1(3).Enabled = False
            Command1(4).Enabled = False
            Command1(8).Enabled = False
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command1_Click 5
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_KeyUp"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    If index = 7 And Button = 2 And Shift = 1 Then
        Unload frmWKLab
        frmWKLab.Show 1
    End If
    
    If index = 1 And Button = 2 And Shift = 1 Then
        If gbLocalSec Then
            ChkLM.Visible = True
            If gbLokalModus Then
                ChkLM.Caption = "Datenbank synchronisieren"
            End If
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Setzedatenbankpfad()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim lAnz    As Long
    Dim ctmp    As String
    Dim sPfad   As String
    Dim cPfad       As String

    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Speichern des Datenbankpfades"
        .InitDir = "C:\"
'        .InitDir = sOldpfad
        .Filter = "Access - Dateien (*.mdb)|kissdata.mdb"
        .ShowSave
        End With
        sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
        
        iFileNr = FreeFile
        Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
        Close iFileNr
        Kill gcPfad & "KISSLITE.INI"
        
        Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
        Put #iFileNr, 1, sPfad
        Close iFileNr
        gcDBPfad = sPfad
        
        cPfad = gcDBPfad 'Datenbankpfad
        If Right(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        
        If Modul6.FindFile(gcDBPfad, "\kissdata.mdb") Then
            Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)

'            ermLastVK
            
            ctmp = "Die Datenbank ist erfolgreich geöffnet worden!" & vbCrLf
            ctmp = ctmp & "Der letzte Verkauf fand statt am:" & vbCrLf & vbCrLf
            ctmp = ctmp & ermLastVK & vbCrLf & vbCrLf
            ctmp = ctmp & "Ist dies die richtige Datenbank?"
            
            MsgBox ctmp, vbInformation, "Winkiss Hinweis:"
            
'            bei Erfolg auch die anderen Pfade verbiegen
            gsUpdPfad = gcDBPfad & "\In"
            gsZinPfad = gcDBPfad & "\Kissdata.mdb"
            gsKinPfad = gcDBPfad & "\In"
            gsZoutPfad = gcDBPfad & "\Kassout"
            gsDabaPfad = gcDBPfad
            gsSicherPfad = gcDBPfad & "\Sicherung"
            speicherpfad
            

    
            If gbDEMO Then
                lAnz = fnPruefeAnzahlArtikelWKL00()
                If lAnz > 200 Then
                    ctmp = "Die Datenbank enthält mehr als 200 Artikel." & vbCrLf & vbCrLf
                    ctmp = ctmp & "Für die freie Version von Winkiss sind maximal 200 Artikel zugelassen." & vbCrLf & vbCrLf
                    ctmp = ctmp & "Das Programm kann nicht gestartet werden!"
                    MsgBox ctmp, vbCritical, gsPname & " Hinweis:"
                    Command12_Click 4
                End If
            End If
            
            
            gcFilNr = "-1"
            
            Dim rsrs As DAO.Recordset

            Set rsrs = gdBase.OpenRecordset("Fila", dbOpenTable)
            If Not rsrs.EOF Then
                gcFilNr = rsrs!fil
            End If
            rsrs.Close: Set rsrs = Nothing
        
            gbFilNr = False
            If Val(gcFilNr) > -1 Then
                gbFilNr = True
            End If
            
            Label1(13).Caption = "F " & gcFilNr
            

            
            
        Else
            ctmp = "Es kann keine Verbindung zur Datenbank aufgebaut werden!" & vbCrLf & vbCrLf
            ctmp = ctmp & "Bitte setzen Sie sich mit uns über unsere Hotline 0511/955910 in Verbindung. " & vbCrLf & vbCrLf
            MsgBox ctmp, vbCritical, "Winkiss - Hotline anrufen:"
            Command12_Click 4
        End If
        
        PruefeRegistryEintragDataMOD01

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Setzedatenbankpfad"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
err:

    MsgBox "Das Programm wird jetzt beendet!", vbInformation, "Winkiss Hinweis:"
    End
End Sub

Private Function ermittleTag(cmdX As sevCommand3.Command) As Byte
    On Error GoTo LOKAL_ERROR
    
    ermittleTag = 255
    If cmdX.Tag <> "" Then
        ermittleTag = CByte(cmdX.Tag)
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleTag"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Command14_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 1, 2, 3, 4, 7, 8, 9, 11
                Command14_Click 4
            Case 0, 5, 6, 10
                Command14_Click 0
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command14_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

  
Private Sub Command15_Click()
On Error GoTo LOKAL_ERROR
' TestZwecks.Show 1

'an EDEKA FTP-Server eine Test-Bestellung schicken  <<<<<<<<<<<<<<<<<<< START
' giKissFtpMode = 49
' frmWKL38.Show 1
'an EDEKA FTP-Server eine Test-Bestellung schicken  <<<<<<<<<<<<<<<<<<< ENDE

'setzedrucker gcListenDrucker & "-oday"
'reportbildschirmToPrinter "aWKL21i"

' MsgBox (DatePart("ww", DateValue(Now), vbMonday, vbFirstFourDays)) '---> hat damals 44 geliefert
' MsgBox (DatePart("ww", DateValue(Now))) '--->  hat damals 45 geliefert
  
' MsgBox ("von :" & gZeiten(7).Von & "bis :" & gZeiten(7).Bis)
  
  
  
'  Dim obb As Object
'  Set obb = GetObject("winmgmts:") _
'        .ExecQuery("select * from win32_process where name='meineSchnitt.exe'")
'
'  If obb.Count > 0 Then
'   MsgBox ("Running ...")
'  End If
   
    
Exit Sub
    
LOKAL_ERROR:
 MsgBox (err.Number & vbNewLine & vbNewLine & err.Description)
End Sub

Private Sub Command7_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 4, 5, 6, 7, 0, 1, 2, 3
                Command7_Click 7
            
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 6, 7, 8, 9, 10
                Command3_Click 6
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Command10_Click 5
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command13_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Command13_Click 5
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 17, 18
                Command4_Click 3
            Case 4, 8, 9, 10, 11
                Command4_Click 8
            Case 5, 8, 9, 10, 11
                Command4_Click 5
            Case 12, 13, 14, 15, 16
                Command4_Click 15
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command11_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2
                Command11_Click 2
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Command6_Click 8
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Command9_Click 5
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 5, 6, 7, 8, 9
                Command5_Click 9
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Command2_Click 7
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 0, 1, 2, 3, 4
                Command8_Click 3
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command12_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Select Case index
            Case 4, 5, 6, 7, 8, 12, 13
                Command12_Click 5
            Case 14, 15, 16, 17, 18, 19, 20
                Command12_Click 19
            Case 9, 10, 11
                Command12_Click 10
            Case 21, 22, 23
                Command12_Click 22
            Case 0, 1, 2, 3
                Command12_Click 2
            Case 24, 25, 26, 27, 28, 29, 30
                Command12_Click 25
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39
                Command12_Click 36
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command12_KeyUp"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctemp As String
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command2(index))
        Select Case index
            Case Is = 0     'Artikel-Stammdaten
                Frame13.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command14(1).SetFocus
                End If
            Case Is = 1     'Stammdaten-Änderung
                Frame18.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command4(7).SetFocus
                End If
            Case Is = 2     'Artikelgruppe
                If glLevel >= ermittlezugriff(byteZGNr) Then
                    Frame21.Visible = True
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = False
                    Next lcount
                    If gbKostenlos = False Then
                        Command4(10).SetFocus
                    End If
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If

            Case Is = 3     'Kundendaten
                OpenProgrammTeil frmWKL13, ermittlezugriff(byteZGNr)
            Case Is = 4     'Bestellvorschläge
            
                If glLevel >= ermittlezugriff(byteZGNr) Then
                
                    Frame19.Visible = True
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = False
                    Next lcount
                    If gbKostenlos = False Then
                        Command12(21).SetFocus
                    End If
            
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            
            Case Is = 5     'Wareneingang
            
                Frame11.Visible = True
                If gcFilNr = 0 Then
                    Command12(1).Enabled = False
                Else
                    Command12(1).Enabled = True
                End If
                
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                
                If gbKostenlos = False Then
                    Command12(3).SetFocus
                End If
                                    
            Case Is = 6     'Lieferanten
                Frame15.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                
                If gbKostenlos = False Then
                    Command14(6).SetFocus
                End If
            
            Case Is = 7
                For lcount = 0 To 6
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame1.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bStammdaten = False
                
            Case Is = 8 'Etiketten
                
                
                Frame3.Visible = True
                If FileExists(App.Path & "\naegr.rpt") Then
                    Command4(18).Enabled = True
                Else
                    Command4(18).Enabled = False
                End If
                
                
                For lcount = 0 To 8
                
                    Command2(lcount).Enabled = False
                Next lcount
                
                If gbKostenlos = False Then
                    Command4(0).SetFocus
                End If
        End Select
    Else
        Select Case index
            Case Is = 0     'Artikel-Stammdaten
                Frame13.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command14(1).SetFocus
                End If
            Case Is = 1     'Stammdaten-Änderung
                Frame18.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command4(7).SetFocus
                End If
            Case Is = 2     'Artikelgruppe
                Frame21.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                
                If gbKostenlos = False Then
                    Command4(10).SetFocus
                End If
            Case Is = 3     'Kundendaten
    
                If glLevel >= DlgZugriff(3).dZugriff Then
                    frmWKL13.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 4     'Bestellvorschläge
                
                If glLevel >= DlgZugriff(12).dZugriff Then
                    Frame19.Visible = True
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = False
                    Next lcount
                    If gbKostenlos = False Then
                        Command12(21).SetFocus
                    End If
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
                
                   
                
                
            Case Is = 5     'Wareneingang
            
                Frame11.Visible = True
                If gcFilNr = 0 Then
                    Command12(1).Enabled = False
                Else
                    Command12(1).Enabled = True
                End If
                
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command12(3).SetFocus
                End If
            Case Is = 6     'Lieferanten
                Frame15.Visible = True
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command14(6).SetFocus
                End If
            
            Case Is = 7
                For lcount = 0 To 6
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame1.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bStammdaten = False
                
            Case Is = 8 'Etiketten
                
                Frame3.Visible = True
                If FileExists(App.Path & "\naegr.rpt") Then
                    Command4(18).Enabled = True
                Else
                    Command4(18).Enabled = False
                End If
                
                For lcount = 0 To 8
                    Command2(lcount).Enabled = False
                Next lcount
                
                If gbKostenlos = False Then
                    Command4(0).SetFocus
                End If
        End Select
    End If
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub OpenProgrammTeil(frmx As Form, byZugriffsrecht As Byte)
    On Error GoTo LOKAL_ERROR
    Dim ctemp As String
    
    If glLevel >= byZugriffsrecht Then
    
        schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
        
        Unload frmx
        
        frmx.Show 1
    Else
        ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
        ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
        schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
        MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
        
        fAnmeldung
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "OpenProgrammTeil"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten. " & glLevel & " " & byZugriffsrecht & " " & frmx.name
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim rsrs As Recordset
    Dim cSQL As String
    Dim ctemp As String
    Dim iRet As Integer
    Dim cPfad As String
    Dim iZaehler As Integer
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command3(index))
        Select Case index
            Case Is = 0     'Kassieren
                Screen.MousePointer = 11
                OpenProgrammTeil frmWKL20, ermittlezugriff(byteZGNr)
                
                
                If gbLocalSec Then
                    If gbAutoLokalModus Then
                        If gbAutoSYN Then
                            'Syn
                            frmWKL00.ChkLM.value = vbChecked
                            'wieder in den A LM
                            frmWKL00.ChkLM.value = vbChecked
                        End If
                    End If
                End If
                
                
                If gbLokalModus Then
                    Command1(0).Enabled = False
                    Command1(1).Enabled = True
                    Command1(2).Enabled = False
                    Command1(3).Enabled = False
                    Command1(4).Enabled = False
                    Command1(8).Enabled = False
                End If
                
            Case Is = 1     'Tagesbericht
            
                cPfad = gcDBPfad
                If Right(cPfad, 1) <> "\" Then
                    cPfad = cPfad & "\"
                End If
                
                iZaehler = 1
            
            
                Do While FileExists(cPfad & "KASSSTOP_ALLE.TXT")
                    Pause 1
                    gsAnzeigeText = "Diese Funktion kann erst in ein paar Sekunden gestartet werden," & vbCrLf
                    gsAnzeigeText = gsAnzeigeText & "da an einem anderen Rechner ein KASSENABSCHLUSS durchgeführt wird." & vbCrLf
            
                    
                    MsgBox gsAnzeigeText, vbInformation, iZaehler & ". Meldung"
                    Pause 1
                    iZaehler = iZaehler + 1
    
                Loop
            
                bAbschlussjetzt = False
                If glLevel >= ermittlezugriff(byteZGNr) Then
                    If gbLokalModus Then
                        If gbLocalSec Then
                            If gbAutoLokalModus Then
                                If Command3(1).Caption = "Z Bon" Then
                                
                                    DabaPfadNew84
                                    cPfad = gcDBPfad    'Datenbankpfad
                                    Do While IsAktionZulaessig_Kassenabschluss_LM = False

                                    Loop
                                    DabaPfadNew83
                                
                                    iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                    If iRet = vbYes Then
                                    
                                        frmWKL21.LeseDatenWKL21
                                        frmWKL00.ChkLM.value = vbChecked
                                        
                                        If gbBargeldEingabe = True Then
                                            
                                            SchubladeOeffnen
                                            frmWK21b.Show
                                            
                                        Else
                                            AktionAustragen "Kassenabschluss"
                                            frmWKL21.LeseDatenWKL21
                                            frmWKL21.Show 1
                                            Me.Refresh
                                        End If
                                        AktionAustragen "Kassenabschluss"
                
                                    Else
                                        DabaPfadNew84
                                        AktionAustragen "Kassenabschluss"
                                        DabaPfadNew83
                                        
                                        frmWKL21.LeseDatenWKL21 'X bon aus autolokalmodus
                                        frmWKL21.Show 1
                                        Me.Refresh
                                        Exit Sub
                                    End If
                                Else
                                    iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                    If iRet = vbYes Then
                                        frmWKL00.ChkLM.value = vbChecked
                                    Else
                                        
                                    End If
                                    
                                    If gbBargeldEingabe = True Then
                                        iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                        If iRet = vbYes Then
                                            SchubladeOeffnen
                                            frmWK21b.Show
                                        Else
                                            
                                        End If
                                    Else
                                        frmWKL21.LeseDatenWKL21
                                        frmWKL21.Show 1
                                        Me.Refresh
                                    End If
                                    
                                    AktionAustragen "Kassenabschluss"
                                End If
                                
                            Else
                                Exit Sub
                                
                            End If
                        End If
                    Else
                        If gbBargeldEingabe = True Then
                        
                            iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                            If iRet = vbYes Then
                                SchubladeOeffnen
                                
                                frmWK21b.Show 1
                                'Bargeldart
'                                Select Case giBARGELDART
'                                    Case 0
'                                        frmWK21b.Show 1
'                                    Case 1
'                                        frmWK21s.Show 1
'                                    Case 2
'                                        frmWK21t.Show 1
'                                End Select
                            Else
                                
                            End If

                        Else
                            frmWKL21.LeseDatenWKL21
                            frmWKL21.Show 1
                            Me.Refresh
                        End If
                    End If
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                    
                    fAnmeldung
                    
                End If
                
            Case Is = 4     'Kreditverkäufe
                   
                    Frame10.Visible = True
                    Frame2.Enabled = False
                    Command3(0).Enabled = False
                    Command3(1).Enabled = False
                    Command3(2).Enabled = False
                    Command3(3).Enabled = False
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(6).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
    
            Case Is = 5     'Protokolle
                    
                    Frame6.Visible = True
                    Frame2.Enabled = False
                    Command3(0).Enabled = False
                    Command3(1).Enabled = False
                    Command3(2).Enabled = False
                    Command3(3).Enabled = False
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(6).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                
            Case Is = 6     'zurück zum Hauptmenue
            
                If gbLokalModus Then
                    Command1(5).Enabled = True
                Else
                
                    For lcount = 0 To 5
                        Command1(lcount).Enabled = True
                    Next lcount
                    Command1(8).Enabled = True
                End If
        
                Frame2.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bKasse = False
    
            Case Is = 7     'Bargeld zählen
                OpenProgrammTeil frmWK21b, ermittlezugriff(byteZGNr)
                
                If gbLokalModus Then
    
                    Command1(0).Enabled = False
                    Command1(1).Enabled = True
                    Command1(2).Enabled = False
                    Command1(3).Enabled = False
                    Command1(4).Enabled = False
                    Command1(8).Enabled = False
                End If
                
                
            Case Is = 8     'Kassenbuch
                OpenProgrammTeil frmWKL128, ermittlezugriff(byteZGNr)
            Case Is = 9     'Zusammenfassung z Bon
                OpenProgrammTeil frmWK21f, ermittlezugriff(byteZGNr)
            Case Is = 10    'Kundenbestellung
                OpenProgrammTeil frmWKL141, ermittlezugriff(byteZGNr)
        End Select
        
    Else
        Select Case index
            Case Is = 0     'Kassieren
                If glLevel >= DlgZugriff(7).dZugriff Then
                    Screen.MousePointer = 11
                    frmWKL20.Show 1
                    
                    If gbLocalSec Then
                        If gbAutoLokalModus Then
                            If gbAutoSYN Then
                                frmWKL00.ChkLM.value = vbChecked
                                frmWKL00.ChkLM.value = vbUnchecked
                            End If
                        End If
                    End If
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
                If gbLokalModus Then
    
                    Command1(0).Enabled = False
                    Command1(1).Enabled = True
                    Command1(2).Enabled = False
                    Command1(3).Enabled = False
                    Command1(4).Enabled = False
                    Command1(8).Enabled = False
                    
                End If
                
            Case Is = 1     'Tagesbericht

                cPfad = gcDBPfad
                If Right(cPfad, 1) <> "\" Then
                    cPfad = cPfad & "\"
                End If
                
                iZaehler = 1
            
            
                Do While FileExists(cPfad & "KASSSTOP_ALLE.TXT")
                    Pause 1
                    gsAnzeigeText = "Diese Funktion kann erst in ein paar Sekunden gestartet werden," & vbCrLf
                    gsAnzeigeText = gsAnzeigeText & "da an einem anderen Rechner ein KASSENABSCHLUSS durchgeführt wird." & vbCrLf
            
                    
                    MsgBox gsAnzeigeText, vbInformation, iZaehler & ". Meldung"
                    Pause 1
                    iZaehler = iZaehler + 1
    
                Loop

            
            
            
            
            
                bAbschlussjetzt = False
                
                If glLevel >= DlgZugriff(8).dZugriff Then
                    If gbLokalModus Then
                        If gbLocalSec Then
                            If gbAutoLokalModus Then
                            
                                If Command3(1).Caption = "Z Bon" Then
                                
                                    iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                    If iRet = vbYes Then
                                        frmWKL00.ChkLM.value = vbChecked
                                    Else
                                        Exit Sub
                                    End If
                                    
                                    
                                    If gbBargeldEingabe = True Then
                                        iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                        If iRet = vbYes Then
                                            SchubladeOeffnen
                                            frmWK21b.Show 1
                                        Else
                                            
                                        End If

                                    Else
                                        frmWKL21.LeseDatenWKL21
                                        frmWKL21.Show 1
                                        
                                        Me.Refresh
                                    End If
                                    
                                Else
                                    iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                    If iRet = vbYes Then
                                        frmWKL00.ChkLM.value = vbChecked
                                    Else
                                        
                                    End If
                                    
                                    If gbBargeldEingabe = True Then
                                        iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                                        If iRet = vbYes Then
                                            SchubladeOeffnen
                                            frmWK21b.Show 1
                                        Else
                                            
                                        End If

                                    Else
                                        frmWKL21.LeseDatenWKL21
                                        frmWKL21.Show 1
                                        Me.Refresh
                                    End If
                                    
                                End If
                            Else
                                Exit Sub
                                
                            End If
                        End If
                    Else
                        If gbBargeldEingabe = True Then
                            iRet = MsgBox("Möchten Sie jetzt Ihren Kassenabschluss vornehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                            If iRet = vbYes Then
                                SchubladeOeffnen
                                frmWK21b.Show 1
                            Else
                                
                            End If

                        Else
                            frmWKL21.LeseDatenWKL21
                            frmWKL21.Show 1
                            Me.Refresh
                        End If
                    End If
                    
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If

            Case Is = 4     'Kreditverkäufe
                If glLevel >= DlgZugriff(10).dZugriff Then
                    
                   
                    Frame10.Visible = True
                    Frame2.Enabled = False
                    Command3(0).Enabled = False
                    Command3(1).Enabled = False
                    Command3(2).Enabled = False
                    Command3(3).Enabled = False
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(6).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 5     'Protokolle
                If glLevel >= DlgZugriff(10).dZugriff Then
                    Frame6.Visible = True
                    Frame2.Enabled = False
                    Command3(0).Enabled = False
                    Command3(1).Enabled = False
                    Command3(2).Enabled = False
                    Command3(3).Enabled = False
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(6).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 6     'zurück zum Hauptmenue
            
                If gbLokalModus Then
                    Command1(5).Enabled = True
                Else
                
                    For lcount = 0 To 5
                        Command1(lcount).Enabled = True
                    Next lcount
                    Command1(8).Enabled = True
                End If
        
                Frame2.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bKasse = False
    
            Case Is = 7     'Bargeld zählen
                If glLevel >= DlgZugriff(12).dZugriff Then
                    frmWK21b.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
                
                
                If gbLokalModus Then
    
                    Command1(0).Enabled = False
                    Command1(1).Enabled = True
                    Command1(2).Enabled = False
                    Command1(3).Enabled = False
                    Command1(4).Enabled = False
                    Command1(8).Enabled = False
                End If
                
                
                
                
            Case Is = 8     'Kassenbuch
                If glLevel >= DlgZugriff(12).dZugriff Then
                    frmWKL128.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 9     'Zusammmenfassung z bon
                If glLevel >= DlgZugriff(12).dZugriff Then
                    frmWK21f.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 10    'Kundenbestellung
                If glLevel >= DlgZugriff(12).dZugriff Then
                    frmWKL141.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
        End Select
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(index As Integer)
On Error GoTo LOKAL_ERROR
        
    Dim lcount As Long
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command4(index))
        
        Select Case index
            Case Is = 0     'Strichcode-Etiketten
                OpenProgrammTeil frmWKL30, ermittlezugriff(byteZGNr)
            Case Is = 1     'Strichcode-Etiketten selbst wählen
                OpenProgrammTeil frmWKL31, ermittlezugriff(byteZGNr)
            Case Is = 2     'Etiketten aus Lieferschein
                OpenProgrammTeil frmWKL03, ermittlezugriff(byteZGNr)
            Case Is = 17     'Plakate
                OpenProgrammTeil frmWKL72, ermittlezugriff(byteZGNr)
            Case Is = 18     'Spezialetikett
                OpenProgrammTeil frmWKL85, ermittlezugriff(byteZGNr)
            Case Is = 20     'Rabatt - Aufkleber
                OpenProgrammTeil frmWKL187, ermittlezugriff(byteZGNr)
            Case Is = 3     'zurück
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame3.Visible = False
            '***************************************Satmmdaten
            Case Is = 7     'KISS Format
                OpenProgrammTeil frmWKL11, ermittlezugriff(byteZGNr)
            Case Is = 6     'andere Formate
                OpenProgrammTeil frmWKL09, ermittlezugriff(byteZGNr)
            Case Is = 19     'andere Formate
                OpenProgrammTeil frmWKL166, ermittlezugriff(byteZGNr)
            Case Is = 5     'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame18.Visible = False
                
            Case 11 'marken
                frmWKL156.Show 1
            Case 10 'agn
                frmWKL12.Show 1
            Case 9 'pgn
                frmWKL07.Show 1
            Case 4 'lpz
                frmWKL06.Show 1
            Case 8 'zurück
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame21.Visible = False
                
            Case 13 'PS allg
                frmWKL39.Show 1
            Case 14 'PS agn
                frmWKL51.Show 1
            Case 16 'PS LINR
'                frmWKL07.Show 1
            Case 12 'PS PGN
                frmWKL121.Show 1
            Case 15 'zurück
                If gbKostenlos = True Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command13(lcount).Enabled = True
                    Next lcount
                End If
                Frame22.Visible = False
        End Select
    Else
        Select Case index
            Case Is = 0     'Strichcode-Etiketten
                If glLevel >= DlgZugriff(13).dZugriff Then
                    frmWKL30.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1     'Strichcode-Etiketten selbst wählen
                If glLevel >= DlgZugriff(14).dZugriff Then
                    frmWKL31.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 2     'Etiketten aus Lieferschein
                If glLevel >= DlgZugriff(14).dZugriff Then
                    frmWKL03.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 17     'Plakate
                If glLevel >= DlgZugriff(14).dZugriff Then
                    frmWKL72.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 18    'Spezialetikett
                If glLevel >= DlgZugriff(14).dZugriff Then
                    frmWKL85.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 20    'Rabatt - Aufkleber
                If glLevel >= DlgZugriff(14).dZugriff Then
                    frmWKL187.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 3     'zurück
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame3.Visible = False
                
            '***************************************Satmmdaten
            
            Case Is = 7     'KISS Format
                If glLevel >= DlgZugriff(1).dZugriff Then
                    frmWKL11.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 6     'andere Formate
                If glLevel >= DlgZugriff(1).dZugriff Then
                    frmWKL09.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 19     'andere Formate
                If glLevel >= DlgZugriff(1).dZugriff Then
                    frmWKL166.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 5     'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame18.Visible = False
            Case 11 'marken
                frmWKL156.Show 1
            Case 10 'agn
                frmWKL12.Show 1
            Case 9 'pgn
                frmWKL07.Show 1
            Case 4 'lpz
                frmWKL06.Show 1
            Case 8 'zurück
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame21.Visible = False
                
            Case 13 'PS allg
                frmWKL39.Show 1
            Case 14 'PS agn
                frmWKL51.Show 1
            Case 16 'PS LINR
'                frmWKL07.Show 1
            Case 12 'PS PGN
'                frmWKL06.Show 1
            Case 15 'zurück
                If gbKostenlos = True Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command13(lcount).Enabled = True
                    Next lcount
                End If
                Frame22.Visible = False
                
        
        End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctemp As String
    
    
    Screen.MousePointer = 11
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command5(index))
        Select Case index
            Case Is = 0     'Artikelliste
                    Frame9.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False
            Case Is = 1     'Kundenliste
                If glLevel >= ermittlezugriff(byteZGNr) Then
'                    schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                    
                    
                    Frame20.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False

    
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    ctemp = ctemp & "Alle Versuche Programmteile zu öffnen werden protokolliert."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            Case Is = 2     'Lieferantenliste
                OpenProgrammTeil frmWKL42, ermittlezugriff(byteZGNr)
            Case Is = 3     'Verkaufsliste
                If glLevel >= ermittlezugriff(byteZGNr) Then
'                    schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                    
                    
                    Frame23.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False

    
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."

'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            
            
            
           
            Case Is = 5     'Favoritenliste
                OpenProgrammTeil frmWKL44, ermittlezugriff(byteZGNr)
            Case Is = 6     'Bestandsliste/Inventur
                OpenProgrammTeil frmWKL46, ermittlezugriff(byteZGNr)
            Case Is = 7     'Warenzugang/Einkauf
                OpenProgrammTeil frmWKL45, ermittlezugriff(byteZGNr)
            Case Is = 8     'Warenzugang/Einkauf
                OpenProgrammTeil frmWKL48, ermittlezugriff(byteZGNr)
            Case Is = 9     'zurück
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame4.Visible = False
                Label2.Visible = True
                Label3.Visible = True
                Label1(1).Visible = True
                
                bListen = False
    
        End Select
    Else
        Select Case index
            Case Is = 0     'Artikelliste
                If glLevel >= DlgZugriff(15).dZugriff Then
                    Frame9.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1     'Kundenliste
                If glLevel >= DlgZugriff(16).dZugriff Then
                    
                    Frame20.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False

                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 2     'Lieferantenliste
                If glLevel >= DlgZugriff(17).dZugriff Then
                    frmWKL42.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 3     'Artikelgruppenliste
                If glLevel >= DlgZugriff(17).dZugriff Then
                    Frame23.Visible = True
                    Frame4.Enabled = False
                    Command5(0).Enabled = False
                    Command5(1).Enabled = False
                    Command5(2).Enabled = False
                    Command5(3).Enabled = False
                    Command5(5).Enabled = False
                    Command5(6).Enabled = False
                    Command5(7).Enabled = False
                    Command5(8).Enabled = False
                    Command5(9).Enabled = False

                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 5     'Favoritenliste
                If glLevel >= DlgZugriff(19).dZugriff Then
                    frmWKL44.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 6     'Bestandsliste/Inventur
                If glLevel >= DlgZugriff(21).dZugriff Then
                    frmWKL46.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 7     'Warenzugang/Einkauf
                If glLevel >= DlgZugriff(20).dZugriff Then
                    frmWKL45.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
        
            Case Is = 8     'Warenzugang/Einkauf
                If glLevel >= DlgZugriff(0).dZugriff Then
                    frmWKL48.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
        
            Case Is = 9     'zurück
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame4.Visible = False
                Label2.Visible = True
                Label3.Visible = True
                Label1(1).Visible = True
                
                bListen = False
        End Select
    End If
    
  
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iRet As Integer
    Dim sPfad As String
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command6(index))
    
        Select Case index
            Case Is = 0     'Einstellungen
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
                
                Frame5.Enabled = False
                Frame8.Visible = True
                
                If gbKostenlos = False Then
                    Command9(0).SetFocus
                End If
            Case Is = 1     'Datenbank...
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame5.Enabled = False
                Frame14.Visible = True
                
                If gbKostenlos = False Then
                    Command12(4).SetFocus
                End If
            Case Is = 2     'Datenaustausch...
                
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame16.Visible = True
                
                If gbKostenlos = False Then
                    Command12(11).SetFocus
                End If
            Case Is = 3     'Email schreiben
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame5.Enabled = False
                Frame17.Visible = True
                
                If gbKostenlos = False Then
                    Command12(20).SetFocus
                End If
            
            Case Is = 4     'DsFinvK Export.
            
            MsgBox ("in Arbeit . . .")
'            ExportFormular.Left = (Me.ScaleWidth - ExportFormular.Width) / 2
'            ExportFormular.Top = (Me.ScaleHeight - ExportFormular.Height) / 2
'            ExportFormular.Show 1
                
            Case Is = 5     'Zugriffsrechte             ehemals 55
                Screen.MousePointer = 11
                OpenProgrammTeil frmWKL54, ermittlezugriff(byteZGNr)
            Case Is = 6     'Programmeinstellung
                Screen.MousePointer = 11
                OpenProgrammTeil frmWKL53, ermittlezugriff(byteZGNr)
            Case Is = 7             'DTA-Diskette
                OpenProgrammTeil frmWKL57, ermittlezugriff(byteZGNr)
            Case Is = 8
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame5.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bService = False
        
        End Select
    Else
        
        Select Case index
            Case Is = 0     'Einstellungen
                If glLevel >= DlgZugriff(25).dZugriff Then
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = False
                    Next lcount
                    
                    Frame5.Enabled = False
                    Frame8.Visible = True
                    
                    If gbKostenlos = False Then
                        Command9(0).SetFocus
                    End If
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1     'Datenbank...
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame5.Enabled = False
                Frame14.Visible = True
                
                If gbKostenlos = False Then
                    Command12(4).SetFocus
                End If
            Case Is = 2     'Datenaustausch...
                
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame16.Visible = True
                
                If gbKostenlos = False Then
                    Command12(11).SetFocus
                End If
            Case Is = 3     'Email schreiben
                For lcount = 0 To 8
                    Command6(lcount).Enabled = False
                Next lcount
    
                Frame5.Enabled = False
                Frame17.Visible = True
                
                If gbKostenlos = False Then
                    Command12(20).SetFocus
                End If
            
            
            
                    
    
            Case Is = 5     'Zugriffsrechte
                If glLevel >= DlgZugriff(28).dZugriff Then
                    Screen.MousePointer = 11
                    frmWKL54.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 6     'Programmeinstellung
                If glLevel >= DlgZugriff(31).dZugriff Then 'soll mal Zugriff 30 werden
                    Screen.MousePointer = 11
                    frmWKL53.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 7             'DTA-Diskette
                If glLevel >= DlgZugriff(31).dZugriff Then
                    frmWKL57.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 8
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame5.Visible = False
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bService = False
        
        End Select
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten." & index
    
    Fehlermeldung1
End Sub
Private Sub Speicher_alle_BTN()
On Error GoTo LOKAL_ERROR

'    loeschNEW "BTNBESCHRIFTUNG", gdBase
'    CreateTableT3 "BTNBESCHRIFTUNG", gdBase
'
'
'    SpeicherBTNCaption frmWKL50
'    SpeicherBTNCaption frmWKL31
'    SpeicherBTNCaption frmWKL10
'    SpeicherBTNCaption frmWKL12
'    SpeicherBTNCaption frmWKL13
'    SpeicherBTNCaption frmWKL18
'    SpeicherBTNCaption frmWKL19
'    SpeicherBTNCaption frmWKL40
'    SpeicherBTNCaption frmWKL41
'    SpeicherBTNCaption frmWKL42
'    SpeicherBTNCaption frmWKL21
'    SpeicherBTNCaption frmWKL22
'    SpeicherBTNCaption frmWKL23
'    SpeicherBTNCaption frmWKL52
'    SpeicherBTNCaption frmWKL15
'    SpeicherBTNCaption frmWKL29
'    SpeicherBTNCaption frmWKL11
'    SpeicherBTNCaption frmWKL99

'    SpeicherBTNCaption frmWKL30
'    SpeicherBTNCaption frmWKL24
'    SpeicherBTNCaption frmWKL28
'    SpeicherBTNCaption frmWK25a
'    SpeicherBTNCaption frmWKL82
'    SpeicherBTNCaption frmWKL81
'    SpeicherBTNCaption frmWK25c
'    SpeicherBTNCaption frmWKL83
'    SpeicherBTNCaption frmWKL44
'    SpeicherBTNCaption frmWK21b
'    SpeicherBTNCaption frmWK00a
'    SpeicherBTNCaption frmWKL16
'    SpeicherBTNCaption frmWK81a
'    SpeicherBTNCaption frmWK81b
'    SpeicherBTNCaption frmWK81c
'    SpeicherBTNCaption frmWK81d
'    SpeicherBTNCaption frmWK81e
'    SpeicherBTNCaption frmWK81f
'    SpeicherBTNCaption frmWKL45
'    SpeicherBTNCaption frmWKL57

'    SpeicherBTNCaption frmWKL46
'    SpeicherBTNCaption frmWKL01
'    SpeicherBTNCaption frmWKL17
'    SpeicherBTNCaption frmWKL58
'    SpeicherBTNCaption frmWKL59
'    SpeicherBTNCaption frmWK20a
'    SpeicherBTNCaption frmWKLab
'    SpeicherBTNCaption frmWK20b
'    SpeicherBTNCaption frmWK21d
'    SpeicherBTNCaption frmWK21g
'    SpeicherBTNCaption frmWK21r
'    SpeicherBTNCaption frmWK15a
'    SpeicherBTNCaption frmWK20c
'    SpeicherBTNCaption frmWKL02
'    SpeicherBTNCaption frmWK40c
'    SpeicherBTNCaption frmWK25d
'    SpeicherBTNCaption frmWK24a

'    SpeicherBTNCaption frmWKLaf
'    SpeicherBTNCaption frmWKLai
'    SpeicherBTNCaption frmWKLaj
'    SpeicherBTNCaption frmWKLak
'    SpeicherBTNCaption frmWKLal
'    SpeicherBTNCaption frmWKL00b
'    SpeicherBTNCaption frmWK24b
'    SpeicherBTNCaption frmWK20d
'    SpeicherBTNCaption frmWK25f
'    SpeicherBTNCaption frmWK25g
'    SpeicherBTNCaption frmWK24c
'    SpeicherBTNCaption frmWKLam
'    SpeicherBTNCaption frmWKLan
'    SpeicherBTNCaption frmWK25h
'    SpeicherBTNCaption frmWK25m
'    SpeicherBTNCaption frmWK10a
'    SpeicherBTNCaption frmWK21f

'    SpeicherBTNCaption frmWK21p
'    SpeicherBTNCaption frmWK20h
'    SpeicherBTNCaption frmWK20e
'    SpeicherBTNCaption frmWK20g
'    SpeicherBTNCaption frmWK20j
'    SpeicherBTNCaption frmWKL48
'    SpeicherBTNCaption frmWK20f
'    SpeicherBTNCaption frmWKL56
'    SpeicherBTNCaption frmWKLar
'    SpeicherBTNCaption frmWKLas
'    SpeicherBTNCaption frmWKLau
'    SpeicherBTNCaption frmWKLav
'    SpeicherBTNCaption frmWK21k
'    SpeicherBTNCaption frmWK21l
'    SpeicherBTNCaption frmWKL53
'    SpeicherBTNCaption dlgAbfrage
'    SpeicherBTNCaption dlgRetoure
'    SpeicherBTNCaption dlgBestand
'    SpeicherBTNCaption dlgAuszahlung
'    SpeicherBTNCaption frmWKL25
'    SpeicherBTNCaption frmWKL26

'    SpeicherBTNCaption frmWKL27
'    SpeicherBTNCaption frmWKL32
'    SpeicherBTNCaption frmWKL33
'    SpeicherBTNCaption frmWKL34
'    SpeicherBTNCaption dlgAbfrage3
'    SpeicherBTNCaption dlgKopieren
'    SpeicherBTNCaption frmWKL35
'    SpeicherBTNCaption frmWKL36
'    SpeicherBTNCaption frmWKL03
'    SpeicherBTNCaption frmWKL38
'    SpeicherBTNCaption frmWKL37
'    SpeicherBTNCaption frmWKL09
'    SpeicherBTNCaption frmWK11a
'    SpeicherBTNCaption frmWK12a
'    SpeicherBTNCaption frmWKL54
'    SpeicherBTNCaption frmWKL47
'    SpeicherBTNCaption frmWKL49
'    SpeicherBTNCaption frmWKL06
'    SpeicherBTNCaption frmWKL07
'    SpeicherBTNCaption frmWKL14
'    SpeicherBTNCaption frmWKL05

'    SpeicherBTNCaption frmWKL39
'    SpeicherBTNCaption frmWKL51
'    SpeicherBTNCaption frmWKL55
'    SpeicherBTNCaption frmWKL60
'    SpeicherBTNCaption frmWK21m
'    SpeicherBTNCaption frmWKL61
'    SpeicherBTNCaption frmWK21s
'    SpeicherBTNCaption frmWK21t
'    SpeicherBTNCaption frmWK21n
'    SpeicherBTNCaption frmWK21o
'    SpeicherBTNCaption frmWKL62
'    SpeicherBTNCaption frmWKL63
'    SpeicherBTNCaption frmWKL64
'    SpeicherBTNCaption frmWKL65
'    SpeicherBTNCaption frmWKL66
'    SpeicherBTNCaption frmWKL67
'    SpeicherBTNCaption frmWKL68
'    SpeicherBTNCaption frmWKL69
'    SpeicherBTNCaption frmWKL70
'    SpeicherBTNCaption frmWK25i
'    SpeicherBTNCaption frmWK25j
'    SpeicherBTNCaption frmWK25k
'    SpeicherBTNCaption frmWKL71
'    SpeicherBTNCaption frmWKL72
'    SpeicherBTNCaption frmWKL73

'    SpeicherBTNCaption frmWKL75
'    SpeicherBTNCaption frmWKL76
'    SpeicherBTNCaption frmWKL77
'    SpeicherBTNCaption frmWKL78
'    SpeicherBTNCaption frmWKL79
'    SpeicherBTNCaption frmWKL80
'    SpeicherBTNCaption frmWKL84
'    SpeicherBTNCaption frmWKL85
'    SpeicherBTNCaption frmWKL86
'    SpeicherBTNCaption frmWKL87
'    SpeicherBTNCaption frmWKL88
'    SpeicherBTNCaption frmWKL89
'    SpeicherBTNCaption frmWKL91
'    SpeicherBTNCaption frmWKL92
'    SpeicherBTNCaption frmWKL93
'    SpeicherBTNCaption frmWKL100
'    SpeicherBTNCaption frmWKL110
'    SpeicherBTNCaption frmWKL101
'    SpeicherBTNCaption frmWKL102
'    SpeicherBTNCaption frmWKL103
'    SpeicherBTNCaption frmWKL111
'    SpeicherBTNCaption frmWKL112
'    SpeicherBTNCaption frmWKL113
'    SpeicherBTNCaption frmWKL114
'    SpeicherBTNCaption frmWKL115
'    SpeicherBTNCaption frmWKL116
'    SpeicherBTNCaption frmWKL117
'    SpeicherBTNCaption frmWKLah
'    SpeicherBTNCaption frmWKL118
'    SpeicherBTNCaption frmWKL119
'    SpeicherBTNCaption frmWKL120
'    SpeicherBTNCaption frmWKL121
'    SpeicherBTNCaption frmWKL122

'    SpeicherBTNCaption frmWKL123
'    SpeicherBTNCaption frmWKL124
'    SpeicherBTNCaption frmWKL125
'    SpeicherBTNCaption frmWKL126
'    SpeicherBTNCaption frmWKL127
'    SpeicherBTNCaption frmWKL128
'    SpeicherBTNCaption frmWKL129
'    SpeicherBTNCaption frmWKL130
'    SpeicherBTNCaption frmWKL131
'    SpeicherBTNCaption frmWKL132
'    SpeicherBTNCaption frmWKL133
'    SpeicherBTNCaption frmWKL134
'    SpeicherBTNCaption frmWKL135
'    SpeicherBTNCaption frmWKL136
'    SpeicherBTNCaption frmWKL137
'    SpeicherBTNCaption frmWKL138
'    SpeicherBTNCaption frmWKL139
'    SpeicherBTNCaption frmWKL140
'    SpeicherBTNCaption frmWKL141
'    SpeicherBTNCaption frmWKL142
'    SpeicherBTNCaption frmWKL143
'    SpeicherBTNCaption frmWKL144
'    SpeicherBTNCaption frmWKL145
'    SpeicherBTNCaption frmWKL146
'    SpeicherBTNCaption frmWKL147
'    SpeicherBTNCaption frmWKL148
'    SpeicherBTNCaption frmWKL149
'    SpeicherBTNCaption frmWKL150
'    SpeicherBTNCaption frmWKL151
'    SpeicherBTNCaption frmWKL152
'    SpeicherBTNCaption frmWKL153
'    SpeicherBTNCaption frmWKL154
'    SpeicherBTNCaption frmWKL155
'    SpeicherBTNCaption frmWKL156


'    SpeicherBTNCaption frmWKL161
'    SpeicherBTNCaption frmWKL162
'    SpeicherBTNCaption frmWKL163
'    SpeicherBTNCaption frmWKL164
'    SpeicherBTNCaption frmWKL165
'    SpeicherBTNCaption frmWKL166
'    SpeicherBTNCaption frmWKL167
'    SpeicherBTNCaption frmWKLao
'    SpeicherBTNCaption frmWK16a
'    SpeicherBTNCaption frmWKL168
'    SpeicherBTNCaption frmWKL169
'    SpeicherBTNCaption frmWKL170
'    SpeicherBTNCaption frmWKL171
'    SpeicherBTNCaption frmWK25l
'    SpeicherBTNCaption frmWKL172
'    SpeicherBTNCaption frmWKL173
'    SpeicherBTNCaption frmWKL08
'    SpeicherBTNCaption frmWKL174
'    SpeicherBTNCaption frmWKL175
'    SpeicherBTNCaption frmWKL176
'    SpeicherBTNCaption frmWKL177
'    SpeicherBTNCaption frmWK20i
'    SpeicherBTNCaption frmWKL179
'    SpeicherBTNCaption frmWKL180
'    SpeicherBTNCaption frmWKL181
'    SpeicherBTNCaption frmWKL182
'    SpeicherBTNCaption frmWKL183
'    SpeicherBTNCaption frmWKL184
'    SpeicherBTNCaption frmWKL185
'    SpeicherBTNCaption frmWKL186
'    SpeicherBTNCaption frmWKL187
'    SpeicherBTNCaption frmWKL188
'    SpeicherBTNCaption frmWKL192
'    SpeicherBTNCaption frmWKL193
'    SpeicherBTNCaption frmWKL94
'    SpeicherBTNCaption frmWKL194
'    SpeicherBTNCaption frmWKL195
'    SpeicherBTNCaption frmWKL196
'    SpeicherBTNCaption frmWKL197
'    SpeicherBTNCaption frmWKL90

'    SpeicherBTNCaption frmWKL198
'    SpeicherBTNCaption frmWKL199
'    SpeicherBTNCaption frmWKL200
'    SpeicherBTNCaption frmWKL201
'    SpeicherBTNCaption frmWKL202
'    SpeicherBTNCaption frmWKL203
'    SpeicherBTNCaption frmWKL204
'    SpeicherBTNCaption frmWKL205
'    SpeicherBTNCaption frmWKL206
'    SpeicherBTNCaption frmWKL209
'    SpeicherBTNCaption frmWKL210
'    SpeicherBTNCaption frmWKL211
'    SpeicherBTNCaption frmWKL212
'    SpeicherBTNCaption frmWKL213
'    SpeicherBTNCaption frmWKL214
'    SpeicherBTNCaption frmWKL215
'    SpeicherBTNCaption frmWKL216
'    SpeicherBTNCaption frmWKL160
'    SpeicherBTNCaption frmWKL217
'    SpeicherBTNCaption frmWKL218
'    SpeicherBTNCaption dlgPW
'    SpeicherBTNCaption dlgTaNr
'    SpeicherBTNCaption dlgGeschenkset
'    SpeicherBTNCaption dlgGutschein
'    SpeicherBTNCaption dlgTermDel

'    SpeicherBTNCaption frmWKL20
'    SpeicherBTNCaption frmWKL43


'    gcArtNrFiliale = "332284"
'    SpeicherBTNCaption frmWKLae
'
'    gckundnr = "1"
'    SpeicherBTNCaption frmWKL74
'
'    SpeicherBTNCaption frmWKL157
'
'    SpeicherBTNCaption frmWKL158
'SpeicherBTNCaption frmWKL159


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Speicher_alle_BTN"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command7(index))
    
        Select Case index
            Case Is = 0     'Protokoll der Bestandsveränderungen
                OpenProgrammTeil frmWKL37, ermittlezugriff(byteZGNr)
            Case Is = 1     'Artikelliste aus MDE/Scanner
                OpenProgrammTeil frmWKL71, ermittlezugriff(byteZGNr)
            Case Is = 3     'allg. Verkaufsvorgänge
                OpenProgrammTeil frmWKL136, ermittlezugriff(byteZGNr)
            Case Is = 4     'Verkaufsprotokoll
                OpenProgrammTeil frmWK25d, ermittlezugriff(byteZGNr)
            Case Is = 5     'Rabattverkäufe
                OpenProgrammTeil frmWK25f, ermittlezugriff(byteZGNr)
            Case Is = 2     'Kassenprotokolle
                OpenProgrammTeil frmWKL73, ermittlezugriff(byteZGNr)
            Case Is = 6     'Arbeitszeit
            
                If gbAA = True Then
                    OpenProgrammTeil frmWK25g, ermittlezugriff(byteZGNr)
                Else
                    OpenProgrammTeil frmWKL14, ermittlezugriff(byteZGNr)
                End If
            Case Is = 7
                Frame2.Enabled = True
                Frame6.Visible = False
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 10
                        Command3(lcount).Enabled = True
                    Next lcount
                End If
                
                
                If gbLokalModus Then
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                End If
                
'                If gbKostenlos Then
'                    SetzefürKostenlos frmWKL00
'                Else
'                    Command3(0).Enabled = True
'                    Command3(1).Enabled = True
'                    Command3(2).Enabled = True
'                    Command3(6).Enabled = True
'                    Command3(3).Enabled = False
'                End If
                
            Case Is = 8     'GDPdU/DATEV
                OpenProgrammTeil frmWKL171, ermittlezugriff(byteZGNr)
        End Select
    Else
        Select Case index
            Case Is = 0     'protokoll der bestandsveränderungen
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL37.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 1     'artikelliste aus mde
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL71.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 3     'allg. Verkaufsvorgänge
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL136.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 4     'Verkaufsprotokoll
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWK25d.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 5     'Rabattverkäufe
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWK25f.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 6     'Arbeitszeit
                If glLevel >= 9 Then
                    If gbBEDKARTE = False Then
                        frmWK25g.Show 1
                    Else
                        frmWKL14.Show 1
                    End If
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 2     'Kassenprotokolle
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL73.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 7
                Frame2.Enabled = True
                Frame6.Visible = False
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 10
                        Command3(lcount).Enabled = True
                    Next lcount
                End If
                
                
                
                If gbLokalModus Then
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                Else
'                    Command3(4).Enabled = True
'                    Command3(5).Enabled = True
'                    Command3(7).Enabled = True
'                    Command3(8).Enabled = True
'                    Command3(9).Enabled = True
'                    Command3(10).Enabled = True
                End If
                
''                If gbKostenlos Then
''                    SetzefürKostenlos frmWKL00
''                Else
''                    Command3(0).Enabled = True
''                    Command3(1).Enabled = True
''                    Command3(2).Enabled = True
''                    Command3(6).Enabled = True
''                    Command3(3).Enabled = False
''                End If
            Case Is = 8    'GDPdU/DATEV
                If glLevel >= DlgZugriff(11).dZugriff Then

                    frmWKL171.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Command8_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command8(index))
        Select Case index
            Case Is = 0     'Vorgaben
                OpenProgrammTeil frmWKL81, ermittlezugriff(byteZGNr)
            Case Is = 1     'Termin-Kalender
                OpenProgrammTeil frmWKL82, ermittlezugriff(byteZGNr)
            Case Is = 2     'Notizen
                OpenProgrammTeil frmWKL83, ermittlezugriff(byteZGNr)
            Case Is = 4     'Reparaturverwaltung
                OpenProgrammTeil frmWKL133, ermittlezugriff(byteZGNr)
            Case Is = 3     '** Zurück **
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(6).Enabled = True
                Command1(7).Enabled = True
                Command1(8).Enabled = True
                
                Frame7.Visible = False
                
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bTermine = False
    
        End Select
    Else
        Select Case index
            Case Is = 0     'Vorgaben
                If glLevel >= DlgZugriff(22).dZugriff Then
                    frmWKL81.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
                
            Case Is = 1     'Termin-Kalender
                If glLevel >= DlgZugriff(23).dZugriff Then
                    frmWKL82.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 2     '** Notizen **
    '            MsgBox "z.Zt. noch nicht möglich!"
    '            Exit Sub
                
                If glLevel >= DlgZugriff(24).dZugriff Then
                    frmWKL83.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 4     '** Reparaturverwaltung **
                If glLevel >= DlgZugriff(24).dZugriff Then
                    frmWKL133.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 3     '** Zurück **
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(6).Enabled = True
                Command1(7).Enabled = True
                Command1(8).Enabled = True
                
                Frame7.Visible = False
                
                Label2.Visible = True
                Label3.Caption = ""
                Label1(1).Visible = True
                
                bTermine = False
    
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim lcount As Long
    Dim ireslt As Integer
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command9(index))
    
        Select Case index
            Case Is = 0     'Unternehmensangaben
                OpenProgrammTeil frmWKL16, ermittlezugriff(byteZGNr)
                LeseFirmenDaten
            Case Is = 1     'Druckereintellungen
                OpenProgrammTeil frmWKL50, ermittlezugriff(byteZGNr)
            Case Is = 2 'Texte Kassenbon
                OpenProgrammTeil frmWKL52, ermittlezugriff(byteZGNr)

            Case Is = 3 'Kartenleser
                OpenProgrammTeil frmWKL58, ermittlezugriff(byteZGNr)
            Case Is = 4 'Warengruppen
                OpenProgrammTeil frmWKL59, ermittlezugriff(byteZGNr)
            Case Is = 5
            
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
                
                Frame5.Enabled = True
                Frame8.Visible = False
                
            Case Is = 6 'Bedienerverwaltung
                OpenProgrammTeil frmWKL22, ermittlezugriff(byteZGNr)
            Case Is = 7 'Bonus auf Bon
                OpenProgrammTeil frmWKL34, ermittlezugriff(byteZGNr)
            Case Is = 8 'MWST
                
                ireslt = MsgBox("Winkiss muss an allen Arbeitplätzen beendet werden, bevor" & vbNewLine & "Sie die MWST ändern können." & vbNewLine & vbNewLine & "gegenwärtige MWST: ( V: " & gdMWStV & ", E: " & gdMWStE & ", O: " & gdMWStO & " )", vbQuestion + vbYesNo, "Winkiss Frage:")
                If ireslt = vbYes Then
                 TabelleMWSTSATZ_Erweiterungen_wiederherstellen
                 OpenProgrammTeil frmWKL56, ermittlezugriff(byteZGNr)
                End If
                
        End Select
    
    Else
    
        Select Case index
            Case Is = 0     'Unternehmensangaben
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL16.Show 1
                    LeseFirmenDaten
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL50.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 2 'Texte Kassenbon
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL52.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 3
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL58.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 4 'Warengruppen
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL59.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 5
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
                Frame5.Enabled = True
                Frame8.Visible = False
                
            Case Is = 6 'Bediener
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL22.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 7 'Bonus auf Bon
                If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL34.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 8
            
                
                ireslt = MsgBox("Winkiss muss an allen Arbeitplätzen beendet werden, bevor" & vbNewLine & "Sie die MWST ändern können." & vbNewLine & vbNewLine & "gegenwärtige MWST: ( V: " & gdMWStV & ", E: " & gdMWStE & ", O: " & gdMWStO & " )", vbQuestion + vbYesNo, "Winkiss Frage:")
                If ireslt = vbYes Then
                
                  TabelleMWSTSATZ_Erweiterungen_wiederherstellen
                  
                  If glLevel >= DlgZugriff(25).dZugriff Then
                    frmWKL56.Show 1
                  Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                  End If
                  
                End If
                 
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctemp   As String
    
'    Screen.MousePointer = 11
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command10(index))
        Select Case index
            Case Is = 0     'Artikelliste nach Lieferanten
                OpenProgrammTeil frmWKL40, ermittlezugriff(byteZGNr)
            Case Is = 1     'Artikelliste nach AGN
                OpenProgrammTeil frmWKL41, ermittlezugriff(byteZGNr)
            Case Is = 2     'negative Bestände
                If glLevel >= ermittlezugriff(byteZGNr) Then
'                    schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                    
                    negart
                
                    Screen.MousePointer = 0
                    Exit Sub
    
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    ctemp = ctemp & "Alle Versuche Programmteile zu öffnen werden protokolliert."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            Case Is = 3 'reduzierte Preise
                If glLevel >= ermittlezugriff(byteZGNr) Then
'                    schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                    
                    MsgBox "Leider außer Betrieb", vbInformation, "Winkiss Hinweis:"
'
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    ctemp = ctemp & "Alle Versuche Programmteile zu öffnen werden protokolliert."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
                
            Case Is = 4 'Gutscheinliste
            
                If gbKL_LIVEGUTSCHEIN Then
                    MsgBox "Leider außer Betrieb", vbInformation, "Winkiss Hinweis:"
                Else
                    OpenProgrammTeil frmWKLal, ermittlezugriff(byteZGNr)
                End If
            
            
                
            Case Is = 5
                Frame9.Visible = False
                Frame4.Enabled = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If
                
            Case Is = 6  'unterschrittene Mindestmenge
                OpenProgrammTeil frmWKLas, ermittlezugriff(byteZGNr)
            Case Is = 7 'Mindestbestand ermitteln
                OpenProgrammTeil frmWKL139, ermittlezugriff(byteZGNr)
            Case Is = 8 'diverse Listen
                OpenProgrammTeil frmWKL55, ermittlezugriff(byteZGNr)
        End Select
    Else
    
        Select Case index
            Case Is = 0
                If glLevel >= DlgZugriff(15).dZugriff Then
                    frmWKL40.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1
                If glLevel >= DlgZugriff(16).dZugriff Then
                    frmWKL41.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 2 'negative Bestände
                If glLevel >= DlgZugriff(15).dZugriff Then
                
                    negart
                    
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 3 'reduzierte Preise
                If glLevel >= DlgZugriff(15).dZugriff Then
                
'                    redart
                    MsgBox "Leider außer Betrieb", vbInformation, "Winkiss Hinweis:"
                    
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 4 'Gutscheinliste
                If glLevel >= DlgZugriff(15).dZugriff Then
                    If gbKL_LIVEGUTSCHEIN Then
                        MsgBox "Leider außer Betrieb", vbInformation, "Winkiss Hinweis:"
                    Else
                        frmWKLal.Show 1
                    End If
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 5
                
                Frame9.Visible = False
                Frame4.Enabled = True
                
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If
                
            Case Is = 6  'unterschrittene Mindestmenge
                If glLevel >= DlgZugriff(15).dZugriff Then
                    frmWKLas.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 7 'Mindestbestand ermitteln
                If glLevel >= DlgZugriff(15).dZugriff Then
                    frmWKL139.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 8 'diverse Listen
                If glLevel >= DlgZugriff(15).dZugriff Then
                    frmWKL55.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
        End Select
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub negart()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "NEGART", gdBase
    CreateTable "NEGART", gdBase
    cSQL = "Insert into NEGART Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " , EAN "
    
    cSQL = cSQL & " from ARTIKEL where BESTAND < 0 "
    cSQL = cSQL & " and not Bestand is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on NEGART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update NEGART inner join LISRT on NEGART.Linr = LISRT.Linr "
    cSQL = cSQL & " set NEGART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    reportbildschirm "WKL024", "aWKL40aa"
    
    Pause (3)
    loeschNEW "NEGART", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "negart"
    Fehler.gsFehlertext = "Beim Ermitteln der negativen Bestände ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command11_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command11(index))
        Select Case index
            Case Is = 0     'Kreditverwaltung
                OpenProgrammTeil frmWKL24, ermittlezugriff(byteZGNr)
            Case Is = 1     'automatische Rechnungserstellung
                OpenProgrammTeil frmWK24c, ermittlezugriff(byteZGNr)
            Case Is = 2
            
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 10
                        Command3(lcount).Enabled = True
                    Next lcount
                End If
                
                If gbLokalModus Then
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                End If
                
                Frame2.Enabled = True
                Frame10.Visible = False
        End Select
    Else
    
        Select Case index
            Case Is = 0
                If glLevel >= DlgZugriff(10).dZugriff Then
                    frmWKL24.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 1
                If glLevel >= DlgZugriff(10).dZugriff Then
                    frmWK24c.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 2
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 10
                        Command3(lcount).Enabled = True
                    Next lcount
                End If
                
                If gbLokalModus Then
                    Command3(4).Enabled = False
                    Command3(5).Enabled = False
                    Command3(7).Enabled = False
                    Command3(8).Enabled = False
                    Command3(9).Enabled = False
                    Command3(10).Enabled = False
                End If
                
                
                Frame2.Enabled = True
                Frame10.Visible = False
                
        End Select
    End If
    
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BestKundenBonus()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim lBestand As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART61", gdBase
    CreateTable "ART61", gdBase
    
    txtStatus.Text = 12
    '1.Schritt alle Artikel auswählen
    
    
    
    

    sSQL = " Insert into ART61 select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , KUNDNR "
    sSQL = sSQL & " , PREIS as VKPREIS "
    sSQL = sSQL & " , MENGE as VKMENGE "
    sSQL = sSQL & " , 0.00 as ERTRAG "
    sSQL = sSQL & " , adate "
    sSQL = sSQL & " , FILIALE "
    sSQL = sSQL & " , EKPR as LEKPR  "
    sSQL = sSQL & " , VKPR as KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , MWST "
    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", '' as Nachname "
    sSQL = sSQL & ", '' as Vorname "
    sSQL = sSQL & ", 0 as Rabatt "
    sSQL = sSQL & ", 0 as Bonus "
    
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where year(adate) = year(now)"
    sSQL = sSQL & " and kundnr <> 0 "
    sSQL = sSQL & " and not kundnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
    
    sSQL = "Update ART61 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStV & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 16
    
    sSQL = "Update ART61 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStE & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ART61 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStO & " )) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError

    
    txtStatus.Text = 18
    
    
    sSQL = " Create index kundnr on ART61(kundnr)  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21
    
    sSQL = " Create index LINR on ART61(LINR)  "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 25
    
    sSQL = " Update ART61 inner join kunden on ART61.kundnr = kunden.kundnr "
    sSQL = sSQL & " set ART61.Nachname = kunden.name "
    sSQL = sSQL & ", ART61.Vorname =  kunden.Vorname "
    sSQL = sSQL & ", ART61.Rabatt =  kunden.Rabatt "
    sSQL = sSQL & ", ART61.Bonus =  kunden.Bonus "
    gdBase.Execute sSQL, dbFailOnError
    

    
    txtStatus.Text = 31
    
    sSQL = " Update ART61 inner join LISRT on ART61.Linr = LISRT.Linr "
    sSQL = sSQL & " set ART61.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART61d", gdBase
    
    txtStatus.Text = 39
    
    sSQL = " select sum(vkpreis) as spreis,kundnr into ART61d  "
    sSQL = sSQL & " from ART61 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 48
    
    sSQL = " Update ART61 inner join ART61d on ART61.kundnr = ART61d.kundnr "
    sSQL = sSQL & " set ART61.SPREIS = ART61d.SPREIS "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "ART61d", gdBase
    
    txtStatus.Text = 52
    
    sSQL = " select sum(vkmenge) as sMenge,kundnr into ART61d  "
    
    sSQL = sSQL & " from ART61 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 55
    
    sSQL = " Update ART61 inner join ART61d on ART61.kundnr = ART61d.kundnr "
    sSQL = sSQL & " set ART61.SMenge = ART61d.SMenge "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ART61d", gdBase
    txtStatus.Text = 57
    
    sSQL = " select sum(ertrag) as sertrag,kundnr into ART61d  "
    
    sSQL = sSQL & " from ART61 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 59
    
    sSQL = " Update ART61 inner join ART61d on ART61.kundnr = ART61d.kundnr "
    sSQL = sSQL & " set ART61.Sertrag = ART61d.Sertrag "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 67
    

    
    loeschNEW "ART61c", gdBase
    
    sSQL = " select *  into ART61c from ART61  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 77
    
    loeschNEW "ART61", gdBase
    
    sSQL = " select *  into ART61 from ART61c order by Bonus desc "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ART61c", gdBase
    
    sSQL = " Update ART61  set SORTI = 'Bonus' "
    gdBase.Execute sSQL, dbFailOnError
        
    txtStatus.Text = 98
    
    reportbildschirm "", "aZEN00y"
    
    Pause 2
    loeschNEW "ART61", gdBase
    loeschNEW "ART61c", gdBase
    
    Screen.MousePointer = 0
    txtStatus.Text = 0
    picprogress.Visible = False
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestKundenBonus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub

Private Sub BestKunden(soerder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim lBestand As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    
    lbl6(28).Visible = True
    lbl6(28).Caption = "Die besten Kunden werden ermittelt..."
    
    lbl6(53).Visible = True
    lbl6(53).Caption = "bitte warten..."
    
    Me.Refresh
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 10
    
    loeschNEW "ART59", gdBase
    CreateTable "ART59", gdBase
    
    loeschNEW "ART61a", gdBase
    CreateTable "ART61a", gdBase
    
    txtStatus.Text = 12
    '1.Schritt alle Artikel auswählen
    
    sSQL = " Insert into ART59 select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , KUNDNR "
    sSQL = sSQL & " , PREIS as VKPREIS "
    sSQL = sSQL & " , MENGE as VKMENGE "
    sSQL = sSQL & " , 0.00 as ERTRAG "
    sSQL = sSQL & " , adate "
    sSQL = sSQL & " , FILIALE "
    sSQL = sSQL & " , EKPR as LEKPR  "
    sSQL = sSQL & " , VKPR as KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , MWST "
    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", '' as Nachname "
    sSQL = sSQL & ", '' as Vorname "
    sSQL = sSQL & ", 0 as Rabatt "
    
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where year(adate) = year(now)"
    sSQL = sSQL & " and kundnr <> 0 "
    sSQL = sSQL & " and not kundnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into ART61a select  distinct(kundnr) as KNUMMER"
    sSQL = sSQL & " from art59 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
    sSQL = "Update ART59 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStV & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 16
    
    sSQL = "Update ART59 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStE & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 17
    
    sSQL = "Update ART59 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStO & " )) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 18
    
    sSQL = " Create index kundnr on art59(kundnr)  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21
    
    sSQL = " Create index knummer on art61a(knummer)  "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 25
    
    sSQL = " Update ART61a inner join kunden on ART61a.knummer = kunden.kundnr "
    sSQL = sSQL & " set ART61a.Nachname = kunden.name "
    sSQL = sSQL & ", ART61a.Vorname =  kunden.Vorname "
    sSQL = sSQL & ", ART61a.Rabatt =  kunden.Rabatt "
    gdBase.Execute sSQL, dbFailOnError
    

    
    txtStatus.Text = 31

    
    loeschNEW "ART59d", gdBase
    
    txtStatus.Text = 39
    
    sSQL = " select sum(vkpreis) as spreis,kundnr into ART59d  "
    sSQL = sSQL & " from ART59 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 48
    
    sSQL = " Update ART61a inner join ART59d on ART61a.knummer = ART59d.kundnr "
    sSQL = sSQL & " set ART61a.SPREIS = ART59d.SPREIS "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "ART59d", gdBase
    
    txtStatus.Text = 52
    
    sSQL = " select sum(vkmenge) as sMenge,kundnr into ART59d  "
    
    sSQL = sSQL & " from ART59 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 55
    
    sSQL = " Update ART61a inner join ART59d on ART61a.knummer = ART59d.kundnr "
    sSQL = sSQL & " set ART61a.SMenge = ART59d.SMenge "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ART59d", gdBase
    txtStatus.Text = 57
    
    sSQL = " select sum(ertrag) as sertrag,kundnr into ART59d  "
    sSQL = sSQL & " from ART59 "
    sSQL = sSQL & " group by kundnr"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 59
    
    sSQL = " Update ART61a inner join ART59d on ART61a.knummer = ART59d.kundnr "
    sSQL = sSQL & " set ART61a.Sertrag = ART59d.Sertrag "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 67
    

    
    loeschNEW "ART59c", gdBase

    sSQL = " select *  into ART59c from art61a  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 77
    
    loeschNEW "ART61a", gdBase

    sSQL = " select *  into ART61a from art59c order by " & soerder & " desc "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78
    
    loeschNEW "ART59c", gdBase
    
    Select Case UCase(soerder)
    
        Case "SPREIS"
            sSQL = " Update ART61a  set SORTI = 'Umsatz' "
            gdBase.Execute sSQL, dbFailOnError
        
        Case "SMENGE"
            sSQL = " Update ART61a  set SORTI = 'Stückzahl' "
            gdBase.Execute sSQL, dbFailOnError
        
        Case "SERTRAG"
            sSQL = " Update ART61a  set SORTI = 'Ertrag' "
            gdBase.Execute sSQL, dbFailOnError
        
    End Select
    
    
    txtStatus.Text = 98
    
    lbl6(28).Caption = "Die besten Kunden werden ermittelt..."
    lbl6(53).Caption = "Die Druckvorschau wird erstellt..."
    
    Me.Refresh
    reportbildschirm "", "aZEN00n"
    
    Pause 2
    
    Screen.MousePointer = 0
    txtStatus.Text = 0
    picprogress.Visible = False
    
'    Me.Refresh
    
    lbl6(28).Visible = False
    lbl6(53).Visible = False
    
    loeschNEW "ART61a", gdBase
    loeschNEW "ART59", gdBase
    loeschNEW "ART59c", gdBase
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestKunden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub
Private Sub Command12_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iRet As Integer
    Dim ctmp As String
    Dim ctemp As String
    Dim cPfad As String
    Dim sSQL As String
    Dim result&, Buff$
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command12(index))
        Select Case index
            Case 39 'Verkäufe nach Zugang
                frmWK25m.Show 1
            Case 35 'Artikelgruppenliste
                frmWK25h.Show 1
            Case 33 'Markenliste
                frmWK25i.Show 1
            Case 31 'Verkaufsanteile mit und ohne Kundenbindung
                If NewTableSuchenDBKombi("VKKUUMS", gdBase) Then
                    If Datendrin("VKKUUMS", gdBase) Then
                        iRet = MsgBox("Möchten Sie die Auswertung nocheinmal neu erstellen", vbYesNo + vbQuestion, "Winkiss Frage:")
                    Else
                        iRet = vbYes
                    End If
                Else
                    iRet = vbYes
                End If
    
                If iRet = vbYes Then
                    Verkaufsanteil
                End If
                
                Label2.Visible = True
                anzeige "normal", "Druckvorschau wird erstellt...", Label2
                Screen.MousePointer = 11
                reportbildschirm "", "aZEN00s"
                anzeige "normal", "", Label2
                Screen.MousePointer = 0
                Label2.Visible = False
            Case 37 'Produktgruppenliste
                frmWK25j.Show 1
            Case 38 'Größenauswertung
                frmWK25l.Show 1
            Case 34 'Linieliste
                frmWK25k.Show 1
            
            Case 26 'allg.Kundenliste
                loeschNEW "AKUNDLI", gdBase
                CreateTable "AKUNDLI", gdBase
                
                sSQL = "Insert into AKUNDLI select "
            
                sSQL = sSQL & " TEL "
                sSQL = sSQL & ", VORNAME "
                sSQL = sSQL & ", KUNDNR "
                sSQL = sSQL & ", NAME "
                sSQL = sSQL & ", STRASSE "
                sSQL = sSQL & ", PLZ "
                sSQL = sSQL & ", stadt as ort "
                sSQL = sSQL & ", TITEL "
                sSQL = sSQL & ", FIRMA "
                sSQL = sSQL & ", RABATT "
                sSQL = sSQL & ", DATUM1 "
                sSQL = sSQL & " from Kunden  "
                gdBase.Execute sSQL, dbFailOnError
                
                reportbildschirm "dWKL12a", "aWKL47"
                
                Pause 4
                loeschNEW "AKUNDLI", gdBase
                Exit Sub
            Case 24 'Rabattliste
                loeschNEW "AKUNDLI", gdBase
                CreateTable "AKUNDLI", gdBase
                
                sSQL = "Insert into AKUNDLI select "
            
                sSQL = sSQL & " TEL "
                sSQL = sSQL & ", VORNAME "
                sSQL = sSQL & ", KUNDNR "
                sSQL = sSQL & ", NAME "
                sSQL = sSQL & ", STRASSE "
                sSQL = sSQL & ", PLZ "
                sSQL = sSQL & ", stadt as ort "
                sSQL = sSQL & ", TITEL "
                sSQL = sSQL & ", FIRMA "
                sSQL = sSQL & ", RABATT "
                sSQL = sSQL & ", DATUM1 "
                sSQL = sSQL & " from Kunden  "
                gdBase.Execute sSQL, dbFailOnError
                reportbildschirm "dWKL12a", "aWKL47a"
                
                Pause 4
                loeschNEW "AKUNDLI", gdBase
                Exit Sub
            Case 27 'bonusliste
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKundenBonus
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 28 'KE
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "Sertrag"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 29 'KU
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "Spreis"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 30 'KS
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "smenge"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 40 'Kunden Feedback
                frmWKL172.Show 1
            Case 36
                If gbKostenlos = True Then
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(3).Enabled = True
                    Command5(9).Enabled = True
                Else
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If

                Frame23.Visible = False
                Frame4.Enabled = True
            Case 25
                If gbKostenlos = True Then
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(3).Enabled = True
                    Command5(9).Enabled = True
                Else
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If

                Frame20.Visible = False
                Frame4.Enabled = True
        
            Case 21 'Bestellung berechnen
'                MsgBox "Hier 2"
                frmWKL43.Show 1
'                MsgBox "Hier 4"
            Case 23 'Bestellung manuell
                frmWKL47.Show 1
            Case Is = 0 ' aus Einzellieferung
                OpenProgrammTeil frmWKL15, ermittlezugriff(byteZGNr)
            Case Is = 1 'aus Filialumverteilung
                OpenProgrammTeil frmWKL23, ermittlezugriff(byteZGNr)
            Case Is = 2 ' Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame11.Visible = False
            Case Is = 3 'aus Bestellung
                OpenProgrammTeil frmWK15a, ermittlezugriff(byteZGNr)
'                OpenProgrammTeil frmWK16a, ermittlezugriff(byteZGNr)
            '*****************************************************
            Case Is = 4 'Pfad zur Datenbank
                If glLevel >= ermittlezugriff(byteZGNr) Then
                    Setzedatenbankpfad
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            Case Is = 13    'Datenbankbefehl
            
            
                dlgPW.Show 1
                    
                If dlgPW.Back = True Then
                    OpenProgrammTeil frmWKLaf, ermittlezugriff(byteZGNr)
                Else
                        
                End If
            
            
               
                
            Case Is = 6 'Reorganisation
'                If glLevel >= ermittlezugriff(byteZGNr) Then

                    iRet = MsgBox("Möchten Sie jetzt die Datenbank optimieren?", vbQuestion + vbYesNo, "Winkiss Frage:")
                    If iRet = vbYes Then
                        Command12_Click 5
                        Command6_Click 8
                        
                        picprogress.Visible = True
                        lbl6(28).Visible = True
                        lbl6(53).Visible = True
            
                        If BistDualleineinderDatenbank Then
                            CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige
                           
                            sSQL = "update dbeinste set lastkomp='" & Date & "'"
                            sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
                            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                        Else
                            anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Kissdata", lbl6(28)
                        End If
                       
                        If BistDualleineinderDatenbankApp Then
                            dbApp_Compri "Kissapp.MDB"
                        Else
                            anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Kissapp", lbl6(28)
                        End If
                       
                        picprogress.Visible = False
                        lbl6(28).Visible = False
                        lbl6(53).Visible = False
                        
                        Screen.MousePointer = 0
                
                        Label2.ForeColor = glS1
                        Label2.Caption = "Anwender aktiv"
                        Label2.Refresh
                    End If

'                Else
'                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
'                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
'                End If
    
            Case Is = 7 'Fremddaten importieren
                OpenProgrammTeil frmWKL153, ermittlezugriff(byteZGNr)
            
            Case Is = 8 'DB wiederherstellen
                OpenProgrammTeil frmWKL32, ermittlezugriff(byteZGNr)
            Case Is = 12 'DB bereinigen
            
            
                dlgPW.Show 1
                    
                If dlgPW.Back = True Then
                    OpenProgrammTeil frmWKL33, ermittlezugriff(byteZGNr)
                Else
                        
                End If
            
            
            
                
            Case Is = 5 'Schließen Datenbank
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
                Frame5.Enabled = True
                Frame14.Visible = False
                
            Case Is = 19 'Schließen Kissnet
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
                Frame5.Enabled = True
                Frame17.Visible = False
                
            '****Ende Datenbankframe******************************
                
            Case Is = 10 'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
                Frame5.Enabled = True
                Frame16.Visible = False
            Case Is = 11    'Kassendaten Bereitstellung
                OpenProgrammTeil frmWKL26, ermittlezugriff(byteZGNr)
            Case Is = 9     'Kassendaten Einlesen
                OpenProgrammTeil frmWKL27, ermittlezugriff(byteZGNr)
            Case Is = 20 'KISNET... Mailbox
                OpenProgrammTeil frmWKL160, ermittlezugriff(byteZGNr)
            Case Is = 14 'Email schreiben
                If glLevel >= ermittlezugriff(byteZGNr) Then
'                    schreibeBEDProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                    
                    Buff = "mailto:"
                    result = ShellExecute(0&, "Open", Buff, "", "", 1)
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    ctemp = ctemp & "Alle Versuche Programmteile zu öffnen werden protokolliert."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If
            Case 22
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame19.Visible = False
        End Select
    Else
        Select Case index
        
            Case 39 'Verkäufe nach Zugang
                frmWK25m.Show 1
            Case 35 'Artikelgruppenliste
                frmWK25h.Show 1
            Case 33 'Markenliste
                frmWK25i.Show 1
            Case 31 'Verkaufsanteile mit und ohne Kundenbindung
                If NewTableSuchenDBKombi("VKKUUMS", gdBase) Then
                    If Datendrin("VKKUUMS", gdBase) Then
                        iRet = MsgBox("Möchten Sie die Auswertung nocheinmal neu erstellen", vbYesNo + vbQuestion, "Winkiss Frage:")
                    Else
                        iRet = vbYes
                    End If
                Else
                    iRet = vbYes
                End If
    
                If iRet = vbYes Then
                    Verkaufsanteil
                End If
                
                Label2.Visible = True
                anzeige "normal", "Druckvorschau wird erstellt...", Label2
                Screen.MousePointer = 11
                reportbildschirm "", "aZEN00s"
                anzeige "normal", "", Label2
                Screen.MousePointer = 0
                Label2.Visible = False
            Case 37 'Produktgruppenliste
                frmWK25j.Show 1
            Case 38 'Größenauswertung
                frmWK25l.Show 1
            Case 34 'Linieliste
                frmWK25k.Show 1
            Case 26 'allg.Kundenliste
            
                loeschNEW "AKUNDLI", gdBase
                CreateTable "AKUNDLI", gdBase
                
                sSQL = "Insert into AKUNDLI select "
            
                sSQL = sSQL & " TEL "
                sSQL = sSQL & ", VORNAME "
                sSQL = sSQL & ", KUNDNR "
                sSQL = sSQL & ", NAME "
                sSQL = sSQL & ", STRASSE "
                sSQL = sSQL & ", PLZ "
                sSQL = sSQL & ", stadt as ort "
                sSQL = sSQL & ", TITEL "
                sSQL = sSQL & ", FIRMA "
                sSQL = sSQL & ", RABATT "
                sSQL = sSQL & ", DATUM1 "
                sSQL = sSQL & " from Kunden  "
                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                
                reportbildschirm "dWKL12a", "aWKL47"
                loeschNEW "AKUNDLI", gdBase
                Exit Sub

            Case 24 'Rabattliste
                loeschNEW "AKUNDLI", gdBase
                CreateTable "AKUNDLI", gdBase
                
                sSQL = "Insert into AKUNDLI select "
            
                sSQL = sSQL & " TEL "
                sSQL = sSQL & ", VORNAME "
                sSQL = sSQL & ", KUNDNR "
                sSQL = sSQL & ", NAME "
                sSQL = sSQL & ", STRASSE "
                sSQL = sSQL & ", PLZ "
                sSQL = sSQL & ", stadt as ort "
                sSQL = sSQL & ", TITEL "
                sSQL = sSQL & ", FIRMA "
                sSQL = sSQL & ", RABATT "
                sSQL = sSQL & ", DATUM1 "
                sSQL = sSQL & " from Kunden  "
                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                reportbildschirm "dWKL12a", "aWKL47a"
                loeschNEW "AKUNDLI", gdBase
                Exit Sub
                reportbildschirm "dWKL12a", "aWKL47"
                Exit Sub
                
            Case 27 'bonusliste
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKundenBonus
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 28 'KE
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "Sertrag"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 29 'KU
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "Spreis"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 30 'KS
                Command1_Click 3
                Me.Refresh
                schreibeProtokollProgrammablauf " löst Liste aus    " & Command12(index).Caption
                BestKunden "smenge"
                schreibeProtokollProgrammablauf " Liste fertig      " & Command12(index).Caption
            Case 40 'KS
                frmWKL172.Show 1
            Case 36
                If gbKostenlos = True Then
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(3).Enabled = True
                    Command5(9).Enabled = True
                Else
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If

                Frame23.Visible = False
                Frame4.Enabled = True
            Case 25
                
                If gbKostenlos = True Then
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(3).Enabled = True
                    Command5(9).Enabled = True
                Else
                
                    Command5(0).Enabled = True
                    Command5(1).Enabled = True
                    Command5(2).Enabled = True
                    Command5(3).Enabled = True
                    Command5(5).Enabled = True
                    Command5(6).Enabled = True
                    Command5(7).Enabled = True
                    Command5(8).Enabled = True
                    Command5(9).Enabled = True
                End If

                Frame20.Visible = False
                Frame4.Enabled = True
            Case 21 'Bestellung berechnen
'                MsgBox "Hier 1"
                frmWKL43.Show 1
'                MsgBox "Hier 5"
            Case 23 'Bestellung manuell
                frmWKL47.Show 1
            Case Is = 0 ' aus Einzellieferung
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWKL15.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 1 'aus Filialumverteilung
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWKL23.Show '1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 2 ' Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame11.Visible = False
                
            Case Is = 3 'aus Bestellung
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWK15a.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            '*****************************************************
            Case Is = 4 'Pfad zur Datenbank
                If glLevel >= DlgZugriff(27).dZugriff Then
                    Setzedatenbankpfad
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 13
                If glLevel >= DlgZugriff(27).dZugriff Then
                
                
                    dlgPW.Show 1
                    
                    If dlgPW.Back = True Then
                        frmWKLaf.Show 1
                    Else
                            
                    End If
                
                   
                    
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 6 'Reorganisation
'                If glLevel >= DlgZugriff(26).dZugriff Then
                
                    iRet = MsgBox("Möchten Sie jetzt die Datenbank optimieren?", vbQuestion + vbYesNo, "Winkiss Frage:")
                    If iRet = vbYes Then
                        Command12_Click 5
                        Command6_Click 8
                        
                        picprogress.Visible = True
                        lbl6(28).Visible = True
                        lbl6(53).Visible = True
            
                        If BistDualleineinderDatenbank Then
                            CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige
                            
                            sSQL = "update dbeinste set lastkomp='" & Date & "'"
                            sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
                            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                        Else
                            anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
                        End If
                       
                        If BistDualleineinderDatenbankApp Then
                            dbApp_Compri "Kissapp.MDB"
                        Else
                            anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
                        End If
                       
                        picprogress.Visible = False
                        lbl6(28).Visible = False
                        lbl6(53).Visible = False
                        
                        Screen.MousePointer = 0
                        
                        Label2.ForeColor = glS1
                        Label2.Caption = "Anwender aktiv"
                        Label2.Refresh
                    End If
'                Else
'                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
'                End If
    
            Case Is = 7 'Fremddaten importieren
                If glLevel >= DlgZugriff(30).dZugriff Then

                    frmWKL153.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            
            Case Is = 8 'DB wiederherstellen
                If glLevel >= DlgZugriff(27).dZugriff Then
                    frmWKL32.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 12 'DB bereinigen
                If glLevel >= DlgZugriff(27).dZugriff Then
                
                    dlgPW.Show 1
                    
                    If dlgPW.Back = True Then
                        frmWKL33.Show 1
                    Else
                            
                    End If
                    
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 5 'Schließen Datenbank
            
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
    
                Frame5.Enabled = True
                Frame14.Visible = False
                
            Case Is = 19 'Schließen Kissnet
            
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
    
                Frame5.Enabled = True
                Frame17.Visible = False
                
            '****Ende Datenbankframe******************************
                
            Case Is = 10 'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command6(lcount).Enabled = True
                    Next lcount
                End If
                
    
                Frame5.Enabled = True
                Frame16.Visible = False
                
            Case Is = 11 ' Bereitstellung
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWKL26.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
    
            Case Is = 9 'Einlesen
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWKL27.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 20 'KISNET... Mailbox
            
                If glLevel >= DlgZugriff(5).dZugriff Then
                    frmWKL160.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 14 'Email schreiben
            
                Buff = "mailto:"
                result = ShellExecute(0&, "Open", Buff, "", "", 1)
            Case 22
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame19.Visible = False
                
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3356 Then
        MsgBox "Die Datenbank ist immernoch durch einen anderen Benutzer oder ein anderes Programm geöffnet!", vbOKOnly, "Winkiss Hinweis:"
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command12_Click"
        Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten. " & index
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub Verkaufsanteil()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim bymonat     As Byte
    Dim lJahr       As Long
    Dim dUmsatz     As Double
    Dim dumsatznull As Double
    Dim dumsatzKU   As Double
    Dim lMengeKU      As Long
    Dim lMengenull  As Long
    Dim lcount      As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    Dim iFil As Integer
    Dim lDat As Long
    
    
    Screen.MousePointer = 11
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 5
    
    loeschNEW "VKKUUMS", gdBase
    CreateTable "VKKUUMS", gdBase
   
    bymonat = Month(Now)
    lJahr = Year(Now)
    
    k = 1
    sSQL = " Insert into VKKUUMS (Nr ,Monat,Monatname,Jahr) values  (" & k & "," & bymonat & ",'" & gcMonat(bymonat) & "'," & lJahr & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    For j = 1 To 12
        bymonat = bymonat - 1
        If bymonat = 0 Then
            bymonat = 12
            lJahr = lJahr - 1
        End If
        
        k = k + 1
        
        sSQL = " Insert into VKKUUMS (Nr ,Monat,Monatname,Jahr) values  (" & k & "," & bymonat & ",'" & gcMonat(bymonat) & "'," & lJahr & ")"
        gdBase.Execute sSQL, dbFailOnError
    Next j

    
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    loeschNEW "KUNZTohneK", gdBase
    
    sSQL = "Select * into KUNZTohneK from Kassjour where  "
    sSQL = sSQL & " Kundnr = 0 "
    sSQL = sSQL & " and adate > " & CLng(DateValue(Now) - 410)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    CheckIndex "KUNZTohneK", "adate", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1
    CheckIndex "KUNZTohneK", "Filiale", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1
    CheckIndex "KUNZTohneK", "Kundnr", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1

    loeschNEW "KUNZTmitK", gdBase
    
    sSQL = "Select * into KUNZTmitK from Kassjour where  "
    sSQL = sSQL & " Kundnr > 0 "
    sSQL = sSQL & " and adate > " & CLng(DateValue(Now) - 410)
    sSQL = sSQL & " and UMS_OK = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    CheckIndex "KUNZTmitK", "adate", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1
    CheckIndex "KUNZTmitK", "Filiale", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1
    CheckIndex "KUNZTmitK", "Kundnr", "", gdBase
    txtStatus.Text = CInt(txtStatus.Text) + 1
    
    Label2.Visible = True

    Set rsrs = gdBase.OpenRecordset("VKKUUMS")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!Monat) Then
                bymonat = rsrs!Monat
            End If
            
            If Not IsNull(rsrs!jahr) Then
                lJahr = rsrs!jahr
            End If
            
            rsrs.Edit
            
            txtStatus.Text = CInt(txtStatus.Text) + 1
            
            dUmsatz = 0
            
            anzeige "normal", "Monat: " & bymonat, Label2
            dumsatznull = ermKundenumsatz(False, bymonat, lJahr, CInt(gcFilNr))
            lMengenull = ermKundenMenge(False, bymonat, lJahr)
            
            anzeige "normal", "Monat: " & bymonat & " Umsatz:" & Format(dumsatznull, "#####0.00"), Label2
            
            rsrs!Mengenull = lMengenull
            rsrs!Umsnull = dumsatznull
            
            dumsatzKU = ermKundenumsatz(True, bymonat, lJahr, CInt(gcFilNr))
            lMengeKU = ermKundenMenge(True, bymonat, lJahr)
            
            anzeige "normal", "Monat: " & bymonat & " Umsatz:" & Format(dumsatzKU, "#####0.00"), Label2
            
            rsrs!MengeKU = lMengeKU
            rsrs!UMSKU = dumsatzKU
            dUmsatz = dumsatzKU + dumsatznull
            
            If dUmsatz <> 0 Then
                rsrs!UMSKUproz = dumsatzKU * 100 / dUmsatz
                rsrs!UMSnullproz = dumsatznull * 100 / dUmsatz
            Else
                rsrs!UMSKUproz = 0
                rsrs!UMSnullproz = 0
            End If
            
            rsrs!UMSKUkunden = ermKundenumsatzproKauf(True, bymonat, lJahr, CInt(gcFilNr))
            anzeige "normal", "Monat: " & bymonat & " " & rsrs!UMSKUkunden, Label2
            rsrs!UMSNullkunden = ermKundenumsatzproKauf(False, bymonat, lJahr, CInt(gcFilNr))
            anzeige "normal", "Monat: " & bymonat & " " & rsrs!UMSNullkunden, Label2
            
            If rsrs!UMSKUkunden <> 0 Then
                rsrs!UMSKUproKauf = rsrs!UMSKU / rsrs!UMSKUkunden
            Else
                rsrs!UMSKUproKauf = 0
            End If
            
            If rsrs!UMSNullkunden <> 0 Then
                rsrs!UMSNULLproKauf = rsrs!Umsnull / rsrs!UMSNullkunden
            Else
                rsrs!UMSNULLproKauf = 0
            End If
            
            If rsrs!UMSKUkunden <> 0 Then
                rsrs!UMSKUschnitt = lMengeKU / rsrs!UMSKUkunden
            Else
                rsrs!UMSKUschnitt = 0
            End If
            
            If rsrs!UMSNullkunden <> 0 Then
                rsrs!UMSNULLschnitt = lMengenull / rsrs!UMSNullkunden
            Else
                rsrs!UMSNULLschnitt = 0
            End If
            
            
            rsrs.Update
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
'    anzeige "normal", "Druckvorschau wird erstellt...", lbl6(4)
    
    loeschNEW "KUNZTmitK", gdBase
    loeschNEW "KUNZTohneK", gdBase
    loeschNEW "KUNZTEMP", gdBase

'    reportbildschirm "aZEN00v"
    
    Label2.Visible = False
    picprogress.Visible = False

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Verkaufsanteil"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub
Private Sub Command13_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctemp As String
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command13(index))
        Select Case index
            Case Is = 0 'Bedienerstatistik
                OpenProgrammTeil frmWKL154, ermittlezugriff(byteZGNr)
            Case Is = 1 'Lieferantenstatistik
                OpenProgrammTeil frmWKLau, ermittlezugriff(byteZGNr)
            Case Is = 2 'Kundenanalyse
                OpenProgrammTeil frmWKLav, ermittlezugriff(byteZGNr)
            Case Is = 3 'Umsatzstatistik
                OpenProgrammTeil frmWK25a, ermittlezugriff(byteZGNr)
            Case Is = 4 'Geschäftsanalyse
                OpenProgrammTeil frmWKL138, ermittlezugriff(byteZGNr)
            Case Is = 5 'zurück zum Hauptmenü
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame12.Visible = False
                Label2.Visible = True
                Label3.Visible = True
                Label1(1).Visible = True
                
                bStatistiken = False
            Case Is = 6 'AGN - Statistik
                OpenProgrammTeil frmWKL18, ermittlezugriff(byteZGNr)
            Case Is = 7 'Zeitenstatistik
                OpenProgrammTeil frmWK25c, ermittlezugriff(byteZGNr)
            Case Is = 8     'Preislagenstatistik
            
                If glLevel >= ermittlezugriff(byteZGNr) Then
                    Frame22.Visible = True
                    For lcount = 0 To 8
                        Command13(lcount).Enabled = False
                    Next lcount
                    If gbKostenlos = False Then
                        Command4(13).SetFocus
                    End If
                Else
                    ctemp = "Zur Zeit ist " & gcUserName & " angemeldet." & vbCrLf
                    ctemp = ctemp & gcUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
'                    schreibeBEDProtokoll "versuchte erfolglos '" & gsProteil & "' zu öffnen"
                    MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                End If

        End Select
    Else
        Select Case index
            Case Is = 0 'Bedienerstatistik
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL154.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 1 'Lieferantenstatistik
                If glLevel >= DlgZugriff(11).dZugriff Then
    
                    frmWKLau.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 2 'Kundenanalyse
                If glLevel >= DlgZugriff(11).dZugriff Then
    
                    frmWKLav.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 3 'Umsatzstatistik
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWK25a.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 4 'Geschäftsanalyse
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL138.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 5 'zurück zum Hauptmenü
                For lcount = 0 To 5
                    Command1(lcount).Enabled = True
                Next lcount
                Command1(8).Enabled = True
                
                Frame12.Visible = False
                Label2.Visible = True
                Label3.Visible = True
                Label1(1).Visible = True
                
                bStatistiken = False
                
            Case Is = 6 'AGN - Statistik
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWKL18.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 7 'Zeitenstatistik
                If glLevel >= DlgZugriff(11).dZugriff Then
                    frmWK25c.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 8     'Preislagenstatistik...
                Frame22.Visible = True
                For lcount = 0 To 8
                    Command13(lcount).Enabled = False
                Next lcount
                If gbKostenlos = False Then
                    Command4(13).SetFocus
                End If
        End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command14_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If gbZugriffNew Then
    
        byteZGNr = ermittleTag(Command14(index))
        Select Case index
        
            Case Is = 1     'Artikel bearbeiten
                OpenProgrammTeil frmWKL10, ermittlezugriff(byteZGNr)
            Case Is = 2     'Bestandskorrektur
                OpenProgrammTeil frmWKL19, ermittlezugriff(byteZGNr)
'            Case Is = 3     'Kalkulation der Preise
'                OpenProgrammTeil frmWKL35, ermittlezugriff(byteZGNr)
            Case Is = 3     'Kalkulation der Preise
                OpenProgrammTeil frmWKL167, ermittlezugriff(byteZGNr)
            Case Is = 7     'Terminpreise
                OpenProgrammTeil frmWKL61, ermittlezugriff(byteZGNr)
            Case Is = 8     'Artikel retournieren
                OpenProgrammTeil frmWKL112, ermittlezugriff(byteZGNr)
            Case Is = 9     'Artikel löschen
                OpenProgrammTeil frmWKL113, ermittlezugriff(byteZGNr)
            Case Is = 11     'Pennerartikel
                OpenProgrammTeil frmWKL168, ermittlezugriff(byteZGNr)
            Case Is = 4     'Schließen
            
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                
                Frame13.Visible = False
            Case Is = 6     'Lieferanten bearbeiten
                gsLinr = ""
                OpenProgrammTeil frmWKL17, ermittlezugriff(byteZGNr)
            Case Is = 5     'Lieferantenverwaltung
                OpenProgrammTeil frmWKL25, ermittlezugriff(byteZGNr)
            Case Is = 10    'Lieferanten Rechnungsübersicht
                OpenProgrammTeil frmWKL130, ermittlezugriff(byteZGNr)
            Case Is = 0     'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame15.Visible = False
            
        End Select
    Else
        Select Case index
            Case Is = 1     'Artikel bearbeiten
                If glLevel >= DlgZugriff(0).dZugriff Then
                    frmWKL10.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
                
            Case Is = 2     'Bestandskorrektur
                If glLevel >= DlgZugriff(0).dZugriff Then
                    frmWKL19.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
'            Case Is = 3     'Kalkulation der Preise
'                If glLevel >= DlgZugriff(4).dZugriff Then
'                    frmWKL35.Show 1
'                Else
'                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
'                End If
            Case Is = 3     'Kalkulation der Preise
                If glLevel >= DlgZugriff(4).dZugriff Then
                    frmWKL167.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 7     'TerminPreise
                If glLevel >= DlgZugriff(4).dZugriff Then
                    frmWKL61.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 8     'Artikel retournieren
                If glLevel >= DlgZugriff(4).dZugriff Then
                    frmWKL112.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 9     'Artikel löschen
                If glLevel >= DlgZugriff(4).dZugriff Then
                    frmWKL113.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 11     'Pennerartikel
                If glLevel >= DlgZugriff(4).dZugriff Then
                    frmWKL168.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 4     'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame13.Visible = False
            Case Is = 6     'Lieferanten bearbeiten
                If glLevel >= DlgZugriff(0).dZugriff Then
                    gsLinr = ""
                    frmWKL17.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 5     'Lieferantenverwaltung
                If glLevel >= DlgZugriff(0).dZugriff Then
                    frmWKL25.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 10     'Lieferanten Rechnungsübersicht
                If glLevel >= DlgZugriff(0).dZugriff Then
                    frmWKL130.Show 1
                Else
                    MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                End If
            Case Is = 0     'Schließen
                If gbKostenlos Then
                    SetzefürKostenlos frmWKL00
                Else
                    For lcount = 0 To 8
                        Command2(lcount).Enabled = True
                    Next lcount
                End If
                Frame15.Visible = False
            
            End Select
        End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command14_Click"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Activate()
 Me.KeyPreview = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = vbKeyF8 Then
 
    'Bin ich BUDNI/EDEKA ?
     Dim rsrs As Recordset
     Set rsrs = gdBase.OpenRecordset("select * FROM LISRT WHERE FORMAT='EDIBHSG' OR FORMAT='EDIBUDNI'")
               
     If Not rsrs.EOF Then
       
           If Not gbBudniNeuesFtpVerfahren Then
                    
                 'BUDNI ----> EDEKA
                 
                 Dim ireslt1 As Integer
                 ireslt1 = MsgBox("Winkiss muss an allen Arbeitplätzen beendet werden, bevor" & vbNewLine & "Sie die Budni-EDEKA-Umstellung durchführen können.", vbQuestion + vbYesNo, "Budni-EDEKA Umstellung durchführen:")
                   
                 If ireslt1 = vbYes Then
                  FTPwechsel.Left = (Me.ScaleWidth - FTPwechsel.Width) / 2
                  FTPwechsel.Top = (Me.ScaleHeight - FTPwechsel.Height) / 2
                  FTPwechsel.Show 1
                 End If
                 
           Else
           
                 'EDEKA ----> BUDNI
                 
                 Dim ireslt2 As Integer
                 ireslt2 = MsgBox("Winkiss muss an allen Arbeitplätzen beendet werden, bevor" & vbNewLine & "Sie die Budni-EDEKA-Umstellung rückgängig machen können.", vbQuestion + vbYesNo, "Budni-EDEKA Umstellung rückgängig machen:")
                   
                 If ireslt2 = vbYes Then
                  FTPwechselAbbruch.Left = (Me.ScaleWidth - FTPwechselAbbruch.Width) / 2
                  FTPwechselAbbruch.Top = (Me.ScaleHeight - FTPwechselAbbruch.Height) / 2
                  FTPwechselAbbruch.Show 1
                 End If
                 
           End If
        
     End If
     
 ElseIf KeyCode = vbKeyF2 Then
         
        If Not FileExists(gcDBPfad & "\EineF-DateiWirdGerettet.txt") Then
            
            Dim ireslt3 As Integer
            ireslt3 = MsgBox("haben Sie zuerst geprüft, ob die fehlten F-Dateien auf 'Chipotle' auch im Ordner 'zenin' nicht existieren ?", vbQuestion + vbYesNo, "F-Dateien Wiederherstellung")
                   
            If ireslt3 = vbYes Then
            
            'F-Dateien Rettung
             FDateienRettung.Left = (Me.ScaleWidth - FDateienRettung.Width) / 2
             FDateienRettung.Top = ((Me.ScaleHeight - FDateienRettung.Height) / 2) + 200
             FDateienRettung.Show 1
             
            End If
            
        Else
        
             MsgBox ("andere Kasse rettet momentan eine F-Datei !!!" & vbNewLine & "warten Sie bitte bis diese fertig ist.")
        
        End If
 End If
 
 
 Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_KeyUp"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(5).ForeColor = glS1
    Label1(4).ForeColor = glS1
    Label1(7).ForeColor = glS1
    Label1(6).ForeColor = glS1
    Label1(8).ForeColor = glS1
    Label1(9).ForeColor = glS1
    Label1(10).ForeColor = glS1
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Update"
    
    If index = 5 Then
        Label1(5).ForeColor = glLink
    End If
    
    If index = 4 Then
        Label1(4).ForeColor = glLink
    End If
    
    If index = 6 Then
        Label1(6).ForeColor = glLink
    End If
    
    If index = 7 Then
        Label1(7).ForeColor = glLink
    End If
    
    If index = 8 Then
        Label1(8).ForeColor = glLink
    End If
    
    If index = 9 Then
        Label1(9).ForeColor = glLink
    End If
    
    If index = 10 Then
        Label1(10).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sTemp As String
    Dim cPfad As String
    Dim iErrCounter As Integer
    Dim i As Integer
    Dim lVon    As Long
    Dim lBis    As Long
    Dim lDiff1  As Long
    Dim lDiff2  As Long
    Dim lDif    As Long
    Dim iTage   As Integer
    Dim lLinr   As Long

    iErrCounter = 0
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If GBZeitschlossVersion Then
        If gsZeitschlossdate = DateValue(Now) Then
        
        Else
            gsAnzeigeText = "Das Programm wird beendet!"
            frmWKL69.Show 1
            End
        End If
    End If

startnochmaL:
    If gbLokalModus = False Then
        If Not FileExists(cPfad & "kissdata.mdb") Then
            Pause 1
            If iErrCounter < 2 Then
                iErrCounter = iErrCounter + 1
                GoTo startnochmaL
            End If
        
            If giUmleitgrund = 1 Then Exit Sub
            giUmleitgrund = 3 'Datenbank abgerissen/Netzwerkfehler
    
            gcUmleittxt = "Die Verbindung zur Datenbank wurde unterbrochen." & vbCrLf
            gcUmleittxt = gcUmleittxt & "Es liegt ein Problem mit Ihrem Netzwerk vor!" & vbCrLf
            gcUmleittxt = gcUmleittxt & "Benachrichtigen Sie Ihren Netzwerktechniker!" & vbCrLf
            
            frmWKL60.Show 1
        End If
    End If
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

'    If FileExists(cPfad & "KASSSTOP.TXT") Then
'        gsAnzeigeText = "Winkiss wird für ein paar Sekunden unterbrochen," & vbCrLf
'        gsAnzeigeText = gsAnzeigeText & "da an einem anderen Rechner die Artikeldaten aktualisiert werden." & vbCrLf
'        gsAnzeigeText = gsAnzeigeText & "Drücken Sie 'Warten', dann wird Winkiss nach ein paar Sekunden wieder freigeschaltet." & vbCrLf
'        gsAnzeigeText = gsAnzeigeText & "Drücken Sie 'Weiter', können Sie mit eventuellen Datenverlusten weiterarbeiten." & vbCrLf
'        frmWK21m.Label1.Caption = "KassStop"
'        frmWK21m.Show 1
'    End If
    
    Label1(3).Caption = Format$(Time, "HH:MM:SS")

    If gbSichernYes And giSICHTYP = 1 Then 'Sichern ja? und immer bei Progstart
    
        If Label1(3).Caption = "00:00:00" Then 'Jeden Tag um 0 Uhr aktiv schalten
            gbSichernHeut = True
        End If
        
        If Format$(Time, "SS") = 30 Then
            If gbSichernHeut Then
                
                frmWKL00.Label2.Caption = "Datenbank wird gesichert..."
                frmWKL00.Label2.Refresh
                
                DabaSicherung
                
                frmWKL00.Label2.Caption = "Anwender aktiv"
                frmWKL00.Label2.Refresh
                
                
                gbSichernHeut = False
            End If
        End If
    End If
    
    
    'Hier kommt die zeitgesteuerte Sicherung
    If gbSichernYes And giSICHTYP = 3 Then 'Sichern ja? und zeitgesteuerte Sicherung
    
        If Format$(Time, "SS") = 35 Then
                
            If Format(TimeValue(gsSICHTIME), "HH:MM") = Format(TimeValue(Now), "HH:MM") Then
                
                frmWKL00.Label2.Caption = "Datenbank wird gesichert..."
                frmWKL00.Label2.Refresh
                
                DabaSicherung
                
                frmWKL00.Label2.Caption = "Anwender aktiv"
                frmWKL00.Label2.Refresh
                
            End If
                
        End If
    End If
    
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "picture\Post.wav"
    
    'Hier die Zeitungsauswertung VMP checken und gegebenenfall erstellen und verschicken
    If gsVMPArt = "2" Then
        If Format$(Time, "SS") = 40 Then
            
            If itsVMPtime Then
                VMPZeitPunktwegschicken
            End If
                
        End If
    End If
    
    Dim bmerke  As Boolean
    bmerke = gbFTPautomatic
    
    'Hier den auto_Export_Artikelbestand checken und gegebenenfall erstellen und verschicken
    If gbAuto_Export_Artikelbestand = True Then
        If Format$(Time, "SS") = 40 Then
            If its_Export_Time Then
            
                lese_Ex_Steu
            
                If gbBL Then
                    If Export_Artikelbestände_Komplett_Vedes Then
                        If gbFtpYes Then
                            gbFTPautomatic = True
                            giKissFtpMode = 25 'FTPMODE= 25 , BEAUTY - Ordner leeren abschicken
                            frmWKL38.Show 1
                            gbFTPautomatic = bmerke
                        End If
                    End If
                End If
            
                If gbEXNOR Then
                    If Export_Artikelbestände Then
                        If gbFtpYes Then
                            gbFTPautomatic = True
                            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
                            frmWKL38.Show 1
                            
                            giKissFtpMode = 44 'FTPMODE= 44 , Out - Ordner leeren
                            frmWKL38.Show 1
                            
                            gbFTPautomatic = bmerke
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    If Format$(Time, "SS") = 10 Then
    
        Dim cpfaddb As String
                        
        cpfaddb = gcDBPfad
        If Right$(cpfaddb, 1) <> "\" Then
            cpfaddb = cpfaddb & "\"
        End If
        
        If FileExists(App.Path & "\Rose.cfg") Then
            newfnCheck4RoseSalesDaten
        End If
        'ende Rose, München
    End If
    
    If gbNacht = False And gsDabaNachtStart <> "" Then
        If Format$(Time, "SS") = 40 Then
        
            Dim sZeitNachtstartDB As Date
            Dim sJetztDB As Date
            sZeitNachtstartDB = Format(TimeValue(gsDabaNachtStart), "HH:MM")
            sJetztDB = Format(TimeValue(Now), "HH:MM")
            If sZeitNachtstartDB = sJetztDB Then
                Timer1.Enabled = False
                gdBase.Close
                schreibeProtokoll "Abmeldung: meldet sich ab(kissdata.mdb)."
                schreibeProtokollBENUTZERablauf "Abmeldung"
                gdApp.Close
                schreibeProtokoll "Abmeldung: meldet sich ab(kissapp.mdb)."
                
                End 'Winkiss beenden
            End If
        
        End If
    End If
    
    If gbNacht = True Then
        If Format$(Time, "SS") = 50 Then
        
            Dim sZeitNachtstart As Date
            Dim sJetzt As Date
        
            'immer zur 50 Sekunde prüfen ob Nachtverarbeitung starten muss
            
            'stimmt die Zeit?
            Dim dNS As Date
            Dim dNow As Date
            
            Dim dAbzug As Date
            Dim dDiff As Date
            Dim dDiff1 As Date
            
            dAbzug = Format(TimeValue("00:10:00"), "HH:MM")
            
        
            sZeitNachtstart = Format(TimeValue(gsNachtstart), "HH:MM")
            dNS = Format(TimeValue(gsNachtstart), "HH:MM")
            sJetzt = Format(TimeValue(Now), "HH:MM")
            dNow = Format(TimeValue(Now), "HH:MM")
            
            dDiff = dNS - dAbzug
            
            dDiff = Format(TimeValue(dDiff), "HH:MM")

            
            dDiff1 = dNow + dAbzug
            dDiff1 = Format(TimeValue(dDiff1), "HH:MM")

            
            If dDiff1 < dNS Then
            
            Else
            
                If dNow > dNS Then
                
                Else
                    If dDiff <= dNow Then
                        anzeige "LASER", "Die Nachtverarbeitung startet um " & Format(TimeValue(dNS), "HH:MM") & " Uhr", frmWKL00.Label2
                    End If
                End If
            End If
            
            If sZeitNachtstart = sJetzt Then
                Timer1.Enabled = False
                
                anzeige "normal", "Die Nachtverarbeitung startet jetzt", frmWKL00.Label2
                schreibeProtokollNachtAblauf "Die Nachtverarbeitung startet jetzt"
                'dann arbeite mal die Nachtverarbeitung ab
                
                If gbHauptg = True Then 'das ist Hauptgeschäft
                    dieNachtverarbeitungHauptG
                    
                    
                ElseIf gbBestAkt = True Then 'das ist Lager
                
                    
                
                    'die Nachtverarbeitung führen wir nur durch, wenn eine Abschlussdatei vorliegt
                    schreibeProtokollNachtAblauf "Kassendateiprüfung jetzt"
                    If Kassendateipruefung_bestanden Then
                        schreibeProtokollNachtAblauf "Kassendateiprüfung erfolgreich"
                        dieNachtverarbeitungLager
                    Else
                        schreibeProtokollNachtAblauf "Abbruch der Nachtverarbeitung (Kassendatei fehlt)"
                        MsgBox "Die Nachtverarbeitung wurde nicht durchgeführt, weil die heutige Kassendatei fehlt.", vbInformation, "Winkiss Hinweis:"
                        Exit Sub
                    End If
                Else
                    dieNachtverarbeitung
                    'fertig mit der Nachtverarbeitung
                End If
                schreibeProtokollNachtAblauf "Die Nachtverarbeitung endet jetzt"
                
                Pause 1
                
                If gbBR Then
                    schreibeProtokollNachtAblauf "Bestellvorschläge erstellen beginnt"
'                    bestellvorschlagrechnen

                    lbl6(28).Visible = True
                    ErmittleBestVorschlag lbl6(28)
                    lWEinBESTLIN lbl6(28)

                    schreibeProtokollNachtAblauf "Bestellvorschläge erstellen endet"
                End If
                
                If gbSTAMDA Then
                    schreibeProtokollNachtAblauf "Stammdaten einlesen beginnt"
'                    Stammdateneinlesen
                    schreibeProtokollNachtAblauf "Stammdaten einlesen endet"
                End If
                
                If gbKABSCH Then 'Lagerumschlag rechnen
                    If LUGSAktuell = False Then
                        schreibeProtokollNachtAblauf "Lagerumschläge werden geschrieben"
            
                        alleLUGnachLief txtStatus, picprogress, False
            
                        schreibeProtokollNachtAblauf "Lagerumschläge erstellen endet"
                    End If
                End If
                
                If gbUmsatzNeu Then
                    schreibeProtokollNachtAblauf "Artikelumsätze werden geschrieben"
                
                    UmsartjNew frmWKL00.Label2
                    Ums_artNew frmWKL00.Label2
                    
                    schreibeProtokollNachtAblauf "Artikelumsätze erstellen endet"
                End If
                
                If gbMB Then
                    schreibeProtokollNachtAblauf "Mindestbestände werden geschrieben"
                    
                    leseMBDetails
        
                    Select Case MBDETAILMON
                        Case 5 '9
                            iTage = 272
                        Case 4 '8
                            iTage = 241
                        Case 3 '7
                            iTage = 211
                        Case 2 '6
                            iTage = 180
                        Case 1 '5
                            iTage = 150
                        Case 0 '4
                            iTage = 119
                        Case Else
                            iTage = 180
                    End Select
                
                    lVon = DateValue(Now) - iTage
                    lBis = DateValue(Now)
                    
                    lDiff1 = lBis - lVon
                    lDiff2 = MBDETAILBIS - MBDETAILVON
                    
                    If MBDETAILVON <= lBis And MBDETAILBIS <= lBis And MBDETAILVON >= lVon And MBDETAILBIS >= lVon Then
                        lDif = lDiff1 - lDiff2
                    ElseIf MBDETAILVON <= lBis And MBDETAILBIS > lBis And MBDETAILVON >= lVon Then
                        lDif = MBDETAILVON - lVon
                    ElseIf MBDETAILBIS <= lBis And MBDETAILVON < lVon And MBDETAILBIS >= lVon Then
                        lDif = lBis - MBDETAILBIS
                    ElseIf MBDETAILVON < lVon And MBDETAILBIS < lVon Then
                        lDif = lDiff1
                    ElseIf MBDETAILVON > lBis And MBDETAILBIS > lBis Then
                        lDif = lDiff1
                    End If
                    
                    schreibeProtokollNachtAblauf MBDETAILBVO & ", " & CInt(lDif) & ", " & 1 & ", " & lVon & ", " & lBis & ", " & lbl6(28) & ", " & MBDETAILVON & ", " & MBDETAILBIS
                    MBrechnen1 MBDETAILBVO, CInt(lDif), 1, lVon, lBis, lbl6(28), MBDETAILVON, MBDETAILBIS
                    
                    schreibeProtokollNachtAblauf "Mindestbestände erstellen endet"
                End If
                
                
                
                
                'Grund.cfg
                
                cpfaddb = gcDBPfad
                If Right$(cpfaddb, 1) <> "\" Then
                    cpfaddb = cpfaddb & "\"
                End If
                
                If FileExists(cpfaddb & "Grund.cfg") Then
                    Schreibe_VKDATEN
                End If
                'Ende Grund.cfg

                
                
                
                
                schreibeProtokollNachtAblauf "Komprimierung beginnt"
                
                
                dabkomp1
                
                
                lbl6(28).ForeColor = vbRed
                lbl6(28).Caption = "Noch nicht ausschalten!"
                lbl6(28).Refresh
                
                
                lbl6(53).ForeColor = vbRed
                lbl6(53).Caption = "KISSAPP wird komprimiert..."
                lbl6(53).Refresh
                
                If BistDualleineinderDatenbankApp Then
                    dbApp_Compri "Kissapp.MDB"
                Else
                    anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
                End If
                
                lbl6(53).ForeColor = vbRed
                lbl6(53).Caption = "GDPDU wird komprimiert..."
                lbl6(53).Refresh
                
                If BistDualleineinderDatenbankGDPDU Then
                    GDPDU_GLAGER_KLEINHALTEN lbl6(28)
                    dbGDPDU_Compri "GDPDU.MDB", lbl6(28)
                Else
                    anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
                End If
                
                
                
                
                
                lbl6(53).ForeColor = vbRed
                lbl6(53).Caption = "KASSBON wird komprimiert..."
                lbl6(53).Refresh
                
                If BistDualleineinderDatenbankKASSBON Then
                    dbKASSBON_Compri "KASSBON.MDB"
                Else
                    anzeige "rot", "Fehler - Es sind noch andere Benutzer in der Datenbank", lbl6(28)
                End If
                
                Dim sTabc As String
                sTabc = kassetabcheck(gdBase, lbl6(53), lbl6(28))
            
        
                If sTabc = "" Then
            
                Else
            '        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
            '                End
                End If
                
                
                schreibeProtokollNachtAblauf "Komprimierung endet"
                
                
                
                
                If gbDSL Then
                    If gbEXTSICH Then
                        schreibeProtokollNachtAblauf "externe Sicherung gestartet"
                        
                        lbl6(28).Visible = True
                        lbl6(28).ForeColor = vbRed
                        
                        ExternSichern txtStatus, lbl6(28)
                        
                        lbl6(28).Visible = False
                        lbl6(28).ForeColor = glS1
                    
                        schreibeProtokollNachtAblauf "externe Sicherung beendet"
                    End If
                End If
                
                If gbSichernYes And giSICHTYP = 2 Then 'Sichern ja? und immer der Nachtv
    
                    schreibeProtokollNachtAblauf "lokale Sicherung gestartet"
                            
                    DabaSicherung
                    
                    schreibeProtokollNachtAblauf "lokale Sicherung beendet"
                    
                        
                End If
                
                
                
                If gsNachtVerarbeitungMail <> "" Then
        
                    'schicke Mail an die hinterlegte Adresse
                    Dim cAbsenderEmail As String
                    Dim cAnEmailadresse As String
                    Dim cBetreff As String
                    Dim cMessagetext As String
                    Dim sAttachment As String
                
                    sAttachment = ""
                    
                    cAbsenderEmail = ermFirmenMail
                    If cAbsenderEmail = "" Then
                        Exit Sub
                    End If
                    
                    cAnEmailadresse = gsNachtVerarbeitungMail
                    cBetreff = "Die Nachtverarbeitung im Winkiss war erfolgreich"
                    
                    cMessagetext = "Die Nachtverarbeitung im Winkiss war erfolgreich" & vbCrLf & vbCrLf
                    cMessagetext = cMessagetext & "Bitte beantworten Sie diese Email nicht."
                    
                    schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, "", cAnEmailadresse _
                    , "bestsend@kisswws.de", gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, cBetreff, cMessagetext, sAttachment
                    
                End If
                
                
                
                
                
                
                
                
                
                If gbPCAus = True Then
                    schreibeProtokollNachtAblauf "Computer soll herunterfahren"
                    allesbeenden
                
                    lbl6(28).Visible = True
                    lbl6(28).ForeColor = vbRed
                    lbl6(28).Caption = "Nachtverarbeitung: PC wird heruntergefahren..."
                    lbl6(28).Refresh
                
                    
                    
                    Timer1.Enabled = False
                    
                    schreibeProtokollNachtAblauf "Computer wird herunterfahren"
                    If SystemDown Then
                        Me.Caption = "Fahre herunter..."
                    Else
                        Me.Caption = "Fehler beim Herunterfahren"
                        schreibeProtokollNachtAblauf "Fehler beim Herunterfahren"
                    End If
                    schreibeProtokollNachtAblauf "Winkiss wird beendet"
                    End
                Else

                    anzeige "normal", "", frmWKL00.Label2
                    anzeige "normal", "letzte Komprimierung:", lbl6(28)
                    anzeige "normal", DateValue(Now) & "   " & TimeValue(Now) & " Uhr", lbl6(53)
                    
                    'test neu start
                    Dim Task$
                    AbmeldungDabaNew
                    
                    
                    'Wir starten Winkiss neu oder beenden Winkiss, so wird auch das eventuell gestartete ZVT-Programm beendet
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    If gsEPartner = "ZVT" Then
                        lese_ZVT_opt
                        
                        'close anwendung
                        Dim hwnd&
                        Dim Y As String
                        Dim result&
                        Dim Title$
                    
                        Y = gZVTPTitel
                                                
                        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                    
                        Do
                            result = GetWindowTextLength(hwnd) + 1
                            Title = Space(result)
                            result = GetWindowText(hwnd, Title, result)
                            Title = Left$(Title, Len(Title) - 1)
                    
                            If InStr(1, Title, Y) Then
                                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                    
                            End If
                    
                            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                        Loop Until hwnd = 0
                    End If
                    
                    If gbBestDateien = True And gsPfadBestandlive <> "" Then
    
                        Y = "BestandLive"
                                                    
                        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                        
                        Do
                            result = GetWindowTextLength(hwnd) + 1
                            Title = Space(result)
                            result = GetWindowText(hwnd, Title, result)
                            Title = Left$(Title, Len(Title) - 1)
                    
                            If InStr(1, Title, Y) Then
                                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                    
                            End If
                    
                            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                        Loop Until hwnd = 0
                    End If
                    
                    
                    
                    
                    
                    'auch die Display.exe
                    
                    If gbZweitMoni Then
    
                        Y = "Ihre Kundeninformationen"
                                                    
                        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                        
                        Do
                            result = GetWindowTextLength(hwnd) + 1
                            Title = Space(result)
                            result = GetWindowText(hwnd, Title, result)
                            Title = Left$(Title, Len(Title) - 1)
                    
                            If InStr(1, Title, Y) Then
                                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                    
                            End If
                    
                            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                        Loop Until hwnd = 0
                    End If
                    
                    If gbWKAUS = True Then
                    
                        gdBase.Close
                        schreibeProtokoll "Abmeldung: meldet sich ab(kissdata.mdb)."
                        schreibeProtokollBENUTZERablauf "Abmeldung"
                        gdApp.Close
                        schreibeProtokoll "Abmeldung: meldet sich ab(kissapp.mdb)."
                        
                        End 'Winkiss beenden
                    
                    Else
            
                        schreibeProtokollUNITXT "Winkiss startet neu", "Start"
                        
                        Task = Shell(App.Path & "\WKSTART.exe", 1) 'WKSTART öffnen
                        Screen.MousePointer = 0
                        gdBase.Close
                        schreibeProtokoll "Abmeldung: meldet sich ab(kissdata.mdb)."
                        schreibeProtokollBENUTZERablauf "Abmeldung"
                        gdApp.Close
                        schreibeProtokoll "Abmeldung: meldet sich ab(kissapp.mdb)."
                        
                        End 'Winkiss beenden
                        
                        
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Schreibe_VKDATEN()
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim cSatz       As String
    Dim rsrs        As Recordset
    Dim cSQL        As String
    Dim sTime       As String
    Dim sDate       As String
    
    Dim cPfad2      As String
    
    cPfad2 = gcDBPfad
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    sTime = Format$(TimeValue(Now), "HHMMSS")
    sDate = Format$(DateValue(Now), "DDMMYYYY")
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "STAT\"
    
    Kill cPfad & "VK_" & gcFilNr & "_" & sDate & "_" & sTime & ".csv"
    
    iFileNr = FreeFile
    Open cPfad & "VK_" & gcFilNr & "_" & sDate & "_" & sTime & ".csv" For Binary As #iFileNr
    
    cSQL = "Select * from AFCBUCH_GRUND "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!EAN) Then
            
                cSatz = ""
                cSatz = cSatz & rsrs!EAN & vbTab
                cSatz = cSatz & rsrs!BEZEICH & vbTab
                cSatz = cSatz & rsrs!Menge & vbTab
                cSatz = cSatz & rsrs!ADATE & vbTab
                cSatz = cSatz & rsrs!AZEIT
                cSatz = cSatz & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr
    
    
    cSQL = "Delete * from AFCBUCH_GRUND"
    gdBase.Execute cSQL, dbFailOnError
    
    Dim bmerke  As Boolean
    bmerke = gbFTPautomatic

    If gbFtpYes Then

        gbFTPautomatic = True
        giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
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
        Fehler.gsFunktion = "Schreibe_VKDATEN"
        Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function Kassendateipruefung_bestanden() As Boolean
On Error GoTo LOKAL_ERROR

    Kassendateipruefung_bestanden = False
    
    Dim sSQL            As String
    Dim lHeute          As Long
    Dim lDateiDatum     As Long
    Dim rsrs            As DAO.Recordset
    
    lHeute = Fix(Now)
    
    sSQL = "Select Max(Datum) as Maxdat from LASTSEND"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        lDateiDatum = rsrs!Maxdat
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lHeute = lDateiDatum Then
        Kassendateipruefung_bestanden = True
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kassendateipruefung_bestanden"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Import_Libri(cpfad_Dat As String, lLinr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim lAnz            As Long
    Dim cSatz1          As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim i               As Integer

    Dim lPosEnde As Long
    Dim cEinzelsatz As String
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lposSemiEnde As Long
    Dim cWert As String
    Dim lfnr1 As Long
    Dim cLEKPR As String
    
    
    Import_Libri = False
    
    loeschNEW "LIBRI", gdBase
    CreateTableT2 "LIBRI", gdBase
        
    lPos = 1
    lPosEnde = 1
    lposSemiEnde = 1
    
    Set rsrs = gdBase.OpenRecordset("LIBRI")
    
    
    iFileNr = FreeFile
    Open cpfad_Dat For Binary As #iFileNr
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
            lfnr1 = lfnr1 + 1
            rsrs!lfnr = lfnr1
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!EAN = cWert
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!LIBESNR = cWert
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            cWert = SwapStr(cWert, "'", " ")
            cWert = SwapStr(cWert, ";", " ")
            cWert = SwapStr(cWert, ",", " ")
            cWert = SwapStr(cWert, "*", " ")
            cWert = SwapStr(cWert, "  ", " ")
            
            rsrs!BEZEICH = Left(cWert, 35)
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            rsrs!NOTIZEN = Left(cWert, 40)
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            rsrs!PRODUKTNR = cWert
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            cWert = SwapStr(cWert, " EUR", "")
            cWert = SwapStr(cWert, ".", ",")
            If IsNumeric(cWert) Then
                rsrs!KVKPR1 = cWert
            Else
                rsrs!KVKPR1 = 0
            End If
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!Produktkuerzel = cWert
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab)
            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            
            rsrs!Produktart = cWert
            rsrs!linr = lLinr
            rsrs.Update
            
        Loop While lLenfil >= lPos
    End If
    
    Close iFileNr
    rsrs.Close: Set rsrs = Nothing
    
    'jetzt in die Artikel
    
    Dim cEAN As String
    Dim cLiBesNr As String
    Dim cBezeich As String
    Dim cNOTIZEN As String
    Dim cArtNr As String
    Dim cPreisKenn As String
    Dim cInhaltBez As String
    Dim dKVkPr1 As Double
    
    Dim cSuchArt As String
    
    Dim rsArt As Recordset
    Dim rsLIB As Recordset
    
    
    Set rsLIB = gdBase.OpenRecordset("LIBRI")
    If Not rsLIB.EOF Then
        rsLIB.MoveFirst
        
        Do While Not rsLIB.EOF
        
            cEAN = ""
            cLiBesNr = ""
            cBezeich = ""
            cNOTIZEN = ""
            cArtNr = ""
            cPreisKenn = ""
            cInhaltBez = ""
            dKVkPr1 = 0
            cSuchArt = ""
        
            If Not IsNull(rsLIB!EAN) Then
                cEAN = rsLIB!EAN
            End If
            
            If Not IsNull(rsLIB!LIBESNR) Then
                cLiBesNr = rsLIB!LIBESNR
            End If
            
            If Not IsNull(rsLIB!Produktkuerzel) Then
                cPreisKenn = rsLIB!Produktkuerzel
            End If
            
            If Not IsNull(rsLIB!KVKPR1) Then
                dKVkPr1 = rsLIB!KVKPR1
            End If
            
            If Not IsNull(rsLIB!BEZEICH) Then
                cBezeich = rsLIB!BEZEICH
            End If
            
            If Not IsNull(rsLIB!NOTIZEN) Then
                cNOTIZEN = rsLIB!NOTIZEN
            End If
            
            If Not IsNull(rsLIB!Produktart) Then
                cInhaltBez = rsLIB!Produktart
            End If
            
            
            If cLiBesNr <> "" And cEAN <> "" Then
                
                cSuchArt = Artikelvorhanden(cEAN, cLiBesNr, lLinr)
                If cSuchArt > 0 Then
                    'updaten
                    Set rsrs = gdBase.OpenRecordset("Select * from Artikel where artnr = " & cSuchArt)
                    
                    If Not rsrs.EOF Then
                        rsrs.Edit
                        rsrs!KVKPR1 = dKVkPr1
                        rsrs.Update
                    End If
                    rsrs.Close
                    
                Else
                    'anlegen
                    Set rsrs = gdBase.OpenRecordset("Select * from Artikel where artnr = -1")
                    
                    cArtNr = HoleFreieArtikelNrWKL10
                    
                    If Val(cArtNr) > 0 Then
                        rsrs.AddNew
                        
                        rsrs!SYNStatus = "A"
                        rsrs!AUFDAT = DateValue(Now)
                        rsrs!EXDAT = 0
                        
                        rsrs!artnr = cArtNr
                        rsrs!linr = lLinr
                        rsrs!BEZEICH = cBezeich
                        rsrs!NOTIZEN = cNOTIZEN
                        rsrs!KVKPR1 = dKVkPr1
                        
                        rsrs!LPZ = 1
                        rsrs!AGN = 617
                        rsrs!PGN = Null
                        rsrs!RKZ = "N"
                        rsrs!lekpr = 0
                        rsrs!vkpr = 0
                        rsrs!BESTAND = 0
                        If cPreisKenn = "fPr" Then
                            rsrs!MWST = "V"
                        Else
                            rsrs!MWST = "E"
                        End If
                        
                        rsrs!GROESSE = ""
                        rsrs!LIBESNR = cLiBesNr
                        rsrs!EAN = cEAN
        
                        rsrs!INHALT = 0
                        rsrs!INHALTBEZ = Left(cInhaltBez, 3)
                        rsrs!GRUNDPREIS = "N"
                        rsrs!MINBEST = 0
                        rsrs!RABATT_OK = "N"
                        rsrs!UMS_OK = "J"
                        rsrs!GEFUEHRT = "J"
                        rsrs!MINMEN = 0
                        rsrs!ekpr = 0
                        rsrs!PREISSCHU = "N"
                        rsrs!BONUS_OK = "N"
                        rsrs!UMS_OK = "J"
                        rsrs!AWM = "0"
        
                        rsrs.Update
                        rsrs.Close
                        
                       
                        Set rsArt = gdBase.OpenRecordset("Select * from Artlief where artnr = -1")
                        rsArt.AddNew
                        rsArt!SYNStatus = "A"
                        rsArt!artnr = cArtNr
                        rsArt!linr = lLinr
                        rsArt!lekpr = 0
                        rsArt!LIBESNR = cLiBesNr
                        rsArt!MINMEN = 0
                        rsArt!SPANNE = 0
                    
                        rsArt.Update
                        rsArt.Close
                    Else
                        MsgBox "Keine freien Artikelnummern vorhanden!", vbInformation, "Winkiss Hinweis:"
                        Exit Function
                    End If
                
                End If
                
            End If
                

                
            
            rsLIB.MoveNext
        Loop

    End If
    rsLIB.Close: Set rsLIB = Nothing
    
    
    
    
    
    Import_Libri = True
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Import_Libri"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Function
Private Function Artikelvorhanden(sEAN As String, slibesnr As String, lLinr As Long) As Long
On Error GoTo LOKAL_ERROR

    Artikelvorhanden = 0
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSQL = "select * from artikel where ean = '" & sEAN & "'"
    sSQL = sSQL & " or ean2 = '" & sEAN & "'"
    sSQL = sSQL & " or ean3 = '" & sEAN & "'"
        
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!artnr) Then
            Artikelvorhanden = rsrs!artnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Artikelvorhanden > 0 Then
        Exit Function
    End If
    
    sSQL = "select * from artlief where libesnr = '" & slibesnr & "'"
    sSQL = sSQL & " and linr  = " & lLinr
        
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!artnr) Then
            Artikelvorhanden = rsrs!artnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Artikelvorhanden"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function checkLinrForLIBRI() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim rsLi As Recordset
    
    checkLinrForLIBRI = 0

    sSQL = "Select libLINR from LIBRIe "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!libLINR) Then
            checkLinrForLIBRI = rsrs!libLINR
            
            sSQL = "Select * from LISRT where LINR = " & checkLinrForLIBRI
            sSQL = sSQL & " and ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
            Set rsLi = gdBase.OpenRecordset(sSQL)
            If rsLi.RecordCount = 0 Then
                checkLinrForLIBRI = 0
            End If
            rsLi.Close
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If checkLinrForLIBRI = 0 Then
    
        Screen.MousePointer = 0
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
        If gF2Prompt.cFeld <> "" Then
            gsAnzeige00a = "Bitte den LIBRI - Lieferant auswählen!"
            frmWK00a.Show 1
        End If
        gsAnzeige00a = ""
        
'        anzeige "normal", "Der Lieferant: " & gF2Prompt.cWahl & " wurde zugeordnet."
        
        If gF2Prompt.cWahl <> "" Then
             checkLinrForLIBRI = CDbl(gF2Prompt.cWahl)
        End If
        
        If checkLinrForLIBRI <> 0 Then
        
            loeschNEW "LIBRIE", gdBase
            CreateTableT2 "LIBRIE", gdBase
            sSQL = "Insert Into LIBRIe (LIBlinr) values (" & checkLinrForLIBRI & ")"
            gdBase.Execute sSQL, dbFailOnError
            
        End If
        
    End If
    
    
   
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkLinrForLIBRI"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Private Function itsVMPtime() As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsZEIT      As Recordset
    Dim cFeld       As String
    
    itsVMPtime = False
    
    If NewTableSuchenDBKombi("TAGVMP", gdApp) Then
    
        If Trim(gcTag) = "" Then
         Exit Function
        End If
    
        sSQL = "select * from TAGVMP where Tag =  '" & WeekdayName(gcTag) & "'"
        Set rsrs = gdApp.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            sSQL = "select * from ZEITVMP where Tag =  '" & WeekdayName(gcTag) & "'"
            sSQL = sSQL & " order by ZEIT  "
            Set rsZEIT = gdApp.OpenRecordset(sSQL)
            If Not rsZEIT.EOF Then
                rsZEIT.MoveFirst
                Do While Not rsZEIT.EOF
                    If Not IsNull(rsZEIT!zeit) Then
                        cFeld = rsZEIT!zeit
                        If Format$(TimeValue(cFeld), "HH:MM") = Format$(TimeValue(Now), "HH:MM") Then
                            rsZEIT.Close
                            rsrs.Close: Set rsrs = Nothing
                            itsVMPtime = True
                            Exit Function
                        End If
                    End If
                    rsZEIT.MoveNext
                Loop
            End If
            rsZEIT.Close
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "itsVMPtime"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function its_Export_Time() As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsZEIT      As Recordset
    Dim cFeld       As String
    
    its_Export_Time = False
    
    If NewTableSuchenDBKombi("TAGAEA", gdApp) Then
        sSQL = "select * from TAGAEA where Tag =  '" & WeekdayName(gcTag) & "'"
        Set rsrs = gdApp.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            sSQL = "select * from ZEITAEA where Tag =  '" & WeekdayName(gcTag) & "'"
            sSQL = sSQL & " order by ZEIT  "
            Set rsZEIT = gdApp.OpenRecordset(sSQL)
            If Not rsZEIT.EOF Then
                rsZEIT.MoveFirst
                Do While Not rsZEIT.EOF
                    If Not IsNull(rsZEIT!zeit) Then
                        cFeld = rsZEIT!zeit
                        If Format$(TimeValue(cFeld), "HH:MM") = Format$(TimeValue(Now), "HH:MM") Then
                            rsZEIT.Close
                            rsrs.Close: Set rsrs = Nothing
                            its_Export_Time = True
                            Exit Function
                        End If
                    End If
                    rsZEIT.MoveNext
                Loop
            End If
            rsZEIT.Close
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "its_Export_Time"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub bestellvorschlagrechnen()
On Error GoTo LOKAL_ERROR

    Dim byteBRformorgen As Byte

    byteBRformorgen = ermbrheute
    
    Select Case byteBRformorgen
    Case 6
        byteBRformorgen = 8
    Case 7
        byteBRformorgen = 8
    Case 13
        byteBRformorgen = 1
    Case 14
        byteBRformorgen = 1
    Case 255
        Exit Sub
    Case Else
        byteBRformorgen = byteBRformorgen + 1
    End Select
    bestellvor byteBRformorgen
        
    
                    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "bestellvorschlagrechnen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub bestellvor(bytebr As Byte)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lLinr As Long
    
    sSQL = "Select * from LISRT where br = " & bytebr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
                erstellevorschlag CStr(lLinr), Label2, Label3
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
    Fehler.gsFunktion = "bestellvor"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub dieNachtverarbeitung()
On Error GoTo LOKAL_ERROR

    Dim bmerke  As Boolean
    Dim iRet    As Long
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    
    bmerke = gbFTPautomatic
    k = 1
    
    If gbUKDAT = True Then
        If gbFtpYes Then
            If gbFtpZENT Then
                
                schreibeProtokollNachtAblauf "Übertragung der Kassendateien beginnt"
                gbFTPautomatic = True
                giKissFtpMode = 10   '6 'FTPMODE= 6 , gsZoutpfad - Ordner leeren abschicken
                frmWKL38.Show 1
                gbFTPautomatic = bmerke
                schreibeProtokollNachtAblauf "Übertragung der Kassendateien endet"
            End If
        End If
    End If
    
    If gbUSTAT = True Then
        If gbFtpYes Then
        
            schreibeProtokollNachtAblauf "Übertragung der Statistikdateien beginnt"
            gbFTPautomatic = True
            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
            frmWKL38.Show 1
            gbFTPautomatic = bmerke
            schreibeProtokollNachtAblauf "Übertragung der Statistikdateien endet"
        
        End If
    End If
    
    
    
    If gbUPRO = True Or gbUSTADA = True Then
        If gbFtpYes Then
        
            If gbUKDAT = True Then
            
            Else
        
                schreibeProtokollNachtAblauf "Übertragung der Stammdaten und Programmupdates beginnt"
                gbFTPautomatic = True
                giKissFtpMode = 1 'FTPMODE= 1
                frmWKL38.Show 1
                
                
                gbFTPautomatic = bmerke
                schreibeProtokollNachtAblauf "Übertragung der Stammdaten und Programmupdates endet"
        
            End If
        End If
    End If
    
    iRet = 0
    
'    Stehen aus irgendeinem Grund KD an dann einlesen

    schreibeProtokollNachtAblauf "Prüfung, ob Kassendateien vorliegen"

    gbFTPautomatic = True
    iRet = newfnCheck4KassenDateiWKL00()
    gbFTPautomatic = bmerke
    
    If iRet <> 0 Then
        schreibeProtokollNachtAblauf "Übertragung der Kassendateien endet"
        Exit Sub
    End If
    
    
    If gbEKDAT = True Then
        If gbFtpYes Then
            If gbFtpZENT Then
            
                schreibeProtokollNachtAblauf "Übertragung der Kassendateien beginnt"
            
'                Pause 1
                
'                Label3.Visible = True
'                Label1(1).Visible = True
                
                '1.Versuch
                
'                anzeige "normal", "1. Versuch die Kassendateien ", Label2
'                anzeige "normal", "von der Zentrale abzuholen ", Label3
                
'                schreibeProtokollNachtAblauf "1. Versuch die Kassendateien abzuholen"
                
'                'gistartmin
'
'                For j = giSTARTMIN - 1 To 0 Step -1
'                    Me.Refresh
'                    For i = 59 To 0 Step -1
'
'
'                        Label1(1).Caption = "in " & j & ":" & Format(i, "0#") & " Minuten "
'                        Label1(1).Refresh
'
'                        Pause 1
'                    Next i
'                Next j
                
                gbFTPautomatic = True
                
                giKissFtpMode = 8 'FTPMODE= 8,Kassendateien holen
                frmWKL38.Show 1
                
                gbFTPautomatic = True
                iRet = newfnCheck4KassenDateiWKL00()
                gbFTPautomatic = bmerke
                
                If iRet = 0 Then
'                    Do While iRet = 0
'                        k = k + 1
'                        If k > 6 Then '12
'                            anzeige "normal", "Alle Versuche Kassendateien", Label2
'                            anzeige "normal", "von der Zentrale abzuholen scheiterten.", Label3
'                            schreibeProtokollNachtAblauf "Alle Versuche Kassendateien abzuholen scheiterten"
'                            Label1(1).Visible = False
'                            Label1(1).Refresh
'
'                            Exit Do
'                        End If
'
'                        anzeige "normal", k & ". Versuch die Kassendateien ", Label2
'                        anzeige "normal", "von der Zentrale abzuholen ", Label3
'                        schreibeProtokollNachtAblauf "Versuch (" & k & ". ) die Kassendateien abzuholen"
'
'                        For j = giINTERV - 1 To 0 Step -1 ' J = 15
'                            Me.Refresh
'                            For i = 59 To 0 Step -1
'
'
'                                Label1(1).Caption = "in " & j & ":" & Format(i, "0#") & " Minuten "
'                                Label1(1).Refresh
'
'                                Pause 1
'                            Next i
'                        Next j
'
'                        gbFTPautomatic = True
'
'                        giKissFtpMode = 8 'FTPMODE= 8,Kassendateien holen
'                        frmWKL38.Show 1
'
'                        gbFTPautomatic = True
'                        iRet = newfnCheck4KassenDateiWKL00()
'                        gbFTPautomatic = bmerke
'                    Loop
                End If
                
                schreibeProtokollNachtAblauf "Übertragung der Kassendateien endet"

            End If
        End If
    End If
    ' sollen auch die Kassendateien eingelesen werden?

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dieNachtverarbeitung"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub check_Lüning_Stammdaten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
                
    Dim sLüningKundnr As String
    Dim rsLi As DAO.Recordset
    
    sLüningKundnr = ""
    
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDILUENING' "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sLüningKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(sLüningKundnr) > 0 Then

        'dann bau mal die Verbindung auf und hol alles für den Kunden ab

        giKissFtpMode = 37
        frmWKL38.Show 1

    End If

    
Exit Sub

LOKAL_ERROR:
    
     Fehler.gsDescr = err.Description
     Fehler.gsNumber = err.Number
     Fehler.gsFormular = Me.name
     Fehler.gsFunktion = "check_Lüning_Stammdaten"
     Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
     
     Fehlermeldung1

End Sub

Private Sub dieNachtverarbeitungHauptG()
On Error GoTo LOKAL_ERROR

    Dim bmerke  As Boolean
    Dim iRet    As Long
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    
    bmerke = gbFTPautomatic
    k = 1
    
'    If gbFtpYes Then
'        gbFTPautomatic = True
'        giKissFtpMode = 1 'FTPMODE= 1
'        frmWKL38.Show 1
'        gbFTPautomatic = bmerke
'    End If
    
    schreibeProtokollNachtAblauf "1. Versuch die Lagerdateien abzuholen"
    
    If gbFtpYes Then
        gbFTPautomatic = True
        giKissFtpMode = 20 'FTPMODE= 20,Lagerdateien holen
        frmWKL38.Show 1
        gbFTPautomatic = bmerke
    End If
        
    iRet = 0
    
'    Stehen aus irgendeinem Grund KD an dann einlesen
    schreibeProtokollNachtAblauf "Prüfung, ob Lagerdateien vorliegen"

    gbFTPautomatic = True
    iRet = newfnCheck4LagerHauptgDateien()
    gbFTPautomatic = bmerke
    
    If iRet <> 0 Then
        schreibeProtokollNachtAblauf "Übertragung der Lagerdateien endet"
        Exit Sub
    End If
    
    Pause 10
    
    schreibeProtokollNachtAblauf "2. Versuch die Lagerdateien abzuholen"
    If gbFtpYes Then
        gbFTPautomatic = True
        giKissFtpMode = 20 'FTPMODE= 20,Lagerdateien holen
        frmWKL38.Show 1
        gbFTPautomatic = bmerke
    End If
    iRet = 0
'    Stehen aus irgendeinem Grund KD an dann einlesen
    schreibeProtokollNachtAblauf "Prüfung, ob Lagerdateien vorliegen"
    gbFTPautomatic = True
    iRet = newfnCheck4LagerHauptgDateien()
    gbFTPautomatic = bmerke
    If iRet <> 0 Then
        schreibeProtokollNachtAblauf "Übertragung der Lagerdateien endet"
        Exit Sub
    End If
    
    Pause 10
    
    schreibeProtokollNachtAblauf "3. Versuch die Lagerdateien abzuholen"
    If gbFtpYes Then
        gbFTPautomatic = True
        giKissFtpMode = 20 'FTPMODE= 20,Lagerdateien holen
        frmWKL38.Show 1
        gbFTPautomatic = bmerke
    End If
    iRet = 0
'    Stehen aus irgendeinem Grund KD an dann einlesen
    schreibeProtokollNachtAblauf "Prüfung, ob Lagerdateien vorliegen"
    gbFTPautomatic = True
    iRet = newfnCheck4LagerHauptgDateien()
    gbFTPautomatic = bmerke
    If iRet <> 0 Then
        schreibeProtokollNachtAblauf "Übertragung der Lagerdateien endet"
        Exit Sub
    End If
                
    schreibeProtokollNachtAblauf "Übertragung der Lagerdateien endet erfolglos"

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dieNachtverarbeitungHauptG"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub dieNachtverarbeitungLager()
On Error GoTo LOKAL_ERROR

    Dim bmerke  As Boolean
    
    bmerke = gbFTPautomatic
    
'    If gbFtpYes Then
'        gbFTPautomatic = True
'        giKissFtpMode = 1 'FTPMODE= 1
'        frmWKL38.Show 1
'        gbFTPautomatic = bmerke
'    End If
    
    schreibeProtokollNachtAblauf "externe Daten abholen"
    picprogress.Visible = True
    ExternAbholen lbl6(53), txtStatus, lbl6(28)
    picprogress.Visible = False
           
    schreibeProtokollNachtAblauf "externe Daten abholen beendet"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dieNachtverarbeitungLager"
    Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loadsprache()
    On Error GoTo LOKAL_ERROR
    
    If NewTableSuchenDBKombi("LANG", gdApp) Then
    
        Command1(0).Caption = ermLangString(giSprache, 1)
        Command1(1).Caption = ermLangString(giSprache, 2)
        Command1(2).Caption = ermLangString(giSprache, 3)
        Command1(3).Caption = ermLangString(giSprache, 4)
        Command1(8).Caption = ermLangString(giSprache, 5)
        Command1(4).Caption = ermLangString(giSprache, 6)
        Command1(5).Caption = ermLangString(giSprache, 7)
        Command1(6).Caption = ermLangString(giSprache, 8)
        Command1(7).Caption = ermLangString(giSprache, 9)

    End If
    
    
Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loadsprache"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub LadeUnternehmensDatenWKL00()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    gRegister.firma = ""
    gRegister.Plz = ""
    gRegister.Ort = ""
    
    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!name) Then
            gRegister.firma = rsrs!name
        End If
        If Not IsNull(rsrs!Plz) Then
            gRegister.Plz = rsrs!Plz
        End If
        If Not IsNull(rsrs!Ort) Then
            gRegister.Ort = rsrs!Ort
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeUnternehmensDatenWKL00"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseProgrammeinstellungenWKL00()
    On Error GoTo LOKAL_ERROR

    Dim rsrs        As Recordset
    Dim rsDB        As Recordset
    Dim rsKA        As Recordset
    Dim cSQL        As String
    Dim sSQL        As String
    Dim slastd      As String
    Dim cPfad       As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "Lochmann.cfg") Then
        gbBestinZ = True
    Else
        gbBestinZ = False
    End If
    
    If FileExists(cPfad & "Lager.cfg") Then
        LeseLagerCFG
        gbBestAkt = True
    Else
        gbBestAkt = False
    End If
    
    If FileExists(cPfad & "Hauptg.cfg") Then
        gbHauptg = True
    Else
        gbHauptg = False
    End If
    
    txtStatus.Text = 1
    
    glButtonHintergrund_from = 0
    glButtonHintergrund_to = 0
    glButtonMouseMove_Hintergrund_from = 0
    glButtonMouseMove_Hintergrund_to = 0
    glButtonMouseMove_Bordercolor = 0
    glButtonBordercolor = 0
    glButtonMouseMove_Forecolor = 0
    glButtonForecolor = 0
    
    glU1 = -2147483630
    glS1 = -2147483630
    glH1 = 12632064
    glH2 = 8421376
    glSelBack1 = 65280
    glLink = 23700
    glWarn = 23700
    gsPname = "Winkiss"
    gsFont = "Arial"
    gsFontsize = 12
    gsUpdPfad = gcDBPfad & "\In"
    
    gbErrPrint = True
    gbEtiFokEan = True
    gbFtpYes = False
    gbBEDKARTE = False
    gsDFU = "KISSHANN"
    gdTabfak = 1
    giWochendat = 0
    gbEtiQuickScanM = False
    
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        glU1 = IIf(IsNull(rsrs!U1), -2147483630, rsrs!U1)
        glS1 = IIf(IsNull(rsrs!S1), -2147483630, rsrs!S1)
        glH1 = IIf(IsNull(rsrs!H1), 12632064, rsrs!H1)
        glH2 = IIf(IsNull(rsrs!H2), 8421376, rsrs!H2)
        glSelBack1 = IIf(IsNull(rsrs!SB1), 65280, rsrs!SB1)
        glWarn = IIf(IsNull(rsrs!WARN), 23700, rsrs!WARN)
        glLink = IIf(IsNull(rsrs!LINK), 23700, rsrs!LINK)
        
        glButtonHintergrund_from = IIf(IsNull(rsrs!ButtonHintergrund_from), 0, rsrs!ButtonHintergrund_from)
        glButtonHintergrund_to = IIf(IsNull(rsrs!ButtonHintergrund_to), 0, rsrs!ButtonHintergrund_to)
        glButtonMouseMove_Hintergrund_from = IIf(IsNull(rsrs!ButtonMouseMove_Hintergrund_from), 0, rsrs!ButtonMouseMove_Hintergrund_from)
        glButtonMouseMove_Hintergrund_to = IIf(IsNull(rsrs!ButtonMouseMove_Hintergrund_to), 0, rsrs!ButtonMouseMove_Hintergrund_to)
        glButtonMouseMove_Bordercolor = IIf(IsNull(rsrs!ButtonMouseMove_Bordercolor), 0, rsrs!ButtonMouseMove_Bordercolor)
        glButtonBordercolor = IIf(IsNull(rsrs!ButtonBordercolor), 0, rsrs!ButtonBordercolor)
        glButtonMouseMove_Forecolor = IIf(IsNull(rsrs!ButtonMouseMove_Forecolor), 0, rsrs!ButtonMouseMove_Forecolor)
        glButtonForecolor = IIf(IsNull(rsrs!ButtonForecolor), 0, rsrs!ButtonForecolor)
        
        If gbDEMO Then
            gsPname = "Winkiss Demo"
        ElseIf gbKostenlos Then
            gsPname = "Winkiss free"
        Else
            gsPname = IIf(IsNull(rsrs!pname), "Winkiss", rsrs!pname)
        End If
        
        
        gsFont = IIf(IsNull(rsrs!Font), "Arial", rsrs!Font)
        gsFontsize = IIf(IsNull(rsrs!FontSize), 12, rsrs!FontSize)
        gsUpdPfad = IIf(IsNull(rsrs!UpdPfad), gcDBPfad & "\In", rsrs!UpdPfad)
        gsZinPfad = IIf(IsNull(rsrs!ZinPfad), gcDBPfad & "\Kissdata.mdb", rsrs!ZinPfad)
        gsZoutPfad = IIf(IsNull(rsrs!ZOUTPFAD), gcDBPfad & "\Kassout", rsrs!ZOUTPFAD)
        gsKinPfad = IIf(IsNull(rsrs!KinPfad), gcDBPfad & "\In", rsrs!KinPfad)
        gsDTAPfad = IIf(IsNull(rsrs!DTAPfad), gcDBPfad & "\DTAHEUTE", rsrs!DTAPfad)
        
        
        gbEtiQuickScanM = IIf(IsNull(rsrs!EtiQuickScanM), False, rsrs!EtiQuickScanM)
        
        gbEtiFokEan = IIf(IsNull(rsrs!EtiFokEan), True, rsrs!EtiFokEan)
        gbErrPrint = IIf(IsNull(rsrs!ErrPrint), True, rsrs!ErrPrint)
        gbFtpYes = IIf(IsNull(rsrs!ftp), False, rsrs!ftp)
        gsMDEGERAET = IIf(IsNull(rsrs!MDEGER), "SCANPAL", rsrs!MDEGER)
        gsWAAGE = IIf(IsNull(rsrs!Waage), "keine Waage", rsrs!Waage)
        gsDFU = IIf(IsNull(rsrs!DFU), "KISSHANN", rsrs!DFU)
        gbLocalSec = IIf(IsNull(rsrs!localsec), False, rsrs!localsec)
        gbAutoLokalModus = IIf(IsNull(rsrs!autolModus), False, rsrs!autolModus)
        gbAutoSYN = IIf(IsNull(rsrs!autosyn), False, rsrs!autosyn)
        gsZBon = IIf(IsNull(rsrs!ZBON), "", rsrs!ZBON)
        gbSichernYes = IIf(IsNull(rsrs!Sichern), False, rsrs!Sichern)
        giSICHTYP = IIf(IsNull(rsrs!SICHTYP), 0, rsrs!SICHTYP)
        gsSICHTIME = IIf(IsNull(rsrs!SICHTIME), "", rsrs!SICHTIME)
        
        gsFotoPfad = IIf(IsNull(rsrs!FotoPfad), "", rsrs!FotoPfad)
        gsWebcamPfad = IIf(IsNull(rsrs!WebcamPfad), "", rsrs!WebcamPfad)
        gsSicherPfad = IIf(IsNull(rsrs!SichPfad), gcDBPfad & "\Sicherung", rsrs!SichPfad)
        gsTankPfad = IIf(IsNull(rsrs!TankPfad), gcDBPfad & "\Box", rsrs!TankPfad)
        gsConverterPfad = IIf(IsNull(rsrs!ConverterPfad), gcDBPfad & "\Box", rsrs!ConverterPfad)
        gsWeEinzMe = IIf(IsNull(rsrs!WeEinzME), "", rsrs!WeEinzME)
        gbscanmodi = IIf(IsNull(rsrs!scanmodi), False, rsrs!scanmodi)
        giSortierung = IIf(IsNull(rsrs!etisort), 0, rsrs!etisort)
        glLokalAktuZeit = IIf(IsNull(rsrs!UPDLOKAL), 30, rsrs!UPDLOKAL)
        gsWeEinzFo = IIf(IsNull(rsrs!WeEinzFo), "LS", rsrs!WeEinzFo)
        glArtNrBeg = IIf(IsNull(rsrs!ArtNrBeg), 500000, rsrs!ArtNrBeg)
        gbEcash = IIf(IsNull(rsrs!ecash), False, rsrs!ecash)
        gsEPartner = IIf(IsNull(rsrs!EPartner), "", rsrs!EPartner)
        gbFTPautomatic = IIf(IsNull(rsrs!FTPauto), False, rsrs!FTPauto)
        gbDSL = IIf(IsNull(rsrs!FTPautoh), False, rsrs!FTPautoh)
        gbPASSIVMODE = IIf(IsNull(rsrs!passivmode), False, rsrs!passivmode)
        gbmv = IIf(IsNull(rsrs!MV), True, rsrs!MV)
        gbOptiStada = IIf(IsNull(rsrs!OptiStada), False, rsrs!OptiStada)
        gbOptiStadaSpiel = IIf(IsNull(rsrs!OptiStadaSpiel), False, rsrs!OptiStadaSpiel)
        gbBEDKARTE = IIf(IsNull(rsrs!bedkarte), False, rsrs!bedkarte)
        gbQPASS = IIf(IsNull(rsrs!QPASS), False, rsrs!QPASS)
        gbGTBON = IIf(IsNull(rsrs!GTBON), False, rsrs!GTBON)
        gbPAEBON = IIf(IsNull(rsrs!PAEBON), False, rsrs!PAEBON)
        gbYtescanPcom = IIf(IsNull(rsrs!MDECOM), 2, rsrs!MDECOM)
        giMDEPAUSE = IIf(IsNull(rsrs!MDEPAUSE), 60, rsrs!MDEPAUSE)
        gbYteWAAGEPcom = IIf(IsNull(rsrs!WAAGECOM), 2, rsrs!WAAGECOM)
        gbNacht = IIf(IsNull(rsrs!NACHT), False, rsrs!NACHT)
        gbGeld = IIf(IsNull(rsrs!Geld), True, rsrs!Geld)
        gdVerBGesrabatt = IIf(IsNull(rsrs!VerBGesrabatt), 0, rsrs!VerBGesrabatt)
        gbPBARGeld = IIf(IsNull(rsrs!PBARGeld), False, rsrs!PBARGeld)
        gbQZBON = IIf(IsNull(rsrs!QZBON), False, rsrs!QZBON)
        gbMitExport = IIf(IsNull(rsrs!MITEXPORT), False, rsrs!MITEXPORT)
        gbZBONDINA4HOCH = IIf(IsNull(rsrs!ZBONDINA4HOCH), False, rsrs!ZBONDINA4HOCH)
        gbPark = IIf(IsNull(rsrs!PARK), False, rsrs!PARK)
        gbParknetto = IIf(IsNull(rsrs!PARKnetto), False, rsrs!PARKnetto)
        gbBILDTAST = IIf(IsNull(rsrs!BILDTAST), False, rsrs!BILDTAST)
        gbBONWG = IIf(IsNull(rsrs!BONWG), False, rsrs!BONWG)
        gb2BONUSMESS = IIf(IsNull(rsrs!BONUSMESS), False, rsrs!BONUSMESS)
        gbTerminNoWarn = IIf(IsNull(rsrs!TNW), False, rsrs!TNW)
        gbAuto_Export_Artikelbestand = IIf(IsNull(rsrs!AEA), False, rsrs!AEA)
        
        gbSPY = IIf(IsNull(rsrs!SPY), False, rsrs!SPY)
        gsServerIP = IIf(IsNull(rsrs!IPADRESS), "", rsrs!IPADRESS)
        gsServerPort = IIf(IsNull(rsrs!Port), "", rsrs!Port)
        gdCheckPreis = IIf(IsNull(rsrs!CheckPreis), 0, rsrs!CheckPreis)
        gdKartenschwellenwert = IIf(IsNull(rsrs!Kartenschwellenwert), 0, rsrs!Kartenschwellenwert)
        gsPfadBestandlive = IIf(IsNull(rsrs!PfadBestandlive), "", rsrs!PfadBestandlive)
        
        
        If gbBONWG Then
            gBYTEWGNR = IIf(IsNull(rsrs!WGNR), 0, rsrs!WGNR)
        Else
            gBYTEWGNR = 0
        End If
        
        gbSPIEGEL = IIf(IsNull(rsrs!Spiegel), False, rsrs!Spiegel)
        gbKKSCHUB = IIf(IsNull(rsrs!KKSCHUB), False, rsrs!KKSCHUB)
        gbKOLSCHUB = IIf(IsNull(rsrs!KOLSCHUB), False, rsrs!KOLSCHUB)
        gbKBSCHUB = IIf(IsNull(rsrs!KBSCHUB), False, rsrs!KBSCHUB)
        gbBARZSCHUB = IIf(IsNull(rsrs!BARZSCHUB), False, rsrs!BARZSCHUB)
        gbBargeldEingabe = IIf(IsNull(rsrs!barein), False, rsrs!barein)
        gb2BONKA = IIf(IsNull(rsrs!BONKA), False, rsrs!BONKA)
        gb2BONKR = IIf(IsNull(rsrs!BONKR), True, rsrs!BONKR)
        gb2BONGUVK = IIf(IsNull(rsrs!BONGUVK), True, rsrs!BONGUVK)
        gb2BONEA = IIf(IsNull(rsrs!BONEA), True, rsrs!BONEA)
        gb2BONTermin = IIf(IsNull(rsrs!BONTERMIN), False, rsrs!BONTERMIN)
        gbWVNOT = IIf(IsNull(rsrs!WVNOT), False, rsrs!WVNOT)
        
        gb2BONKB = IIf(IsNull(rsrs!BONKB), False, rsrs!BONKB)
        gb2BONST = IIf(IsNull(rsrs!BONST), False, rsrs!BONST)
        gb2BONFI = IIf(IsNull(rsrs!BONFI), False, rsrs!BONFI)
        gb2BONVerleih = IIf(IsNull(rsrs!BONVERLEIH), False, rsrs!BONVERLEIH)
        gb2BONKOLLVK = IIf(IsNull(rsrs!BONKOLLVK), False, rsrs!BONKOLLVK)
        
        gbBONNRUNTER = IIf(IsNull(rsrs!BONNRUNTER), False, rsrs!BONNRUNTER)
        gbKASSNRUNTER = IIf(IsNull(rsrs!KASSNRUNTER), False, rsrs!KASSNRUNTER)
        gsSTERNZEICH = IIf(IsNull(rsrs!STERNZEICH), "*", rsrs!STERNZEICH)
        glZeichenAnzahlBon = IIf(IsNull(rsrs!ANZZEICHENBON), 32, rsrs!ANZZEICHENBON)
        
        gdBONFONTSIZE = IIf(IsNull(rsrs!BONFONTSIZE), 8, rsrs!BONFONTSIZE)
        gsBONFONTNAME = IIf(IsNull(rsrs!BONFONTNAME), "Standard", rsrs!BONFONTNAME)
        
        gbDritteArtikelzeile = IIf(IsNull(rsrs!DritteArtikelzeile), False, rsrs!DritteArtikelzeile)
        
        
        gdTabfak = IIf(IsNull(rsrs!Tabfak), 1, rsrs!Tabfak)
        gbBonkopie = IIf(IsNull(rsrs!BONKOPIE), True, rsrs!BONKOPIE)
        gbKSF = IIf(IsNull(rsrs!KSF), True, rsrs!KSF)
        gbAABSCHL = IIf(IsNull(rsrs!AABSCHL), False, rsrs!AABSCHL)
        gsKassDatstart = IIf(IsNull(rsrs!KASSDATSTART), "", rsrs!KASSDATSTART)
        gsTerminReminderstart = IIf(IsNull(rsrs!TerminReminderstart), "", rsrs!TerminReminderstart)
        glTageVorTermin = IIf(IsNull(rsrs!TageVorTermin), 2, rsrs!TageVorTermin)
        gbBONNEIN = IIf(IsNull(rsrs!BONNEIN), False, rsrs!BONNEIN)
        gbBARBON2 = IIf(IsNull(rsrs!BARBON2), False, rsrs!BARBON2)
        gbNOBONDRUCKER = IIf(IsNull(rsrs!NOBONDRUCKER), False, rsrs!NOBONDRUCKER)
        gbBARDINA4 = IIf(IsNull(rsrs!BARDINA4), False, rsrs!BARDINA4)
        gbDINA4VIS = IIf(IsNull(rsrs!DINA4VIS), True, rsrs!DINA4VIS)
        gbDINA4RECHFU = IIf(IsNull(rsrs!DINA4RECHFU), False, rsrs!DINA4RECHFU)
        gbDabakompautoNo = IIf(IsNull(rsrs!NOAUTO), False, rsrs!NOAUTO)
        gbGiltAlsRechnung = IIf(IsNull(rsrs!GILTRE), False, rsrs!GILTRE)
        gbEtiEan = IIf(IsNull(rsrs!ETIEAN), False, rsrs!ETIEAN)
        gbOhneAnzeige = IIf(IsNull(rsrs!OhneAnzeige), False, rsrs!OhneAnzeige)
        gsZählbeleg = IIf(IsNull(rsrs!ZAEHLBELEG), "", rsrs!ZAEHLBELEG)
        gbFtpZENT = IIf(IsNull(rsrs!FTPZENT), False, rsrs!FTPZENT)
        gsKL_ADRESSE = IIf(IsNull(rsrs!KL_ADRESSE), "", rsrs!KL_ADRESSE)
        gsKL_BENUTZER = IIf(IsNull(rsrs!KL_BENUTZER), "", rsrs!KL_BENUTZER)
        gsKL_PASSWORT = IIf(IsNull(rsrs!KL_PASSWORT), "", rsrs!KL_PASSWORT)
        gsKL_DATENBANKNAME = IIf(IsNull(rsrs!KL_DATENBANKNAME), "", rsrs!KL_DATENBANKNAME)
        gbKL_LIVEBESTAND = IIf(IsNull(rsrs!KL_LIVEBESTAND), False, rsrs!KL_LIVEBESTAND)
        gbKL_LIVEBESTAND_DIFF = IIf(IsNull(rsrs!KL_LIVEBESTAND_DIFF), False, rsrs!KL_LIVEBESTAND_DIFF)
        gbKL_LIVEKVKPR = IIf(IsNull(rsrs!KL_LIVEKVKPR), False, rsrs!KL_LIVEKVKPR)
        gbKL_LIVEGUTSCHEIN = IIf(IsNull(rsrs!KL_LIVEGUTSCHEIN), False, rsrs!KL_LIVEGUTSCHEIN)
        gbKL_LIVEFarbe = IIf(IsNull(rsrs!KL_LIVEFARBE), False, rsrs!KL_LIVEFARBE)
        gbKL_LIVEGefSperr = IIf(IsNull(rsrs!KL_LIVEGefSperr), False, rsrs!KL_LIVEGefSperr)
        gbKL_LIVENACHRICHTEN = IIf(IsNull(rsrs!KL_LIVENACHRICHTEN), False, rsrs!KL_LIVENACHRICHTEN)
        
        gbZweitMoni = IIf(IsNull(rsrs!ZWEITMONI), False, rsrs!ZWEITMONI)
        gbZweitMoniMinimieren = IIf(IsNull(rsrs!ZWEITMONIMINI), False, rsrs!ZWEITMONIMINI)
        gbSound = IIf(IsNull(rsrs!Sound), True, rsrs!Sound)
        gbISDEMO = IIf(IsNull(rsrs!isdemo), False, rsrs!isdemo)
        gbLeiste2Start = IIf(IsNull(rsrs!Leiste2Start), False, rsrs!Leiste2Start)
        gbSTADAP = IIf(IsNull(rsrs!stadap), True, rsrs!stadap)
        gbEDITKASSNR = IIf(IsNull(rsrs!EDITKASSNR), False, rsrs!EDITKASSNR)
        gbKopOhneAuswertung = IIf(IsNull(rsrs!KopOhneAuswertung), False, rsrs!KopOhneAuswertung)
        gsKL_DSN = IIf(IsNull(rsrs!KL_DSN), "", rsrs!KL_DSN)
        
        gsKaMail = IIf(IsNull(rsrs!KAMAIL), "", rsrs!KAMAIL)
        
        gbUmsAnz = IIf(IsNull(rsrs!UmsAnz), False, rsrs!UmsAnz)
        gsKassPass = IIf(IsNull(rsrs!KassPass), "", rsrs!KassPass)
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    
    Set rsrs = gdBase.OpenRecordset("WEBSHOP_E", dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst

        gsMySQL_PHP_SCRIPT_PFAD = IIf(IsNull(rsrs!MySQL_PHP_SCRIPT_PFAD), "", rsrs!MySQL_PHP_SCRIPT_PFAD)
        gbMySQL_LIVEBESTAND = IIf(IsNull(rsrs!MySQL_LIVEBESTAND), False, rsrs!MySQL_LIVEBESTAND)
        
        gsMySQL_BESTAND_TAB = IIf(IsNull(rsrs!MySQL_BESTAND_TAB), "", rsrs!MySQL_BESTAND_TAB)
        gsMySQL_BESTAND_INDEXSPALTE = IIf(IsNull(rsrs!MySQL_BESTAND_INDEXSPALTE), "", rsrs!MySQL_BESTAND_INDEXSPALTE)
        gsMySQL_BESTAND_SPALTE = IIf(IsNull(rsrs!MySQL_BESTAND_SPALTE), "", rsrs!MySQL_BESTAND_SPALTE)
    End If
    rsrs.Close: Set rsrs = Nothing
    
    txtStatus.Text = 10
    
    Dim ctmpSp As String
    ctmpSp = "(Ja: Keine weitere Lieferung mehr (MB wird auf 0 gesetzt))" & vbCrLf
    ctmpSp = ctmpSp & "(Nein: MB bleibt unverändert)" & vbCrLf
    
    Set rsKA = gdBase.OpenRecordset("KASSEIN", dbOpenTable)
    If Not rsKA.EOF Then
        gbAUSBLDU = IIf(IsNull(rsKA!AUSBLDU), False, rsKA!AUSBLDU)
        gbAUSBLSH = IIf(IsNull(rsKA!AUSBLSH), False, rsKA!AUSBLSH)
        gbAUSBLLS = IIf(IsNull(rsKA!AUSBLLS), False, rsKA!AUSBLLS)
        gbNoGrafik = IIf(IsNull(rsKA!nografik), False, rsKA!nografik)
        gbOpenSchubRetoure = IIf(IsNull(rsKA!OpenSchubRetoure), True, rsKA!OpenSchubRetoure)
        gbMBBLOCKFrage = IIf(IsNull(rsKA!mbblockfrage), True, rsKA!mbblockfrage)
        gbNeukunden = IIf(IsNull(rsKA!Neukunden), True, rsKA!Neukunden)
        gbKASSMBEST = IIf(IsNull(rsKA!KASSMBEST), False, rsKA!KASSMBEST)
        gbEKMAX = IIf(IsNull(rsKA!EKMAX), True, rsKA!EKMAX)
        gbTPbf = IIf(IsNull(rsKA!TPbf), False, rsKA!TPbf)
        gbSterne = IIf(IsNull(rsKA!Sterne), True, rsKA!Sterne)
        gbNoBonGu = IIf(IsNull(rsKA!nobongu), False, rsKA!nobongu)
        gbBonGu2J = IIf(IsNull(rsKA!BonGu2J), False, rsKA!BonGu2J)
        
        
        gbSonderPreisDarstellen = IIf(IsNull(rsKA!SonderPreisDarstellen), False, rsKA!SonderPreisDarstellen)
        gbNoBonPÄ = IIf(IsNull(rsKA!nobonPAE), False, rsKA!nobonPAE)
        gdRESTGU = IIf(IsNull(rsKA!RESTGU), 0, rsKA!RESTGU)
        giFILALI = IIf(IsNull(rsKA!FILALI), 0, rsKA!FILALI)
        gbPrintLOGO = IIf(IsNull(rsKA!PL), True, rsKA!PL)
        gbKUDU = IIf(IsNull(rsKA!KUDU), False, rsKA!KUDU)
        
        glLieblingArtnr = IIf(IsNull(rsKA!LieblingArtnr), 0, rsKA!LieblingArtnr)
        glLieblingAR = IIf(IsNull(rsKA!LieblingAR), 0, rsKA!LieblingAR)
        
        gdWarenkorbWert = IIf(IsNull(rsKA!WarenkorbWert), 0, rsKA!WarenkorbWert)
        gdWarenkorbGR = IIf(IsNull(rsKA!WarenkorbGR), 0, rsKA!WarenkorbGR)
        
        glBaganzArtnr = IIf(IsNull(rsKA!BaganzArtnr), 0, rsKA!BaganzArtnr)
        glBaganzAR = IIf(IsNull(rsKA!BaganzAR), 0, rsKA!BaganzAR)
        
        glZehnProzLinr = IIf(IsNull(rsKA!ZehnProzLinr), 0, rsKA!ZehnProzLinr)
        glZehnProzArtnr = IIf(IsNull(rsKA!ZehnProzArtnr), 0, rsKA!ZehnProzArtnr)
        
        gbGEBRABK = IIf(IsNull(rsKA!GEBRABK), False, rsKA!GEBRABK)
        
        gbBestDateien = IIf(IsNull(rsKA!BestDateien), False, rsKA!BestDateien)
        gbOhnebestProt = IIf(IsNull(rsKA!OhnebestProt), False, rsKA!OhnebestProt)
        gbKeineBestVerWarengru = IIf(IsNull(rsKA!KeineBestVerWarengru), False, rsKA!KeineBestVerWarengru)
        
        gbBarAnz = IIf(IsNull(rsKA!BarAnz), False, rsKA!BarAnz)
        gbEinfacheZollErstattung = IIf(IsNull(rsKA!EinfacheZollErstattung), False, rsKA!EinfacheZollErstattung)
'        gbUmsAnz = IIf(IsNull(rsKA!UmsAnz), False, rsKA!UmsAnz)
        
        gbKBmBI = IIf(IsNull(rsKA!KBmBI), True, rsKA!KBmBI)
        gbmGDetails = IIf(IsNull(rsKA!mGDetails), False, rsKA!mGDetails)
        gbArtrabhalten = IIf(IsNull(rsKA!artrabh), False, rsKA!artrabh)
        gbBEDLEER = IIf(IsNull(rsKA!BEDLEER), False, rsKA!BEDLEER)
        gbBONWAHL = IIf(IsNull(rsKA!BONwahl), False, rsKA!BONwahl)
        gbkassgefuehrt = IIf(IsNull(rsKA!kassgefuehrt), False, rsKA!kassgefuehrt)
        gbMitPreis = IIf(IsNull(rsKA!MITPREIS), False, rsKA!MITPREIS)
        gbSparsatz = IIf(IsNull(rsKA!SparSatz), False, rsKA!SparSatz)
        gbRETVK = IIf(IsNull(rsKA!RETVK), False, rsKA!RETVK)
        gbMitMwstAnteile = IIf(IsNull(rsKA!MitMwstAnteile), False, rsKA!MitMwstAnteile)
        gbRabVs = IIf(IsNull(rsKA!RABVS), False, rsKA!RABVS)
        gbRabatt = IIf(IsNull(rsKA!RABATT), True, rsKA!RABATT)
        gbIdentUser = IIf(IsNull(rsKA!IdentUser), False, rsKA!IdentUser)
        gbNachKBbeiEC = IIf(IsNull(rsKA!NachKBbeiEC), False, rsKA!NachKBbeiEC)
        gbKUWAHLROT = IIf(IsNull(rsKA!KUWAHLROT), False, rsKA!KUWAHLROT)
        
        gbKUWAHLGESPERRTROT = IIf(IsNull(rsKA!KUWAHLGESPERRTROT), True, rsKA!KUWAHLGESPERRTROT)
        
        
        
        
        gbKUWAHLfbimmer = IIf(IsNull(rsKA!KUWAHLfbimmer), False, rsKA!KUWAHLfbimmer)
        gbKUWAHLMAIL = IIf(IsNull(rsKA!KUWAHLMAIL), False, rsKA!KUWAHLMAIL)
        gbKUBONUS = IIf(IsNull(rsKA!KUBONUS), True, rsKA!KUBONUS)
        gbArtikelTextSuche = IIf(IsNull(rsKA!AUTOSEEK), True, rsKA!AUTOSEEK)
        gsZOLLARTBEZ = IIf(IsNull(rsKA!ZOLLARTBEZ), "Parfümerieartikel", rsKA!ZOLLARTBEZ)
        gbCoupon = IIf(IsNull(rsKA!COUPON), False, rsKA!COUPON)
        gbGuStattBar = IIf(IsNull(rsKA!GuStattBar), False, rsKA!GuStattBar)
        gbMitStaffelPreis = IIf(IsNull(rsKA!MitStaffelPreis), True, rsKA!MitStaffelPreis)
        gbSaveReport = IIf(IsNull(rsKA!SAVEREPORT), False, rsKA!SAVEREPORT)
        gbDSDRUCKEN = IIf(IsNull(rsKA!DSDRUCKEN), False, rsKA!DSDRUCKEN)
        gbDS_GEB_DRUCKEN = IIf(IsNull(rsKA!DS_GEB_DRUCKEN), True, rsKA!DS_GEB_DRUCKEN)
        gbDSMeldungErfolg = IIf(IsNull(rsKA!DSMeldungErfolg), True, rsKA!DSMeldungErfolg)
        gbDSKLEIN = IIf(IsNull(rsKA!DSKLEIN), False, rsKA!DSKLEIN)
        gbPLZGEBIET = IIf(IsNull(rsKA!PLZGEBIET), False, rsKA!PLZGEBIET)
        gbMitKundeWahlHinweis = IIf(IsNull(rsKA!MitKundeWahlHinweis), False, rsKA!MitKundeWahlHinweis)
        gbPLZGEBIET_AuchBeiKUWAHL = IIf(IsNull(rsKA!PLZGEBIET_AuchBeiKUWAHL), False, rsKA!PLZGEBIET_AuchBeiKUWAHL)
        gbZOLLmMWST = IIf(IsNull(rsKA!ZOLLmMWST), False, rsKA!ZOLLmMWST)
        gbZOLLonlyFirstPage = IIf(IsNull(rsKA!ZOLLonlyFirstPage), False, rsKA!ZOLLonlyFirstPage)
        gbZOLLPrintDirekt = IIf(IsNull(rsKA!ZOLLPrintDirekt), False, rsKA!ZOLLPrintDirekt)
        gbHandelsspanne_Ausblenden = IIf(IsNull(rsKA!Handelsspanne_Ausblenden), False, rsKA!Handelsspanne_Ausblenden)
        gbAlterGutschein_Ausblenden = IIf(IsNull(rsKA!AlterGutschein_Ausblenden), False, rsKA!AlterGutschein_Ausblenden)
        gbKUBONUS_WENN = IIf(IsNull(rsKA!KUBONUS_WENN), True, rsKA!KUBONUS_WENN)
        gsiKUBONUS_SCHWELLE = IIf(IsNull(rsKA!KUBONUS_SCHWELLE), 0, rsKA!KUBONUS_SCHWELLE)
        gbNoKUBONUS_wenn_Art_and_Ges_rab = IIf(IsNull(rsKA!NOKUBONUS_AGRAB), False, rsKA!NOKUBONUS_AGRAB)
        
        gsVorEinPLZ1 = IIf(IsNull(rsKA!VorEinPLZ1), "PLZ 1", rsKA!VorEinPLZ1)
        gsVorEinPLZ2 = IIf(IsNull(rsKA!VorEinPLZ2), "PLZ 2", rsKA!VorEinPLZ2)
        glRRArtnr = IIf(IsNull(rsKA!rrartnr), 0, rsKA!rrartnr)
        glBonusGrenzeArtnr = IIf(IsNull(rsKA!BonusGrenzeArtnr), 0, rsKA!BonusGrenzeArtnr)
        
        glBonusAuszahlungArtnr = IIf(IsNull(rsKA!BonusAuszahlungArtnr), 0, rsKA!BonusAuszahlungArtnr)
        
        glAutoKundnrforKundBest = IIf(IsNull(rsKA!AutoKundnrforKundBest), 0, rsKA!AutoKundnrforKundBest)
        glAutoAusSchFiliale = IIf(IsNull(rsKA!AutoAusSchFiliale), 0, rsKA!AutoAusSchFiliale)
        
        
        gdSCHWELLEWK = IIf(IsNull(rsKA!SCHWELLEWK), 0, rsKA!SCHWELLEWK)
        gbNurBonusfRunden = IIf(IsNull(rsKA!NurBonusfRunden), False, rsKA!NurBonusfRunden)
        gbNachfragenbeiWGNohnePreis = IIf(IsNull(rsKA!NachfragenbeiWGNohnePreis), False, rsKA!NachfragenbeiWGNohnePreis)
        gsAbrunden = IIf(IsNull(rsKA!Runden), "1", rsKA!Runden)
        gsFARBKASSE = IIf(IsNull(rsKA!FARBKASSE), "1", rsKA!FARBKASSE)
        gsECBILD = IIf(IsNull(rsKA!ECBILD), "1", rsKA!ECBILD)
        glGSArtnr = IIf(IsNull(rsKA!gsARTNR), 0, rsKA!gsARTNR)
        giFarbebeiPark = IIf(IsNull(rsKA!FarbebeiPark), 0, rsKA!FarbebeiPark)
        gdZeitungsSpanne = IIf(IsNull(rsKA!ZeitungsSpanne), 0, rsKA!ZeitungsSpanne)
        glPrimLinr = IIf(IsNull(rsKA!PrimLinr), 0, rsKA!PrimLinr)
        glZeitungsLinr = IIf(IsNull(rsKA!ZeitungsLinr), 0, rsKA!ZeitungsLinr)
        glPaketLinr = IIf(IsNull(rsKA!PaketLinr), 0, rsKA!PaketLinr)
        gsSpezArtikel = IIf(IsNull(rsKA!SpezArtikel), "", rsKA!SpezArtikel)
        gsRabattAusnahmeArtikel = IIf(IsNull(rsKA!RabattAusnahmeArtikel), "", rsKA!RabattAusnahmeArtikel)
        glSpezFotoartikel = IIf(IsNull(rsKA!SpezFotoartikel), 0, rsKA!SpezFotoartikel)
        glSpezLottoauszahlartikel = IIf(IsNull(rsKA!SpezLottoauszahlartikel), 0, rsKA!SpezLottoauszahlartikel)
        glECAuszahlArtnr = IIf(IsNull(rsKA!ECAuszahlArtnr), 0, rsKA!ECAuszahlArtnr)
        
        gskPW = IIf(IsNull(rsKA!kPW), "ß", rsKA!kPW)
        
        gsSpezBontext = IIf(IsNull(rsKA!Spezbontext), "", rsKA!Spezbontext)
        gsSpezBontext2 = IIf(IsNull(rsKA!Spezbontext2), "", rsKA!Spezbontext2)
        gsSpezBontext3 = IIf(IsNull(rsKA!Spezbontext3), "", rsKA!Spezbontext3)
        gsSpezBontextU = IIf(IsNull(rsKA!SpezbontextU), "", rsKA!SpezbontextU)
        gsSpezBonArtRab = IIf(IsNull(rsKA!SpezBonArtRab), "10", rsKA!SpezBonArtRab)
       
        gsSperrFrage = IIf(IsNull(rsKA!SperrFrage), ctmpSp, rsKA!SperrFrage)
        giBARGELDART = IIf(IsNull(rsKA!BARGELDART), 0, rsKA!BARGELDART)
        gbRESTinBAR = IIf(IsNull(rsKA!RESTinBAR), True, rsKA!RESTinBAR)
        
        gbKK_Visa = IIf(IsNull(rsKA!KK_Visa), True, rsKA!KK_Visa)
        gbKK_EurocardMastercard = IIf(IsNull(rsKA!KK_EurocardMastercard), True, rsKA!KK_EurocardMastercard)
        gbKK_AmericanExpress = IIf(IsNull(rsKA!KK_AmericanExpress), True, rsKA!KK_AmericanExpress)
        gbKK_DinersClub = IIf(IsNull(rsKA!KK_DinersClub), True, rsKA!KK_DinersClub)
        gbKK_ECKarte = IIf(IsNull(rsKA!KK_ECKarte), True, rsKA!KK_ECKarte)
        gbKK_Sonstige = IIf(IsNull(rsKA!KK_Sonstige), True, rsKA!KK_Sonstige)
        
        gbKK_AliPay = IIf(IsNull(rsKA!KK_AliPay), True, rsKA!KK_AliPay)
        gbKK_ApplePay = IIf(IsNull(rsKA!KK_ApplePay), True, rsKA!KK_ApplePay)
        gbKK_GooglePay = IIf(IsNull(rsKA!KK_GooglePay), True, rsKA!KK_GooglePay)
        
        
        gbKK_PayPal = IIf(IsNull(rsKA!KK_PayPal), True, rsKA!KK_PayPal)
        gbKK_YabandPay = IIf(IsNull(rsKA!KK_YabandPay), True, rsKA!KK_YabandPay)
        
        gbArtsucheArtFarb = IIf(IsNull(rsKA!ArtsucheArtFarb), True, rsKA!ArtsucheArtFarb)
        
        gcSMTP_SERVER = IIf(IsNull(rsKA!SMTP_SERVER), "smtp.strato.de", rsKA!SMTP_SERVER)
        gcSMTP_USER = IIf(IsNull(rsKA!SMTP_USER), "bestsend@kisswws.de", rsKA!SMTP_USER)
        gcSMTP_PW = IIf(IsNull(rsKA!SMTP_PW), "Ki55!Ww52020", rsKA!SMTP_PW)
        gcSMTP_PORT = IIf(IsNull(rsKA!SMTP_PORT), "587", rsKA!SMTP_PORT)
        
        gbSMTP_SSL = IIf(IsNull(rsKA!SMTP_SSL), True, rsKA!SMTP_SSL)
        
        If (gcSMTP_SERVER = "smtp.strato.de") And (gcSMTP_USER = "bestsend@kisswws.de") And (gcSMTP_PW = "geheim") Then
            gcSMTP_PW = "Ki55!Ww52020"
        End If
        
        
    Else
    
    
    
    
        gcSMTP_SERVER = "smtp.strato.de"
        gcSMTP_USER = "bestsend@kisswws.de"
        gcSMTP_PW = "Ki55!Ww52020"
        gcSMTP_PORT = "587"
        gbSMTP_SSL = True
        
        gbArtsucheArtFarb = True
        gbPLZGEBIET_AuchBeiKUWAHL = False
        gbRESTinBAR = True
        gsAbrunden = "1"
        gsFARBKASSE = "1"
        gsRabattAusnahmeArtikel = ""
        gsSpezArtikel = ""
        gsSperrFrage = ctmpSp
        giFarbebeiPark = 0
        glPrimLinr = 0
        glZeitungsLinr = 0
        glRRArtnr = 0
        glBonusGrenzeArtnr = 0
        glBonusAuszahlungArtnr = 0
        glAutoKundnrforKundBest = 0
        glAutoAusSchFiliale = 0
        gdSCHWELLEWK = 0
        gbNurBonusfRunden = False
        gbNachfragenbeiWGNohnePreis = False
        glGSArtnr = 0
        gsVorEinPLZ1 = "PLZ 1"
        gsVorEinPLZ2 = "PLZ 2"
        gbHandelsspanne_Ausblenden = False
        gbAlterGutschein_Ausblenden = False
        gsZOLLARTBEZ = "Parfümerieartikel"
        gbKUBONUS = True
        gbNoKUBONUS_wenn_Art_and_Ges_rab = False
        gbKUWAHLMAIL = False
        gbKUWAHLROT = False
        gbKUWAHLfbimmer = False
        gbNoGrafik = False
        gbNoBonGu = False
        gbAUSBLDU = False
        gbAUSBLSH = False
        gbAUSBLLS = False
        gdRESTGU = 0
        giFILALI = 0
        gbKUDU = False
        gbZweitMoni = False
        gbGEBRABK = False
        gbOhnebestProt = False
        gbKeineBestVerWarengru = False
        gbBestDateien = False
        gbBarAnz = False
        gbEinfacheZollErstattung = False
        gbUmsAnz = False
        gbEDITKASSNR = False
        gbKBmBI = True
        gbmGDetails = False
        gbArtrabhalten = False
        gbBEDLEER = False
        gbBONNEIN = False
        gbBONWAHL = False
        gbkassgefuehrt = False
        gbMitPreis = False
        gbSparsatz = False
        gbRabVs = False
        gbRabatt = True
        gbLeiste2Start = False
        gbNachKBbeiEC = False
        gbOpenSchubRetoure = True
        gbArtikelTextSuche = True
        gbPLZGEBIET = False
        gbGuStattBar = False
        gbMitStaffelPreis = True
        gdZeitungsSpanne = 0
        giBARGELDART = 0
        
    End If
    rsKA.Close
    
    txtStatus.Text = 19
    
    Set rsDB = gdBase.OpenRecordset("DBEINSTE", dbOpenTable)
        gsSpanne = IIf(IsNull(rsDB!SPANNE), "LEK", rsDB!SPANNE)
        gdBonusGrenze = IIf(IsNull(rsDB!BONUSGRENZ), 0, rsDB!BONUSGRENZ)
        gdBonusGutscheinBeiGrenze = IIf(IsNull(rsDB!BonusGutscheinBeiGrenze), 0, rsDB!BonusGutscheinBeiGrenze)
        giAufrunden = IIf(IsNull(rsDB!Aufrunden), 0, rsDB!Aufrunden)
        giAbrunden = IIf(IsNull(rsDB!ABRUNDEN), 0, rsDB!ABRUNDEN)
        giRundkrit = IIf(IsNull(rsDB!Rundkrit), 0, rsDB!Rundkrit)
        gbGutsch = IIf(IsNull(rsDB!HaGuNr), False, rsDB!HaGuNr)
        gsMWST = IIf(IsNull(rsDB!mwstbeg), "V", rsDB!mwstbeg)
'        gbSTADAP = IIf(IsNull(rsDB!stadap), True, rsDB!stadap)
        gbFTH = IIf(IsNull(rsDB!FTH), True, rsDB!FTH)
        gbSondRab = IIf(IsNull(rsDB!SondRab), False, rsDB!SondRab)
'        gsKassPass = IIf(IsNull(rsDB!KassPass), "", rsDB!KassPass)
        glUPDCOUNT = IIf(IsNull(rsDB!updcount), 1000, rsDB!updcount)
        glUPDTime = IIf(IsNull(rsDB!updTIME), 100, rsDB!updTIME)
        gbKUNDENA = IIf(IsNull(rsDB!KUIMBOY), True, rsDB!KUIMBOY)
        gdDBPAUSE = IIf(IsNull(rsDB!DBPAUSE), 0, rsDB!DBPAUSE)
        gbFilKasDat = IIf(IsNull(rsDB!FilKasDat), False, rsDB!FilKasDat)
        
        gbUnistatWeek = IIf(IsNull(rsDB!UstatW), False, rsDB!UstatW)
        gbUnistatMonat = IIf(IsNull(rsDB!UstatM), False, rsDB!UstatM)
        gbAbschlussNummer = IIf(IsNull(rsDB!abnr), False, rsDB!abnr)
        gbAbschlussDatum = IIf(IsNull(rsDB!abda), False, rsDB!abda)
        gbAGNAusw = IIf(IsNull(rsDB!AGNAUSW), False, rsDB!AGNAUSW)
        gbARTKUMWGN = IIf(IsNull(rsDB!ARTKUMWGN), False, rsDB!ARTKUMWGN)
        gbKUMSUM = IIf(IsNull(rsDB!KUMSUM), True, rsDB!KUMSUM)
        gbDabakompfrueh = IIf(IsNull(rsDB!Storno), True, rsDB!Storno)
        gbPenner_faerben = IIf(IsNull(rsDB!PENNERFARB), True, rsDB!PENNERFARB)
        gbNewArt = IIf(IsNull(rsDB!NewArt), True, rsDB!NewArt)
        gbNewArtNrVorschlag = IIf(IsNull(rsDB!NewArtNrVorschlag), True, rsDB!NewArtNrVorschlag)
        gbDruck27 = IIf(IsNull(rsDB!Druck27), True, rsDB!Druck27)
        gbFILMEK = IIf(IsNull(rsDB!FILMEK), True, rsDB!FILMEK)
        gbFILBONI = IIf(IsNull(rsDB!FILBONI), True, rsDB!FILBONI)
        gbECTOZ = IIf(IsNull(rsDB!ECTOZ), True, rsDB!ECTOZ)
        gbBonusBNB = IIf(IsNull(rsDB!BonusBNB), True, rsDB!BonusBNB)
        gbAA = IIf(IsNull(rsDB!aa), False, rsDB!aa)
        gbTagAkt = IIf(IsNull(rsDB!tagakt), False, rsDB!tagakt)
        gbOGV = IIf(IsNull(rsDB!OGV), False, rsDB!OGV)
        gbRGO = IIf(IsNull(rsDB!RGO), False, rsDB!RGO)
        gbGutschnrKomplett = IIf(IsNull(rsDB!GutschnrKomplett), False, rsDB!GutschnrKomplett)
        gbGutscheinBeiVKversteuern = IIf(IsNull(rsDB!GutscheinBeiVKversteuern), False, rsDB!GutscheinBeiVKversteuern)
        gbREGEB = IIf(IsNull(rsDB!REGEB), False, rsDB!REGEB)
        gbTerminReminderSMS = IIf(IsNull(rsDB!TerminReminderSMS), False, rsDB!TerminReminderSMS)
        gbSPEZRU = IIf(IsNull(rsDB!SPEZRU), False, rsDB!SPEZRU)
        gbSPEZVAR = IIf(IsNull(rsDB!SPEZVAR), 0, rsDB!SPEZVAR)
        gbSCHUBMB = IIf(IsNull(rsDB!SCHUBMB), False, rsDB!SCHUBMB)
        gbGUTSCHBARAUSZAHLUNGMITUNTER = IIf(IsNull(rsDB!GUTSCHBARAUSZAHLUNGMITUNTER), False, rsDB!GUTSCHBARAUSZAHLUNGMITUNTER)
        gbKurzerStorni = IIf(IsNull(rsDB!KurzerStorni), False, rsDB!KurzerStorni)
        gbSTORNOcheck2Bed = IIf(IsNull(rsDB!STORNOcheck2Bed), False, rsDB!STORNOcheck2Bed)
        gbOLDSTADADEL = IIf(IsNull(rsDB!OLDSTADADEL), True, rsDB!OLDSTADADEL)
        gdStadaPause = IIf(IsNull(rsDB!StadaPause), 0, rsDB!StadaPause)
        gbDELBDAT = IIf(IsNull(rsDB!DELBDAT), False, rsDB!DELBDAT)
        gbDIFFPROT = IIf(IsNull(rsDB!DIFFPROT), False, rsDB!DIFFPROT)
        gbUEBERPROT = IIf(IsNull(rsDB!UEBERPROT), False, rsDB!UEBERPROT)
        gsiGESRAB = IIf(IsNull(rsDB!GESRAB), 0, rsDB!GESRAB) 'Gesamtrabatt an Kasse voreingestellt
        gsGESRABBEZ = IIf(IsNull(rsDB!GESRABBEZ), "Jubiläumsrabatt:", rsDB!GESRABBEZ) '
        gsGZBez = IIf(IsNull(rsDB!GZBez), "", rsDB!GZBez) '
        gbKundRabattDeaktiv = IIf(IsNull(rsDB!KundRabattDeaktiv), False, rsDB!KundRabattDeaktiv)
        gbyLugBe = IIf(IsNull(rsDB!LugBe), 2, rsDB!LugBe)
        giTageZugang = IIf(IsNull(rsDB!LugTAGZ), 365, rsDB!LugTAGZ)
        giTageVerkauf = IIf(IsNull(rsDB!LugTAGV), 365, rsDB!LugTAGV)
        gsKUPFAD = IIf(IsNull(rsDB!KUPFAD), "", rsDB!KUPFAD)
        gbGesEKWert_anzeigen = IIf(IsNull(rsDB!GESEK), True, rsDB!GESEK)
        giGebTage = IIf(IsNull(rsDB!GEBTAGE), 2, rsDB!GEBTAGE)
        
        If gbBEDKARTE = False Then
            gbBEDKARTE = IIf(IsNull(rsDB!bedkarte), False, rsDB!bedkarte)
        End If
    
        gbZugriffNew = IIf(IsNull(rsDB!ZugriffNew), False, rsDB!ZugriffNew)
        gbNoSpruch = IIf(IsNull(rsDB!NoSpruch), False, rsDB!NoSpruch)
        gbWEautoGef = IIf(IsNull(rsDB!WEautoGef), True, rsDB!WEautoGef)
        gbNONEGZU = IIf(IsNull(rsDB!NONEGZU), False, rsDB!NONEGZU)
        gbAutoZwsp = IIf(IsNull(rsDB!AutoZwsp), False, rsDB!AutoZwsp)
        gbETIKVKAE = IIf(IsNull(rsDB!ETIKVKAE), False, rsDB!ETIKVKAE)
        gbArtEindeut = IIf(IsNull(rsDB!ARTEINDEUT), True, rsDB!ARTEINDEUT)
        gsEdeka = IIf(IsNull(rsDB!Edeka), "", rsDB!Edeka)
        gbGebAdresse = IIf(IsNull(rsDB!GebAdresse), False, rsDB!GebAdresse)
        gbETIONLYME = IIf(IsNull(rsDB!ETIONLYME), False, rsDB!ETIONLYME)
        gbNoETIWeAusBe = IIf(IsNull(rsDB!NoETIWeAusBe), False, rsDB!NoETIWeAusBe)
        gbKVKSicher = IIf(IsNull(rsDB!KVKSicher), False, rsDB!KVKSicher)
        gbNOWOCHENDATEN = IIf(IsNull(rsDB!NOWOCHENDATEN), False, rsDB!NOWOCHENDATEN)
        gsJUGENDSCHUTZFARBE = IIf(IsNull(rsDB!JUGENDSCHUTZFARBE), "", rsDB!JUGENDSCHUTZFARBE)
        gsUnbekanntStrichMail = IIf(IsNull(rsDB!UnbekanntStrichMail), "", rsDB!UnbekanntStrichMail)
        gsNachtVerarbeitungMail = IIf(IsNull(rsDB!NachtVerarbeitungMail), "", rsDB!NachtVerarbeitungMail)
        
        gbETIBEIFARB = IIf(IsNull(rsDB!ETIBEIFARB), False, rsDB!ETIBEIFARB)
        gsDabaNachtStart = IIf(IsNull(rsDB!NachtStart), "", rsDB!NachtStart)
        gbJBTART = IIf(IsNull(rsDB!JBTART), False, rsDB!JBTART)
        
    rsDB.Close

    txtStatus.Text = 28
    
    If gbEcash = False Then gsEPartner = ""
    
    If gbKUNDENA = True Then
        leseKundenimBon
    Else
        gbKUIBONname = False
        gbKUIBONvorname = False
        gbKUIBONtitel = False
        gbKUIBONfirma = False
        gbKUIBONplz = False
        gbKUIBONort = False
        gbKUIBONstrasse = False
        gbKUIBONtel = False
        gbKUIBONmobil = False
    End If

    If gbNacht = True Then
        leseNacht
    Else
        gbPCAus = False
        gbUKDAT = False
        gbEKDAT = False
        gbUSTAT = False
        gbUPRO = False
        gbEXTSICH = False
        gbUSTADA = False
        gsNachtstart = ""
    End If
    
    txtStatus.Text = 41
    
    'IPSTAT auslesen
    If NewTableSuchenDBKombi("IPSTAT", gdBase) Then
        If CBool(leseIPstat("live")) = True Then
            gbIPSTAT = True
            gsIPMarktNr = leseIPstat("marktnr")
        End If
    End If
    
    'VEDESSTAT auslesen
    If NewTableSuchenDBKombi("VEDESSTAT", gdBase) Then
        If CBool(leseVEDESstat("live")) = True Then
            gbVEDESSTAT = True
            gsVEDESMarktNr = leseVEDESstat("marktnr")
        End If
    End If
    
    
    txtStatus.Text = 45
    
    leseZeitSteu
    
    gbKONTIN = False
    If Datendrin("KONTIN", gdBase) Then
        gbKONTIN = True
    End If
    If gbKUDU Then
        If Day(DateValue(Now)) = 17 Then
            loeschNEW "KUTITEL", gdBase
            loeschNEW "KUSTADT", gdBase
            loeschNEW "KUPLZ", gdBase
        End If
    
        If NewTableSuchenDBKombi("KUTITEL", gdBase) = False Then
            fülledistinctTabelle "TITEL", "", "KUTITEL", "KUNDEN"
        End If
        txtStatus.Text = 48
        If NewTableSuchenDBKombi("KUSTADT", gdBase) = False Then
            fülledistinctTabelle "STADT", "PLZ", "KUSTADT", "KUNDEN"
        End If
        txtStatus.Text = 51
        If NewTableSuchenDBKombi("KUPLZ", gdBase) = False Then
            fülledistinctTabelle "PLZ", "STADT", "KUPLZ", "KUNDEN"
        End If
        txtStatus.Text = 58
    End If
    
    If gbQZBON = True Then
        leseABREport
    Else
        gbTAGFILT = False
        gbARTKUM = False
        gbKK = False
        gbEA = False
        gbMitExport = False
    End If
    
    txtStatus.Text = 62
    If gbPrintLOGO = True Then
        leseLOGOS
    Else
        gbLOGO1 = False
        gbLOGO2 = False
        gbLOGO3 = False
    End If

    txtStatus.Text = 67

    If gbFtpYes Then    'Detailinfos FTP
        If NewTableSuchenDBKombi("StammFTP", gdBase) Then
            LeseStammFtp
        End If
    Else
        gbFTPautomatic = False
'        gbDSL = False
    End If
    
    txtStatus.Text = 74
    If gBYTEWGNR > 0 Then
        gsWGart = ermartnrausWGN(CStr(gBYTEWGNR))
        gsWGBEZEICH = ermBezeichausWGN(gsWGart)
    Else
        gsWGart = ""
        gsWGBEZEICH = ""
    End If
    
    txtStatus.Text = 77
    
    If gbUnistatWeek Then    'Detailinfos Statistik Uniformate Wochenauswertung
        If tableSuchenDBKombi("Statist", 1) Then
            LeseStatistWoche
        End If
    End If
    
    If gbUnistatMonat Then    'Detailinfos Statistik Uniformate Monatsauswertung
        If tableSuchenDBKombi("Statist", 1) Then
            LeseStatistMonat
        End If
    End If
    
    txtStatus.Text = 79
    
    If gbSichernYes Then    'Detailinfos Sicherung
        If NewTableSuchenDBKombi("SICHERUNG", gdBase) = False Then
            CreateTableT2 "SICHERUNG", gdBase
            
            sSQL = "Insert into Sicherung (Lastdate) values ('" & DateValue(Now) - 1 & "')"
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        slastd = LeselastSicherung
        
        If Format$(slastd, "DD.MM.YY") = Format$(DateValue(Now), "DD.MM.YY") Then
            gbSichernHeut = False
        Else
            gbSichernHeut = True
        End If
    Else
        gbSichernHeut = False
    End If
    
    If Not gbLocalSec Then
        Kill "c:\aleer\kissdata.mdb"
    End If
    
    txtStatus.Text = 87
    
    DBEngine.SetOption dbLockRetry, glUPDCOUNT
    DBEngine.SetOption dbLockDelay, glUPDTime
'    DBEngine.SetOption dbPageTimeout, 100
    
    gsZoutPfad = ShortPath(gsZoutPfad)
    gsUpdPfad = ShortPath(gsUpdPfad)
    gsKinPfad = ShortPath(gsKinPfad)
    
    txtStatus.Text = 99
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 3265 Or err.Number = 75 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "LeseProgrammeinstellungenWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
         
    End If
   
End Sub
Private Sub LeseDatenVerbindung()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As String
    
    iFileNr = FreeFile
    Open gcPfad & "VERBIND.TXT" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        Get #iFileNr, 1, gVerbindung
    End If
    Close iFileNr
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseProgrammeinstellungenWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    MsgBox "frmWKL00.LeseDatenVerbindung: " & err.Number & " / " & err.Number
End Sub
Private Sub zwangsoptimierung()
    On Error GoTo LOKAL_ERROR
    
    Dim lLastKompDate As Long
    Dim lHeute  As Long
    Dim sSQL As String
    
    lHeute = Fix(Now)
    lLastKompDate = DateValue(DatumLastKompAnzeigen)
    If lLastKompDate < (lHeute - 7) Then

        If BistDualleineinderDatenbank Then
                
             lbl6(28).ForeColor = vbRed
             lbl6(28).Visible = True
             lbl6(28).Caption = "Heute ist der Zeitpunkt erreicht, um die Datenbank zu komprimieren."
             lbl6(28).Refresh

             lbl6(53).ForeColor = vbRed
             lbl6(53).Caption = "Nehmen Sie sich noch  ein paar Minuten Zeit!"
             lbl6(53).Refresh
             
            Pause (2)
             
            lbl6(28).ForeColor = vbRed
            lbl6(28).Caption = "Im Falle eines Abbruchs werden Uhrzeit und Bediener protokolliert."
            lbl6(28).Refresh
             
            lbl6(53).ForeColor = vbRed
            lbl6(53).Visible = True
            lbl6(53).Caption = "Schalten Sie den Rechner NICHT aus!!!"
            lbl6(53).Refresh
             
            picprogress.Visible = True
            lbl6(28).Visible = True
            lbl6(53).Visible = True
             
            CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige

            sSQL = "update dbeinste set lastkomp='" & Date & "'"
            sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            picprogress.Visible = False
            lbl6(28).Visible = True
            lbl6(53).Visible = True
             
         End If
    End If

    'Zwangsoptimierung Ende
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zwangsoptimierung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    Dim lWert   As Long
    Dim cPfad   As String
    Dim lRet    As Long
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Update"
    
    Select Case index
        Case Is = 0
            If Label1(2).Visible = True Then
                Label1(2).Visible = False
            Else
                If Not gbDEMO And Not gbKostenlos Then
                    ctmp = FileDateTime(gcPfad & "WINKISS.EXE")
                    Label1(2).Caption = "Stand: " & ctmp
                    Label1(2).Visible = True
                End If
            End If
        Case Is = 4
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/neuigkeiten.html"
        Case Is = 5 'Video Grundkurs
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/video-schulung.html"
        Case Is = 6
        
            Screen.MousePointer = 11
            URLGoTo Me.hwnd, "http://kisslive.de/winkiss/winkiss-beitraege/266-tse.html"
            
''            lRet = Shell(App.Path & "\TV_wk.exe", vbNormalFocus)
            
            Screen.MousePointer = 0
            
        Case Is = 10
        
            Screen.MousePointer = 11
            lRet = Shell(App.Path & "\pcvisit_KISS.c.F0312623622.exe", vbNormalFocus)
            
            Screen.MousePointer = 0
'            URLGoTo Me.hwnd, "http://www.kisslive.de/downloads/winkiss/Thomas_Heinz.zip"
            
            
'            Dim lResult As Long
'            Dim sURL As String
'            Dim sLocalFile As String
'
'            ' URL-Link der Datei, die heruntergeladen werden soll
'            sURL = "http://www.kisslive.de/bilder/home/KISS_Hannover.jpg"
'
'            ' Dateiname auf dem lokalen System
'            sLocalFile = "C:\Bild1.jpg"
'
'            ' Download ausführen
'            Screen.MousePointer = vbHourglass
'            lResult = URLDownloadToFile(0, sURL, sLocalFile, 0, 0)
'            Screen.MousePointer = vbNormal
'
'            ' Rückgabewert auswerten
'            If lResult = 0 Then
'              MsgBox "Download erfolgreich ausgeführt!"
'            Else
'              MsgBox "Fehler beim Download: " & _
'               "Entweder existiert die URL nicht, oder Sie haben " & _
'               "einen ungültigen Dateinamen angegeben!", vbCritical
'            End If
            
        Case Is = 7
            URLGoTo Me.hwnd, "http://www.kisslive.de/"
        Case Is = 8 'Bonus - Briefe 'Geburtstags - Brief
        
            URLGoTo Me.hwnd, "http://www.kisslive.de/service/bonusbriefe.html"
'            URLGoTo Me.hwnd, "http://www.kisslive.de/service/geburtstagsbriefe.html"
        Case Is = 9 'Karten - Terminal
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/178-elektonische-zahlungssysteme.html"
        
    End Select
    
    

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Label1_DblClick(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim ctmp        As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim cMess       As String
    
    
    cPfad = App.Path      'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Update"
    
    Select Case index
        Case 0
            If gbLokalModus Then
                gcDBPfad = "C:\aLeer"
            Else
            
                cdatei = "KISSLITE.INI"
                
                iFileNr = FreeFile
                Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
                If LOF(iFileNr) > 0 Then
                    
                    ctmp = Space$(LOF(iFileNr))
                    Get #iFileNr, 1, ctmp
                    gcDBPfad = ctmp
                    Close iFileNr
                End If
            End If
            
            cMess = "Datenbankpfad: " & gcDBPfad & vbCrLf & vbCrLf
            cMess = cMess & "Datenbankgröße: " & DabaFileSize & vbCrLf
            cMess = cMess & "letzte Komprimierung: " & DatumLastKompAnzeigen & " " & DatumLastKompZeitAnzeigen
            
            MsgBox cMess, vbInformation + vbOKOnly, "aktuelle Winkiss - Datenbankinformationen"
    End Select
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseProgrammeinstellungenWKL00"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    MsgBox "frmWKL00.Label1_DblClick: " & err.Number & " / " & err.Description
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim iRet As Integer
    
    If Button = 2 And Shift = 1 Then
        
        ctmp = gRegister.firma & vbCrLf
        ctmp = ctmp & gRegister.Plz & " " & gRegister.Ort & vbCrLf & vbCrLf
        ctmp = ctmp & gRegister.KdWert1 & " / "
        ctmp = ctmp & gRegister.KdWert2 & " / "
        ctmp = ctmp & gRegister.KdWert3 & " / "
        ctmp = ctmp & gRegister.KdWert4 & vbCrLf & vbCrLf
        ctmp = ctmp & gRegister.Confirm1 & " / "
        ctmp = ctmp & gRegister.Confirm2 & " / "
        ctmp = ctmp & gRegister.Confirm3 & " / "
        ctmp = ctmp & gRegister.Confirm4 & vbCrLf & vbCrLf
        ctmp = ctmp & gRegister.Datum & vbCrLf & vbCrLf & vbCrLf
        ctmp = ctmp & "Registrierung löschen?"
    
        iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "REGISTRIERUNG")
        If iRet = vbYes Then
            Kill gcSysPfad & gcRegDatei
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "LeseProgrammeinstellungenWKL00"
        Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        MsgBox "frmWKL00.Label2_MouseUp: " & err.Number & " / " & err.Description
    End If

End Sub




Private Sub Timer2_Timer()
On Error GoTo LOKAL_ERROR

    Dim sZeitKassDatstart As Date
    Dim sJetztforKassdat As Date
                

    If gbKSF = False Then
        If Format$(Time, "SS") = 15 Then
        
            If gsKassDatstart <> "" Then
        
                
                
                sJetztforKassdat = Format(TimeValue(Now), "HH:MM")
                sZeitKassDatstart = Format(TimeValue(gsKassDatstart), "HH:MM")
                
                If sZeitKassDatstart = sJetztforKassdat Then
                    Timer2.Enabled = False
                    anzeige "LASER", "", frmWKL00.Label2
                    
                    If gbBestAkt Then
                        anzeige "normal", "Der Tagesabschluss wird jetzt durchgeführt.", frmWKL00.Label2
                        schreibeProtokollNachtAblauf "Der Tagesabschluss wird jetzt durchgeführt."
                        
                        BESTAKTweg txtStatus
                    
                        If LoescheTagesAbschlussMODUL7(gcKasNum, picprogress, txtStatus, Label2, Label4(0), Label4(2), Label4(1)) Then
    
                        End If
                    End If
                
                    anzeige "normal", "Die Kassendatei wird jetzt übertragen.", frmWKL00.Label2
                    schreibeProtokollNachtAblauf "Die Kassendatei wird jetzt übertragen."
                    
                    theBigFTPFehlerZähler = 0
                    theBigFTPFehler = False
                
                    KassendatundStatcheck 'FTP Transfer bei Statistik oder F - Dateien
                    schreibeProtokollNachtAblauf "Die Kassendatei wird jetzt übertragen endet "
                    
                    Timer2.Enabled = True
                    anzeige "normal", "", frmWKL00.Label2
                   
                End If
            End If
        End If
    End If
    
    If gbTerminReminderSMS = True Then
        If Format$(Time, "SS") = 25 Then
        
            If gsTerminReminderstart <> "" Then
        
                
                sJetztforKassdat = Format(TimeValue(Now), "HH:MM")
                sZeitKassDatstart = Format(TimeValue(gsTerminReminderstart), "HH:MM")
                
                If sZeitKassDatstart = sJetztforKassdat Then
                    Timer2.Enabled = False
                    anzeige "LASER", "", frmWKL00.Label2
                    
                    Dim dateAuswertungstag      As Date
   
                    dateAuswertungstag = DateValue(Now) + glTageVorTermin
                    LeseOpeningsWKL82
                    VersendeTermineSMS dateAuswertungstag
        
                    Timer2.Enabled = True
                    anzeige "normal", "", frmWKL00.Label2
                   
                End If
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."
    
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
Private Function fnPruefeAnzahlArtikelWKL00() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnPruefeAnzahlArtikelWKL00 = 0
    
    cSQL = "Select count(*) as ANZ from ARTIKEL"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!anz) Then
            fnPruefeAnzahlArtikelWKL00 = rsrs!anz
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeAnzahlArtikelWKL00"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim k As Integer
    
    LogtoEnd Me
    dabkomp
    
    
    If gsEPartner = "ZVT" Then
        lese_ZVT_opt
        
        'close anwendung
        Dim hwnd&
        Dim Y As String
        Dim result&
        Dim Title$
    
        Y = gZVTPTitel
                                
        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
    
        Do
            result = GetWindowTextLength(hwnd) + 1
            Title = Space(result)
            result = GetWindowText(hwnd, Title, result)
            Title = Left$(Title, Len(Title) - 1)
    
            If InStr(1, Title, Y) Then
                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
    
            End If
    
            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop Until hwnd = 0
    End If
    
    If gbBestDateien = True And gsPfadBestandlive <> "" Then
    
        Y = "BestandLive"
                                
        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
    
        Do
            result = GetWindowTextLength(hwnd) + 1
            Title = Space(result)
            result = GetWindowText(hwnd, Title, result)
            Title = Left$(Title, Len(Title) - 1)
    
            If InStr(1, Title, Y) Then
                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
    
            End If
    
            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop Until hwnd = 0
    End If
    
    'auch die Display.exe
                    
    If gbZweitMoni Then

        Y = "Ihre Kundeninformationen"
                                    
        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
        
        Do
            result = GetWindowTextLength(hwnd) + 1
            Title = Space(result)
            result = GetWindowText(hwnd, Title, result)
            Title = Left$(Title, Len(Title) - 1)
    
            If InStr(1, Title, Y) Then
                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
    
            End If
    
            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop Until hwnd = 0
    End If
    
    
    
    
    
    allesbeenden
    
    'Was ist mit Nachtverarbeitung?
    
    If gbNacht = True Then
    
        gsAnzeigeText = ""
        Dim ctmp As String
        ctmp = "Nicht ausschalten!" & vbCrLf & vbCrLf
'        ctmp = ctmp & "Auf diesem Rechner ist die automatische Nachtverarbeitung aktiviert." & vbCrLf & vbCrLf
        ctmp = ctmp & "Die Nachtverarbeitung startet um " & gsNachtstart & " Uhr nur, wenn Sie das Programm 'Winkiss' eingeschaltet lassen." & vbCrLf

        gsAnzeigeText = ctmp
            
        frmWK21l.Show
        frmWK21l.Refresh
        frmWK21l.Command1.Visible = False
        frmWK21l.Refresh
        
        For k = 5 To 1 Step -1
        Pause 1
        frmWK21l.lbl6(0).Caption = ctmp & " ..." & k & " sec"
        frmWK21l.lbl6(0).Refresh
        Next k
    End If
    
    
    
    
    
    End 'Ende
    

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub dabkomp()
    On Error GoTo LOKAL_ERROR

    Dim lLastKompDate As Long
    Dim lHeute  As Long
    Dim sTime   As String
    Dim lTime   As Long
    Dim sSQL As String
    
    Screen.MousePointer = 11
'    MsgBox Weekday(DateValue(Now))
    
    If gbLokalModus = False Then
        'Zwangsoptimierung Anfang
        '1. Sind 7 Tage erreicht?
        '2. Zeit nach 18:00 Uhr ?
        '3. auch nicht Freitag ?
        '4. Ist man alleine in der Datenbank?
        
        sTime = TimeValue(Now)
        sTime = Format$(sTime, "HH:MM:SS")
        sTime = SwapStr(sTime, ":", "")
        lTime = CLng(sTime)
        
        lHeute = Fix(Now)
        lLastKompDate = DateValue(DatumLastKompAnzeigen)
        If lLastKompDate < (lHeute - 7) Then
    
            If lTime >= 180000 Then
            
                If Weekday(DateValue(Now)) <> 6 Then
    
                    If BistDualleineinderDatenbank Then
                        If gbDabakompautoNo = False Then
                        
                            lbl6(28).ForeColor = vbRed
                            lbl6(28).Visible = True
                            lbl6(28).Caption = "Heute ist der Zeitpunkt erreicht, um die Datenbank zu komprimieren."
                            lbl6(28).Refresh
            
                            lbl6(53).ForeColor = vbRed
                            lbl6(53).Caption = "Nehmen Sie sich noch  ein paar Minuten Zeit!"
                            lbl6(53).Refresh
                            
                            Pause (2)
                            
                            lbl6(28).ForeColor = vbRed
                            lbl6(28).Caption = "Im Falle eines Abbruchs werden Uhrzeit und Bediener protokolliert."
                            lbl6(28).Refresh
                            
                            lbl6(53).ForeColor = vbRed
                            lbl6(53).Visible = True
                            lbl6(53).Caption = "Schalten Sie den Rechner NICHT aus!!!"
                            lbl6(53).Refresh
                            
                            picprogress.Visible = True
                            lbl6(28).Visible = True
                            lbl6(53).Visible = True
                            
                            CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige
                         
                            sSQL = "update dbeinste set lastkomp='" & Date & "'"
                            sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
                            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                            picprogress.Visible = False
                            lbl6(28).Visible = True
                            lbl6(53).Visible = True
                            
                        
                        End If
                    End If
                End If
            End If
        End If
        'Zwangsoptimierung Ende
    End If
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dabkomp"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub allesbeenden()
    On Error GoTo LOKAL_ERROR

    Dim lLastKompDate As Long
    Dim lHeute  As Long
    Dim sTime   As String
    Dim lTime   As Long
    Dim bmerke As Boolean
    Dim i As Integer
    bmerke = gbFTPautomatic
    
    Screen.MousePointer = 11
    
    picprogress.Visible = True
    txtStatus.Text = 90
    
    If gbLokalModus = False Then
        gbFTPautomatic = True
        
        
        '*****Begin
        theBigFTPFehlerZähler = 0
        theBigFTPFehler = False
    
        KassendatundStatcheck 'FTP Transfer bei Statistik oder F - Dateien
        
         '*****Ende
        gbFTPautomatic = bmerke
    End If
    
    
    txtStatus.Text = 80
    If gbLocalSec Then 'ist unter Programmeinstellungen/Voreinstellungen Localsecurity gefordert?
        If gbLokalModus = False Then
        
            If gbAutoLokalModus Then

                If Not FileExists("C:\aleer\kissdata.mdb") Then
                    HoleLokalDB
                End If
            Else
                HoleLokalDB
                
            End If
            
        End If
    End If
    
    txtStatus.Text = 70
    
    Label2.Caption = "Moment bitte ... Schließe Datenbanken ... Gebe Ressourcen frei ..."
    Label2.Refresh
    
    gcAnwender = ""
    
    txtStatus.Text = 60
    
    If gbLokalModus = False Then
        txtStatus.Text = 50
'        AbmeldungDABA
    End If
        
    txtStatus.Text = 40
    AbmeldungDabaNew
    
    KomprimieredieAPP
    
    Label2.Caption = "Alles ordnungsgemäß beendet..."
    Label2.Refresh
    
    txtStatus.Text = 30
    If gbLokalModus = False Then
        schreibeProtokoll "Abmeldung: meldet sich ab(kissdata.mdb)."
        schreibeProtokollBENUTZERablauf "Abmeldung"
        
        schreibeProtokoll "Abmeldung: meldet sich ab(kissapp.mdb)."
    End If
    txtStatus.Text = 20
    gdApp.Close
    txtStatus.Text = 10
    gdBase.Close
    
    If gbNetzLW Then
        KappeNetzVerbindung
    End If
    txtStatus.Text = 0
    picprogress.Visible = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "allesbeenden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub systembildcheck_all()
On Error GoTo LOKAL_ERROR

    systembildcheck "Tabelle.jpg"
    systembildcheck "Tastatur.jpg"
    systembildcheck "Kalender.jpg"
    systembildcheck "Zurück.jpg"
    systembildcheck "Vor.jpg"
    systembildcheck "Visa.jpg"
    systembildcheck "Visa_kl.jpg"
    systembildcheck "American-Express.jpg"
    systembildcheck "American-Express_kl.jpg"
    systembildcheck "Diners-Club.jpg"
    systembildcheck "Diners-Club_kl.jpg"
    systembildcheck "Mastercard.jpg"
    systembildcheck "Mastercard_kl.jpg"
    systembildcheck "Maestro.jpg"
    systembildcheck "Maestro_kl.jpg"
    systembildcheck "diverse.jpg"
    systembildcheck "switch.jpg"
    systembildcheck "leute1.jpg"
    systembildcheck "Rechts.jpg"
    systembildcheck "Links.jpg"
    systembildcheck "Rauf.jpg"
    systembildcheck "Runter.jpg"
    systembildcheck "futura.jpg"
    systembildcheck "WinEwws.gif"
    systembildcheck "sqlserver.jpg"
    systembildcheck "EC.jpg"
    systembildcheck "Brief.gif"
    systembildcheck "Briefrot.gif"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "systembildcheck_all"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub waehrungbildcheck_all()
On Error GoTo LOKAL_ERROR

    waehrungbildcheck "0k.jpg"
    waehrungbildcheck "1k.jpg"
    waehrungbildcheck "2k.jpg"
    waehrungbildcheck "3k.jpg"
    waehrungbildcheck "4k.jpg"
    waehrungbildcheck "5k.jpg"
    waehrungbildcheck "6k.jpg"
    waehrungbildcheck "7k.jpg"
    waehrungbildcheck "8k.jpg"
    waehrungbildcheck "9k.jpg"
    waehrungbildcheck "10k.jpg"
    waehrungbildcheck "11k.jpg"
    waehrungbildcheck "12k.jpg"
    waehrungbildcheck "13k.jpg"
    waehrungbildcheck "14k.jpg"
    
    waehrungbildcheck "0g.jpg"
    waehrungbildcheck "1g.jpg"
    waehrungbildcheck "2g.jpg"
    waehrungbildcheck "3g.jpg"
    waehrungbildcheck "4g.jpg"
    waehrungbildcheck "5g.jpg"
    waehrungbildcheck "6g.jpg"
    waehrungbildcheck "7g.jpg"
    waehrungbildcheck "8g.jpg"
    waehrungbildcheck "9g.jpg"
    waehrungbildcheck "10g.jpg"
    waehrungbildcheck "11g.jpg"
    waehrungbildcheck "12g.jpg"
    waehrungbildcheck "13g.jpg"
    waehrungbildcheck "14g.jpg"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "waehrungbildcheck_all"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub systembildcheck(cBild As String)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    
    'check ob Bild vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & cBild) Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & cBild

        cZiel = cPfad & "\PICTURE\System\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & cBild

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "systembildcheck"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub waehrungbildcheck(cBild As String)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    
    'check ob Bild vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & cBild) Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & cBild

        cZiel = cPfad & "\PICTURE\EUR\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & cBild

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "waehrungbildcheck"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub dabkomp1()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    Screen.MousePointer = 11
    
    If gbLokalModus = False Then
        If BistDualleineinderDatenbank Then
             
            Pause 1
            lbl6(28).ForeColor = vbRed
            lbl6(28).Caption = "Nachtverarbeitung: Datenbank wird komprimiert..."
            lbl6(28).Refresh
             
            lbl6(53).ForeColor = vbRed
            lbl6(53).Visible = True
            lbl6(53).Caption = ""
            lbl6(53).Refresh
             
            picprogress.Visible = True
            lbl6(28).Visible = True
            lbl6(53).Visible = True
             
            CompactmyDaba gcDBPfad, "KISSDATA.MDB", gdBase, lbl6(53), txtStatus, lbl6(28), gbOhneAnzeige

            sSQL = "update dbeinste set lastkomp='" & Date & "'"
            sSQL = sSQL & " ,lastkomptime='" & TimeValue(Now) & "'"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            picprogress.Visible = False
            lbl6(28).Visible = True
            lbl6(53).Visible = True
             
        End If
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dabkomp1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub erstellevorschlag(cLinr As String, labglo As Label, lab As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim lDatDiff    As Long
    Dim lAnzSatz    As Long
    Dim lAnzahl     As Long
    Dim lEindeckung As Long
    Dim lMengeVor   As Long
    Dim lMengeAkt   As Long
    Dim lcount      As Long
    Dim lMinMen     As Long
    Dim lMenge      As Long
    Dim lBestand    As Long
    Dim lMinBest    As Long
    Dim counter     As Long
    
    Dim sSQLKOPF    As String
    Dim sSQLALLE    As String
    Dim cPfad       As String
    Dim cLfdJahr    As String
    Dim cLfdMonat   As String
    Dim cFeld       As String
    Dim cSQL        As String
    Dim cSQL2       As String
    Dim cSQL3       As String
    Dim cLPZMin     As String
    Dim cLPZMax     As String
    Dim cEindeck    As String
    Dim cBevorrat   As String
    Dim cProdLinien As String
    Dim cJahr       As String
    Dim cMonat      As String
    Dim cArtNr      As String
    Dim cTmpFilnr   As String
    Dim cDatum      As String
    Dim ctmp        As String
    Dim sAccPfad    As String
    Dim cPLINR      As String
    Dim sSQLCheckLPZ As String
    
    Dim dMindest    As Double
    Dim dFaktor     As Double
    Dim dBestellung As Double
    Dim dMenge      As Double
    Dim dMengeTag   As Double
    
    Dim bFehler     As Boolean
    Dim bMinBest    As Boolean
    
    Dim iTmp        As Integer
    Dim i           As Integer
    Dim iGefunden   As Integer
    Dim iFiliale    As Integer
    
    Dim tdTd        As TableDef
    
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim rsRs2       As Recordset
    Dim rsRs3       As Recordset
    Dim rsVorschlz  As Recordset
    Dim rsArtlief   As Recordset
    
    picprogress.Visible = True
    txtStatus.Text = 0
    
    labglo.Visible = True
    labglo.Caption = "Bestellvorschlag wird errechnet..."
    labglo.Refresh
    
    lab.Visible = True
    lab.Caption = cLinr & " " & ermLiefBez(CLng(cLinr))
    lab.Refresh
    
    cPfad = App.Path
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    bFehler = False

    txtStatus.Text = 5

    
    vorschlzRefresh
    
    bFehler = False
    

    cLPZMin = "1" 'Trim$(Text1(1).Text)
    cLPZMax = "999" 'Trim$(Text1(2).Text)
    cEindeck = "4" ' Trim$(Str$(Val(Text1(3).Text)))
    
    
    cBevorrat = "1" ' Trim$(Text1(4).Text)
    

    txtStatus.Text = 10

    sAccPfad = App.Path & "\kissapp.mdb"
    sSQLKOPF = "insert into vorschlz in '" & sAccPfad & "' "

       
    sSQLALLE = " Select distinct"
    sSQLALLE = sSQLALLE & "  A.ARTNR "
    sSQLALLE = sSQLALLE & ", A.BEZEICH "
    sSQLALLE = sSQLALLE & ", A.AGN "
    sSQLALLE = sSQLALLE & ", A.PGN "
    sSQLALLE = sSQLALLE & ", B.LEKPR "
    sSQLALLE = sSQLALLE & ", A.AWM "
    sSQLALLE = sSQLALLE & ", A.EKPR "
    sSQLALLE = sSQLALLE & ", A.VKPR "
    sSQLALLE = sSQLALLE & ", " & cLinr & " as LINR "
    sSQLALLE = sSQLALLE & ", B.LIBESNR "
    sSQLALLE = sSQLALLE & ", A.KVKPR1 "
    sSQLALLE = sSQLALLE & ", A.EAN "
    sSQLALLE = sSQLALLE & ", 0 as MOPREIS "
    sSQLALLE = sSQLALLE & ", A.RKZ "
    sSQLALLE = sSQLALLE & ", A.LPZ "
    
    sSQLALLE = sSQLALLE & ", A.NOTIZEN "
    sSQLALLE = sSQLALLE & ", B.MINMEN "
    sSQLALLE = sSQLALLE & ", 0 as MINBEST "
    sSQLALLE = sSQLALLE & ", 0 as INBEST "
    sSQLALLE = sSQLALLE & ", 0 as ANZEIGE "
    sSQLALLE = sSQLALLE & ", 0 as FAKTOR "
    sSQLALLE = sSQLALLE & ", A.BESTAND "
    sSQLALLE = sSQLALLE & ", NULL as LPZ_BIS "
    sSQLALLE = sSQLALLE & ", NULL as LPZ_VON "
    sSQLALLE = sSQLALLE & ", " & cBevorrat & " as BEVORRAT "
    sSQLALLE = sSQLALLE & ", " & cEindeck & " as EINDECK "
    sSQLALLE = sSQLALLE & ", NULL as VKAMo1 "
    sSQLALLE = sSQLALLE & ", NULL as VKVMo1 "
    sSQLALLE = sSQLALLE & ", NULL as VKLJ1 "
    sSQLALLE = sSQLALLE & ", NULL as VKVJ1 "
    sSQLALLE = sSQLALLE & ", NULL as MITTEILUNG"
    sSQLALLE = sSQLALLE & ", NULL as LJ1 , NULL as LJ2, NULL as LJ3, NULL as LJ4, NULL as LJ5,NULL as LJ6, NULL as LJ7,NULL as LJ8,NULL as LJ9,NULL as LJ10,NULL as LJ11,NULL as LJ12 "
    sSQLALLE = sSQLALLE & ", NULL as VJ1 , NULL as VJ2, NULL as VJ3, NULL as VJ4, NULL as VJ5,NULL as VJ6, NULL as VJ7,NULL as VJ8,NULL as VJ9,NULL as VJ10,NULL as VJ11,NULL as VJ12 "
     
    cSQL = "Select * from LINBEZ where LINR = " & cLinr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        cSQL = sSQLKOPF
        cSQL = cSQL & sSQLALLE
        cSQL = cSQL & " from ARTIKEL A, ARTLIEF B"
        cSQL = cSQL & " where B.LINR = " & cLinr & " "
        cSQL = cSQL & " and A.ARTNR = B.ARTNR "
        cSQL = cSQL & " and A.GEFUEHRT = 'J'"
        cSQL = cSQL & " and (A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null)"
        cSQL = cSQL & " and A.LPZ >= " & cLPZMin & " and A.LPZ <= " & cLPZMax & " "

        cSQL = cSQL & " order by A.LPZ, A.BEZEICH "
        
        gdBase.Execute cSQL, dbFailOnError
          
    Else  '** rsRs.EOF ( Lieferant hat keine Produktlinien!)
        
        cSQL = sSQLKOPF
        cSQL = cSQL & sSQLALLE
        cSQL = cSQL & " from ARTIKEL A, ARTLIEF B "
        cSQL = cSQL & " where B.LINR = " & cLinr & " "
        cSQL = cSQL & " and A.ARTNR = B.ARTNR "
        cSQL = cSQL & " and A.GEFUEHRT = 'J'"
        cSQL = cSQL & " and (A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null)"
        cSQL = cSQL & " order by A.LPZ, A.BEZEICH "
        gdBase.Execute cSQL, dbFailOnError
       
    End If   '** ende von Hole aus der ARTIKEL und ARTLIEF
    '**************************************************************************************

    rsrs.Close: Set rsrs = Nothing
    
    txtStatus.Text = 20

    
    'Vorhandene Bestellungen einlesen
    Set rsrs = gdApp.OpenRecordset("VORSCHLZ", dbOpenTable)
    rsrs.index = "ARTNR"
    
    '************* BESTREST hat kein FILIALNR ******************
    cSQL = "Select ARTNR, SUM(BESTVOR) as INBEST from BESTREST "
    cSQL = cSQL & "where LINR = " & cLinr & " group by ARTNR "
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    
    If Not rsRs2.EOF Then
        rsRs2.MoveFirst
        Do While Not rsRs2.EOF
            If Not IsNull(rsRs2!artnr) Then
                cArtNr = rsRs2!artnr
            Else
                cArtNr = "-1"
            End If
            rsrs.Seek "=", cArtNr
            If Not rsrs.NoMatch Then
                rsrs.Edit
                rsrs!INBEST = rsRs2!INBEST
                rsrs.Update
            End If
            
            rsRs2.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            
    cLfdJahr = Year(Now)
    cLfdMonat = Month(Now)
    
    txtStatus.Text = 35

    Dim sSQL2 As String
    Dim sSQL3 As String
   
   
    
   
    cSQL = "Select * from VORSCHLZ"
    Set rsrs1 = gdApp.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!artnr) Then
                cArtNr = rsrs1!artnr
            Else
                cArtNr = ""
            End If
            
            '** Verkäufe laufendes Jahr einlesen **
            sSQL3 = "select * from UMSARTJ where ARTNR = " & cArtNr
            sSQL3 = sSQL3 & " and  JAHR = " & cLfdJahr
            Set rsrs = gdBase.OpenRecordset(sSQL3)
        
            If Not rsrs.EOF Then
                rsrs1.Edit
                rsrs1!VKLJ1 = rsrs!ANZAHLJ
            Else
                rsrs1.Edit
                rsrs1!VKLJ1 = 0
            End If
            rsrs.Close: Set rsrs = Nothing
            
            '** Verkäufe laufender Monat einlesen **
            
            sSQL2 = "select * from UMS_ART where ARTNR = " & cArtNr
            sSQL2 = sSQL2 & " and  JAHR = " & cLfdJahr
            sSQL2 = sSQL2 & " and  MONAT = " & cLfdMonat
            Set rsRs2 = gdBase.OpenRecordset(sSQL2)


            If Not rsRs2.EOF Then
                rsrs1!VKAMo1 = rsRs2!ANZAHL
            Else
                rsrs1!VKAMo1 = 0
            End If
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            
            
            
            rsrs1.Update
            rsrs1.MoveNext
        Loop
        
    End If
    rsrs1.Close: Set rsrs1 = Nothing
    
    txtStatus.Text = 40

    cSQL = "Select * from VORSCHLZ"
    Set rsrs1 = gdApp.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!artnr) Then
                cArtNr = rsrs1!artnr
            Else
                cArtNr = ""
            End If
            
            cLfdJahr = Year(Now)
            cLfdMonat = Month(Now)
            cLfdMonat = Trim$(Str$(((Val(cLfdMonat)) - 1)))
            If cLfdMonat = "0" Then
                cLfdMonat = "12"
                cLfdJahr = Trim$(Str$(((Val(cLfdJahr)) - 1)))
            End If
            '** Verkäufe Vor-Monat einlesen **
            
            sSQL2 = "select * from UMS_ART where ARTNR = " & cArtNr
            sSQL2 = sSQL2 & " and  JAHR = " & cLfdJahr
            sSQL2 = sSQL2 & " and  MONAT = " & cLfdMonat
            Set rsRs2 = gdBase.OpenRecordset(sSQL2)

            If Not rsRs2.EOF Then
                rsrs1.Edit
                rsrs1!VKVMO1 = rsRs2!ANZAHL
            Else
                rsrs1.Edit
                rsrs1!VKVMO1 = 0
            End If
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing

            cLfdJahr = Year(Now)
            cLfdJahr = Trim$(Str$(((Val(cLfdJahr)) - 1)))
            '** Verkäufe Vor-Jahr einlesen **
            
            sSQL3 = "select * from UMSARTJ where ARTNR = " & cArtNr
            sSQL3 = sSQL3 & " and  JAHR = " & cLfdJahr
            Set rsrs = gdBase.OpenRecordset(sSQL3)
 
            If Not rsrs.EOF Then
                rsrs1!VKVJ1 = rsrs!ANZAHLJ
            Else
                rsrs1!VKVJ1 = 0
            End If
            rsrs.Close: Set rsrs = Nothing
            rsrs1.Update
            rsrs1.MoveNext
        Loop
    End If
    rsrs1.Close: Set rsrs1 = Nothing

    txtStatus.Text = 45

''    '********** GraphFüllenWKL43 ************
    GraphFuellenWKL43 "UMS_ART"
    
    txtStatus.Text = 48
     
    FaktorBerechnungWKL43
        
    txtStatus.Text = 51
    If UeberlaufFehler Then
        Screen.MousePointer = 0

        Exit Sub
    End If

    lEindeckung = CLng(cEindeck) 'Text1(3).Text
    
    dFaktor = Val(cBevorrat)

    txtStatus.Text = "55"
    
    cSQL = "Select ARTNR, FAKTOR, BESTAND, MINMEN, MINBEST, BESTVOR "
    cSQL = cSQL & " from VORSCHLZ"
    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        i = 1
        Do While Not rsrs.EOF
            If IsNull(rsrs!artnr) Then
            Else
                cArtNr = rsrs!artnr

                If cArtNr = 300101 Then
                    cArtNr = cArtNr
                End If
            End If
            If IsNull(rsrs!Faktor) Then
                dMenge = 0
            Else
                If rsrs!Faktor > 0 Then
                    dMenge = rsrs!Faktor
                Else
                    dMenge = 0
                End If
            End If
            
            If IsNull(rsrs!BESTAND) Then
                lBestand = 0
            Else
                lBestand = rsrs!BESTAND
                If lBestand < 0 Then
                    lBestand = 0
                End If
            End If
            
            If IsNull(rsrs!MINMEN) Then
                lMinMen = 0
            Else
                lMinMen = rsrs!MINMEN
            End If
            
            If IsNull(rsrs!MINBEST) Then
                lMinBest = 0
            Else
                lMinBest = rsrs!MINBEST
            End If
            
            dBestellung = dMenge * ((lEindeckung * 7) * dFaktor)
            
            If bMinBest Then
                lBestand = lBestand - lMinBest
            End If
            
            dBestellung = dBestellung - lBestand
            If dBestellung > 0 Then
                If Not IsNull(rsrs!MINMEN) Then
                    dMindest = rsrs!MINMEN
                Else
                    dMindest = 1
                End If
                If dMindest = 0 Then
                    dMindest = 1
                End If
            Else
                dBestellung = 0
                dMindest = 1
            End If

            dBestellung = dBestellung / dMindest
            If dBestellung <> Fix(dBestellung) Then
                dBestellung = Fix(dBestellung) + 1
            End If
            dBestellung = dBestellung * dMindest

            rsrs.Edit
            rsrs!BESTVOR = dBestellung
            rsrs!MINBEST = iTmp
            rsrs.Update
            i = i + 1
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    txtStatus.Text = "60"

    cSQL = "Update VORSCHLZ set BESTVOR = BESTVOR - INBEST"
    gdApp.Execute cSQL, dbFailOnError
    txtStatus.Text = "62"
    
    cSQL = "Update VORSCHLZ set BESTVOR = 0 where BESTVOR < 0"
    gdApp.Execute cSQL, dbFailOnError
    txtStatus.Text = "64"
    
    cSQL = "Update VORSCHLZ set ANZEIGE = 1 where BESTVOR <> 0 or BESTAND <> 0 or FAKTOR <> 0 "
    gdApp.Execute cSQL, dbFailOnError
    txtStatus.Text = "66"
    

    cSQL = "Update VORSCHLZ set AWM = '0' where AWM = '99' "  'löschen der "Artikel anfügen" Farbe
    gdApp.Execute cSQL, dbFailOnError
    txtStatus.Text = "67"
    
    loeschNEW "X" & cLinr, gdApp
    
    txtStatus.Text = "69"
    
    cSQL = "Select * into X" & cLinr & " from vorschlz "
    gdApp.Execute cSQL, dbFailOnError
    
    schreibalias "X" & cLinr, "automatisch erstellt"
    
    txtStatus.Text = "88"
    
    labglo.Caption = ""
    labglo.Refresh
    
    
    lab.Caption = ""
    lab.Refresh
    picprogress.Visible = False
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleVORSCHLag"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Dim lAnz                As Long
    Dim lRet                As Long
    Dim lfail               As Long
    Dim lHeute              As Long
    Dim cPfad               As String
    Dim cPfadA              As String
    Dim ctmp                As String
    Dim cSQL                As String
    Dim sSQL                As String
    Dim iFileNr             As Integer
    Dim iRet                As Integer
    Dim rsrs                As Recordset
    Dim ctemp               As String
    Dim bPfad               As Boolean
    Dim i                   As Integer
    Dim DateNow             As Date
    Dim iDayUnterschied     As Long
    Dim iStep               As Integer
    Dim sRechner            As String
    Dim cQuelle             As String
    Dim cZiel               As String
    Dim j                   As Integer
    Dim Task$
    Dim lFileSize           As Long
    Dim cPfadQ As String
    Dim cpfadZ As String
    
    
    gbSQLSERVER = False
    
    glErrtime = 0
''    If FileExists(App.Path & "\DAO360.dll") = True Then
''        lFileSize = fnFileSize(App.Path & "\DAO360.dll")
''
''        If FileExists(App.Path & "\DAO.txt") = False Then
''
''            If DLLcheckZentStart("C:\Programme\Gemeinsame Dateien\Microsoft Shared\DAO", "DAO360.dll", lFileSize) = False Then
''
''                schreibeDAOtxt
''
''                'Start Zentralestarter
''                Task = Shell(App.Path & "\Winsta.exe", 1) 'Updater öffnen
''
''                'Zentrale Ende
''                End
''
''            End If
''
''        Else
''            Kill App.Path & "\DAO.txt"
''        End If
''    End If
''
''    doIt "C:\Programme\Gemeinsame Dateien\Microsoft Shared\DAO\DAO360.dll", True  'False für unregister
''
''    systemdatcheck "MSWINSCK.ocx"

    mailDLLcheck

    
    If (GetKeyState(vbKeyCapital) = 1) Then
    
'        'CAPS-Lock deaktivieren (falls aktiviert)
'        MsgBox "CAPS-Lock (Großschreibmodus) ist eingeschaltet und wird jetzt abgestellt!", vbInformation, "Winkiss Hinweis:"
        
        
        KeyboardChangeState vbKeyCapital
    End If
    
    
    

    'Spruchrahmen aus
    lbl6(0).BorderStyle = 0
    
    systemdatcheck "sevcmd3.ocx"
    systemdatcheck "sevlmlib.dll"
    
    'mailocxcheck
    
'    Kill App.Path & "\sevZip32.dll"
    zip32check

    gsAnforderung = ""
    j = 0
    gsLokalTabellen(j) = "Kunden"
    j = j + 1
    gsLokalTabellen(j) = "Artikel"
    j = j + 1
    gsLokalTabellen(j) = "Artlief"
    j = j + 1
    gsLokalTabellen(j) = "Gutsch"
    j = j + 1
    gsLokalTabellen(j) = "Lisrt"
    j = j + 1
    gsLokalTabellen(j) = "Kredit"
    j = j + 1
    gsLokalTabellen(j) = "Zugang"
    
    bAbschlussjetzt = False
    UeberlaufFehler = False
    
    picprogress.Visible = True
    '1.Ist Winkiss schon gestartet
    iStep = 1
    txtStatus.Text = iStep * 2
    
    If App.PrevInstance Then
        MsgBox "Winkiss ist schon aktiv", vbInformation, "Winkiss Hinweis:"
        End
    End If
    
    Screen.MousePointer = 11
    
    '2.Standardwerte setzen
    iStep = 2
    txtStatus.Text = iStep * 2
    
    Label2.Caption = "Programm - Standardwerte setzen..."
    Label2.Refresh
    
    gbfrm27 = False
    bPfad = True
    bStammdaten = False
    bKasse = False
    bStatistiken = False
    bTermine = False
    bListen = False
    bService = False
    byteSortReihen = 1
    lHeute = Fix(Now)
    gbStornoErlaubt = True
    gbRabatt = True
    gbAPI = True
    gbNetzLW = False
    gcWaehrung = "EUR"
    giCopyMod = 0
    gbLibesnrSeek = False
    
    giUmleitgrund = 60
    sErrDabapfad = ""
    
    '3.Das Formular positionieren
    iStep = 3
    txtStatus.Text = iStep * 2
    Positionieren
    
    '4.Das Formular skalieren
    iStep = 4
    txtStatus.Text = iStep * 2
    Modul6.Skalieren Me, True, True
    
    Me.WindowState = 2
    

    
    '5.Berechtigungen setzen
    iStep = 5
    txtStatus.Text = iStep * 2
    SetzeCommandTags frmWKL00
    
    '6.Vollversion oder Demoversion?
    iStep = 6
    txtStatus.Text = iStep * 2
    
    Show
    DoEvents
    
    '7.Anwendungspfad
    iStep = 7
    txtStatus.Text = iStep * 2
    
    gcPfad = App.Path
    If Right(gcPfad, 1) <> "\" Then
        gcPfad = gcPfad & "\"
    End If
    
    cPfadA = gcPfad
    
    '8.festen Artikeltext definieren
    iStep = 8
    txtStatus.Text = iStep * 2
    
    gbDivKosmetik = False
    iFileNr = FreeFile
    Open cPfadA & "BonText.CFG" For Binary As iFileNr
    If LOF(iFileNr) > 0 Then
        gbDivKosmetik = True
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        gcDivKosmetik = ctmp
        Close iFileNr
    Else
        gbDivKosmetik = False
        Close iFileNr
        Kill cPfadA & "BonText.CFG"
    End If
    
    '*******************************************************
    '* Hintertür für REGISTRIERUNG offen?
    '*******************************************************
    iStep = 9
    txtStatus.Text = iStep * 2
    iFileNr = FreeFile
    Open cPfadA & "REGISTER.CFG" For Binary As iFileNr
    If LOF(iFileNr) > 0 Then
        Close iFileNr
        gbRegister = False
    Else
        Close iFileNr
        Kill cPfadA & "REGISTER.CFG"
        gbRegister = True
    End If
    
    '**********************************************************
    '* jeder Verkauf nur mit Kundennummer möglich
    '**********************************************************
    iStep = 10
    txtStatus.Text = iStep * 2
    gbZwangsKdNr = False
    iFileNr = FreeFile
    Open cPfadA & "KDNR.CFG" For Binary As iFileNr
    If LOF(iFileNr) > 0 Then
        Close iFileNr
        gbZwangsKdNr = True
    Else
        Close iFileNr
        Kill cPfadA & "KDNR.CFG"
    End If
    
    '***************************************************************
    '* Kassenschublade über COM-Port statt über Bondrucker öffnen?
    '***************************************************************
    iStep = 11
    txtStatus.Text = iStep * 2
    gbLadeCom = False
    iFileNr = FreeFile
    Open cPfadA & "LADECOM.CFG" For Binary As iFileNr
    If LOF(iFileNr) > 0 Then
        gcLadeCom = LOF(iFileNr)
        Get #iFileNr, 1, gcLadeCom
        Close iFileNr
        gbLadeCom = True
    Else
        Close iFileNr
        Kill cPfadA & "LADECOM.CFG"
    End If

    iStep = 12
    txtStatus.Text = iStep * 2
    '***************************************************************
    '* Lese Zugriffspfad zur WinKISS-Datenbank
    '***************************************************************
    Label2.Caption = "Verbindung mit Datenbanken..."
    Label2.Refresh
    
    iRet = fnLeseIniDateiWKL00()
    
    If iRet = 1 Then
        gbLokalModus = True '***Hier wird der lokale Modus gesetzt
                                '***nur die Kasse funktioniert
                                '***jetzt wird eine Lokal.cfg geschrieben (in den Anwendungspfad)
                                '***ist diese vorhanden, dann geht es lokal weiter
    Else
        If Not Modul6.FindFile(gcDBPfad, "kissdata.mdb") Then
            gbLokalModus = True
        Else
            gbLokalModus = False
        End If
    End If
    
    
    '*******
    Do While FileExists(gcDBPfad & "\Komprimierung.txt") = True
   
        iFileNr = FreeFile
        Open gcDBPfad & "\Komprimierung.txt" For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then
            ctmp = Space$(LOF(iFileNr))
            Get #iFileNr, 1, ctmp
            
            Close iFileNr
        Else
            Close iFileNr
        End If
    
        gsAnzeigeText = ctmp & vbCrLf
        gsAnzeigeText = gsAnzeigeText & "Achten Sie darauf, dass die Komprimierung nie abgebrochen wird!" & vbCrLf
        gsAnzeigeText = gsAnzeigeText & "Gehen Sie auf SERVICE/ PROGRAMMEINSTELLUNGEN/ DATENBANK und drücken dort 'komprimieren'!" & vbCrLf
        gsAnzeigeText = gsAnzeigeText & gcDBPfad

        frmWK21m.Label1.Caption = "Komprimierung"
        frmWK21m.Show 1
        
        If gbnachkomp = True Then
            Command3(1).Enabled = False
            Exit Do
        End If
    Loop
    '******
    
    
    iStep = 13
    txtStatus.Text = iStep * 2
    If gbLokalModus Then
        gcDBPfad = "C:\aLeer"
        iFileNr = FreeFile
        Open gcPfad & "\Lokal.CFG" For Binary As #iFileNr

        If LOF(iFileNr) > 0 Then
            Close iFileNr
        Else
            ctemp = "Rechner befindet sich zur Zeit im lokalen Modus"
            Put #iFileNr, 1, ctemp
            Close iFileNr
        End If
    End If
    
    iStep = 14
    txtStatus.Text = iStep * 2
    
    
    
    If gcDBPfad = "" Then

        'Umleitung in wkl60

        giUmleitgrund = 1 'Datenbankpfad weg

        gcUmleittxt = "Es kann keine Verbindung zur Datenbank aufgebaut werden!" & vbCrLf
        gcUmleittxt = gcUmleittxt & "Drücken Sie 'Weiter', um dann den Datenbankpfad zu speichern!" & vbCrLf

        frmWKL60.Show 1

        Command12_Click 4


        iRet = fnLeseIniDateiWKL00()


    End If

    

    'Ab hier ist der Zugriffspfad zur Datenbank bekannt.
    '**********************************************************************************
    '* Bei Netzwerkbetrieb und nicht vorhandener Verbindung zum Datenbank-Server kommt
    '* es beim Öffnen der Datenbank zu Fehlern. Diese Fehler werden jetzt abgefangen!
    '*
    '* Reparatur: Datenbankpfad im Windows Explorer wiederherstellen
    '**********************************************************************************
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iStep = 15
    txtStatus.Text = iStep * 2
    
    
    If Not FileExists(cPfad & "kissdata.mdb") Then

        If FileExists(cPfadA & "kissapp.mdb") Then
            sErrDabapfad = cPfadA & "kissapp.mdb"
            Set gdApp = OpenDatabase(cPfadA & "kissapp.mdb", False)

'            schreibeProtokoll "Anmeldung: meldet sich an(kissapp.mdb)."
        Else

            MsgBox "Die KISSAPP.MDB wird nicht gefunden.", vbInformation, "Winkiss Hinweis:"
            End 'ende
        End If

        giUmleitgrund = 1 'Datenbankpfad weg

        gcUmleittxt = "Es kann keine Verbindung zur Datenbank aufgebaut werden!" & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Drücken Sie 'Weiter', um dann den Datenbankpfad zu speichern!" & vbCrLf

        Kill cPfadA & "Lokal.CFG"

        frmWKL60.Show 1

        Command12_Click 4
        gdApp.Close
    End If
    
    
    
    iStep = 15
    txtStatus.Text = iStep * 2
    '********************************************************************
    '* Die Informationen über das Programm und den Zugriffspfad zur
    '* Datenbank werden in der Registry gespeichert, damit diese
    '* Informationen für ein automatisiertes Update zur Vefügung stehen
    '********************************************************************
    Label2.Caption = "Prüfe Programmregistrierung..."
    Label2.Refresh
    
    PruefeRegistryEintragProgMOD01
    PruefeRegistryEintragDataMOD01
    
    '********************************************************************
    '* Lese Informationen über die zugeordneten Drucker
    '* - Listendrucker
    '* - Bondrucker
    '* - Etikettendrucker
    '* - Faxdrucker
    '********************************************************************
    

    
    gcDBPfad = ShortPath(gcDBPfad)
'    MsgBox gcDBPfad
    
    iStep = 16
    txtStatus.Text = iStep * 2
    
    Label2.Caption = "Prüfe alle Unterverzeichnisse..."
    Label2.Refresh
    '********************************************************************
    '* Prüfe, ob alle notwendigen Unterverzeichnisse im Datenbank-
    '* verzeichnis vorhanden sind. Wenn nicht, dann an
    '* in das bekannte DOS-Format (Begrenzung auf 8 Stellen)
    '********************************************************************
    PruefeSubDir
    
    
    sRechner = rechnername
    srechnertab = sRechner

    
    '********************************************************************
    '* WinKISS-Datenbank öffnen
    '********************************************************************
    
    
    iStep = 17
    txtStatus.Text = iStep * 2
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    'Falle

'''''''    Dim sdatnow As Date
'''''''    Dim scheckdat As Date
'''''''
'''''''    scheckdat = "22.12.2005"
'''''''    sdatnow = DateValue(Now)
'''''''
'''''''
'''''''    If CLng(sdatnow) > CLng(scheckdat) Then
'''''''
'''''''        Dim lWert As Long
'''''''        Dim cSysPfad As String
'''''''
'''''''        cSysPfad = Space$(255)
'''''''        lWert = GetSystemDirectory(cSysPfad, Len(cSysPfad))
'''''''        cSysPfad = Left(cSysPfad, lWert)
'''''''        If Right(cSysPfad, 1) <> "\" Then
'''''''            cSysPfad = cSysPfad & "\"
'''''''        End If
'''''''
'''''''        Kill cSysPfad & gcRegDatei
'''''''
'''''''        Kill App.Path & "\lokal.mdb"
'''''''
'''''''        Set gdBase = OpenDatabase(cpfad & "kissdata.mdb", True, False)
'''''''        gdBase.NewPassword "", "gameoveR"
'''''''        gdBase.Close
'''''''
'''''''        giUmleitgrund = 3 'unbekannter Fehler
'''''''        gcUmleittxt = "Das Zeitschloss ist aktiviert." & vbCrLf
'''''''        gcUmleittxt = gcUmleittxt & "Bitte rufen Sie die Hotline an!"
'''''''        frmWKL60.Show 1
'''''''        End
'''''''    End If
'''''''
'''''''    'Falle ende

' Guten Rutsch
'    pict1.Visible = False
'    If DateValue(Now) > CLng(DateValue("24.12.2008")) And DateValue(Now) < CLng(DateValue("01.01.2009")) Then
'        If Modul6.FindFile(App.Path, "\rutsch.jpg") = True Then
'            pict1.Picture = LoadPicture(App.Path & "\rutsch.jpg")
'            pict1.Visible = True
'
'        End If
'    End If
    
    DoEvents
    
    If FileExists(cPfad & "kissdata.mdb") Then

        sErrDabapfad = cPfad & "kissdata.mdb"
        Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)

        schreibeProtokoll "Anmeldung: meldet sich an(kissdata.mdb)."
        schreibeProtokollBENUTZERablauf "Anmeldung"
        
        If FileExists(cPfadA & "kissapp.mdb") Then
            sErrDabapfad = cPfadA & "kissapp.mdb"
            Set gdApp = OpenDatabase(cPfadA & "kissapp.mdb", False)
            
            schreibeProtokoll "Anmeldung: meldet sich an(kissapp.mdb)."
        Else
        
            MsgBox "Die KISSAPP.MDB wird nicht gefunden.", vbInformation, "Winkiss Hinweis:"
            End 'ende
        End If
    Else
 

    End If
    
    Dim sTabc As String
    sTabc = kassetabcheck(gdBase, Label2, Label3)

    If sTabc = "" Then

    Else
        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden. Winkiss wird beendet.", vbInformation, "Winkiss Hinweis:"
        End
    End If
    
    iStep = 18
    txtStatus.Text = iStep * 2
    
    
    
    '********************************************************************
    '* Hat die Artikel-Datenbank mehr als 100 Sätze, kann die DEMO
    '* nicht gestartet werden
    '********************************************************************
    iStep = 19
    txtStatus.Text = iStep * 2
    If gbDEMO Then
        lAnz = fnPruefeAnzahlArtikelWKL00()
        If lAnz > 200 Then
            ctmp = "Die Datenbank enthält mehr als 200 Artikel." & vbCrLf & vbCrLf
            ctmp = ctmp & "Für die freie Version von Winkiss sind maximal 200 Artikel zugelassen." & vbCrLf & vbCrLf
            ctmp = ctmp & "Das Programm kann nicht gestartet werden!"
            MsgBox ctmp, vbCritical, gsPname & " Hinweis:"
            Command12_Click 4
        End If
    End If
    
    
    
    
    
    
    'Welche Filialnummer bin ich?
    
    gcFilNr = "-1"

    Set rsrs = gdBase.OpenRecordset("Fila", dbOpenTable)
    If Not rsrs.EOF Then
        gcFilNr = rsrs!fil
    End If
    rsrs.Close: Set rsrs = Nothing

    gbFilNr = False
    If Val(gcFilNr) > -1 Then
        gbFilNr = True
    End If
    
    Label1(13).Caption = "F " & gcFilNr
    
    '********************************************************************
    '* Korrektur- und Update-Routinen bei Versionswechsel
    '********************************************************************
    
    Label2.Caption = "Überprüfe Datenbank-Strukturen..."
    Label2.Refresh
    
    
    iStep = 20
    txtStatus.Text = iStep * 2
    CheckProgrammVersion
    
    iStep = 21
    txtStatus.Text = iStep * 2
    LeseProgrammeinstellungenWKL00
    
    
    If gbNoSpruch = False Then
        lbl6(0).Caption = Spruch_des_Tages
        
    End If
    
    
    
    Label2.Caption = "Prüfe auf lokalen Modus..."
    Label2.Refresh
    
    If gbLokalModus = False Then
        iFileNr = FreeFile

        Open gcPfad & "\Lokal.CFG" For Binary As #iFileNr

        If LOF(iFileNr) > 0 Then
            Close iFileNr
            
            If gbLocalSec Then
                If gbAutoLokalModus Then
                    If gbBONWG Then
                        synchronisiereDB
                        HoleLokalDB

'                    ABSCHIEBENDB
                    End If
                Else

                    synchronisiereDB
                    HoleLokalDB
                    
                End If
            End If
            
'            synchronisiereDB
            
            
            Kill gcPfad & "\Lokal.CFG"
        Else
            Close iFileNr
            Kill gcPfad & "\Lokal.CFG"
        End If
    End If

    iStep = 22
    txtStatus.Text = iStep * 2
''    checkdatum
    
    iStep = 23
    txtStatus.Text = iStep * 2
    
''    'DB SIZE
''    Label1(10).Caption = DabaFileSize
''    Label1(10).Refresh
''
''    Label1(11).Caption = DatumLastKompAnzeigen
''    Label1(11).Caption = Label1(11).Caption & " " & DatumLastKompZeitAnzeigen
    
''    Label1(11).Refresh
    
    If gbSchreibRechnerProto Then
        SchreibeWKVersionen
        SchreibRechnerProtokoll
    End If
    
    iStep = 24
    txtStatus.Text = iStep * 2
    Label2.Caption = "Farbeinstellungen werden gesetzt..."
    Label2.Refresh
    
    If gbLokalModus = True Then
        If gbLocalSec Then
            If gbAutoLokalModus = False Then
                schreibelokalFehlerproto "____________________________________"
                schreibelokalFehlerproto "Fehlernummer: 9" & CStr(giUmleitgrund) & vbCrLf & "lokaler Modus "
            End If
        End If
    End If
    
    Me.Refresh
    Modul6.Farbform Me, Me.Label1(0)
    Me.Refresh
    iStep = 25
    txtStatus.Text = iStep * 2
    Label2.Caption = "Schriftart wird gesetzt..."
    Label2.Refresh
    
    
    Modul6.Schrift Me
    Me.Refresh
    
    If gbLokalModus = False Then
    
        If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        
            giUmleitgrund = 5 'Komp fehler
        
            gcUmleittxt = "Beim Komprimieren der Datenbank ist ein schwerwiegender Fehler aufgetreten!" & vbCrLf
            
            
            frmWKL60.Show 1
           
        End If
    
    End If
    
    
    gbBestandsgrund = True
    '**********************************************************
    '    GBZeitschlossVersion = ermdatSAP
    
'    CheckThis
    
    If ermdatSAP = True Then
        If Day(DateValue(Now)) > 13 Then
            gsAnzeigeText = "Geben Sie Passwort Nr. " & Month(DateValue(Now)) & " ein!"
        
        ElseIf Day(DateValue(Now)) > 10 Then
            gsAnzeigeText = "Ihre Programmversion läuft demnächst ab." & vbCrLf
            gsAnzeigeText = gsAnzeigeText & "Passwort Nr. " & Month(DateValue(Now)) & " wird dann benötigt."
        Else
            gsAnzeigeText = "Sie verwenden eine zeitlich begrenzte Programmversion."
        End If
        
        gsZeitschlossdate = DateValue(Now)
        Screen.MousePointer = 0
        frmWKL69.Show 1
        
        If Day(DateValue(Now)) > 13 Then
            Select Case Month(DateValue(Now))
                Case 1
                    If gsZeitPass <> "Schlümpfe" Then
                        End
                    End If
                Case 2
                    If gsZeitPass <> "Single" Then
                        End
                    End If
                Case 3
                    If gsZeitPass <> "Ölschock" Then
                        End
                    End If
                Case 4
                    If gsZeitPass <> "Zweierkiste" Then
                        End
                    End If
                Case 5
                    If gsZeitPass <> "Mitte" Then
                        End
                    End If
                Case 6
                    If gsZeitPass <> "Wende" Then
                        End
                    End If
                Case 7
                    If gsZeitPass <> "Havarie" Then
                        End
                    End If
                Case 8
                    If gsZeitPass <> "Waldsterben" Then
                        End
                    End If
                Case 9
                    If gsZeitPass <> "Molkepulver" Then
                        End
                    End If
                Case 10
                    If gsZeitPass <> "Tiefflug" Then
                        End
                    End If
                Case 11
                    If gsZeitPass <> "Realo" Then
                        End
                    End If
                Case 12
                    If gsZeitPass <> "Eurogeld" Then
                        End
                    End If
                    
            End Select
        End If
    End If
    '**********************************************************
    
    iStep = 26
    txtStatus.Text = iStep * 2
    
    Label1(0).Caption = gsPname & Label1(0).Caption
    
    If gbLokalModus = False Then
        If gbSichernHeut And giSICHTYP = 1 Then
            frmWKL00.Label2.Caption = "Datenbank wird gesichert, nicht ausschalten!!!"
            frmWKL00.Label2.Refresh
            
            DabaSicherung
            
            frmWKL00.Label2.Caption = "Anwender aktiv"
            frmWKL00.Label2.Refresh
            
            
            gbSichernHeut = False
        End If
    End If
    
    '********************************************************************
    '* Spezielle Programmeinstellungen laden , Farben + Programmname
    '********************************************************************
    
    
    '********************************************************************
    '* Überprüfung, ob das Programm sauber registriert ist
    '********************************************************************
    
    Label2.Caption = "Anwender nicht aktiv"
    Label2.Refresh
    
    
    iStep = 27
    txtStatus.Text = iStep * 2
    
    glLevel = -1
    
    If Not gbDEMO And Not gbKostenlos Then
        If gbRegister Then
            iRet = fnPruefeRegistrierungWKL00()
        Else
            iRet = 0
        End If
        If iRet <> 0 Then
            If gbDebug Then
                Label2_MouseUp 2, 1, 0, 0
            End If
                
            ctmp = "Sie verwenden eine nicht- oder falschregistrierte Programmversion!"
            ctmp = ctmp & vbCrLf & vbCrLf
            ctmp = ctmp & "Bitte setzen Sie sich mit der KISS-Hotline Tel.:(0511) 95 59 110 in Verbindung!"
            ctmp = ctmp & vbCrLf & vbCrLf
            ctmp = ctmp & "Das Programm wird jetzt beendet!"
            MsgBox ctmp, vbInformation, "WICHTIGER HINWEIS"
            
            
'            Kill gcSysPfad & gcRegDatei

            Unload frmWKL00
            End 'Ende
        Else
            If gbDebug Then
                Label2_MouseUp 2, 1, 0, 0
            End If
            If Not gbRegister Then
                LadeUnternehmensDatenWKL00
            End If
            ctmp = Trim$(gRegister.firma)
            If InStr(ctmp, " & ") > 0 Then
                ctmp = Left(ctmp, InStr(ctmp, " & ") + 1) & "&" & Right(ctmp, Len(ctmp) - InStr(ctmp, " & ") - 1)
            End If
            Label1(1).Caption = "Registriert für: " & ctmp & ", " & Trim$(gRegister.Plz) & " " & Trim$(gRegister.Ort)
            Label1(1).Refresh
        End If
        Screen.MousePointer = 11
    End If
    
    If gbKostenlos Then
        If NewTableSuchenDBKombi("FREE", gdBase) = False Then
            frmWKL164.Show 1
        End If
    End If
    
    '********************************************************************
    '* Computername geht an die gemeinsame Datenbank
    '* Aktionen wie Kassenabschluss,Programmupdate uvm. werden hier geregelt
    '*
    '********************************************************************
    iStep = 28
    txtStatus.Text = iStep * 2
    
    Label2.Caption = "Setze Listendrucker als Standard..."
    Label2.Refresh
    
    iStep = 29
    txtStatus.Text = iStep * 2
    iRet = fnLeseIniPrinterWKL00()
    Do While iRet <> 0
        frmWKL50.Show 1
        iRet = fnLeseIniPrinterWKL00()
    Loop
    
    iStep = 30
    txtStatus.Text = iStep * 2
    setzedrucker gcListenDrucker
    
'    If Modul6.FindFile(App.Path, "\Handbuch.pdf") Then
'        Label1(7).Visible = True
'    Else
'        Label1(7).Visible = False
'    End If

    iStep = 31
    txtStatus.Text = iStep * 2
    AnmeldungDabaNew
    
    '********************************************************************
    '* Anwender muß sich mit Name und Passwort anmelden
    '* Ab diesem Zeitpunkt greifen die Zugriffsrechte auf die einzelnen
    '* Dialog für den Anwender
    '********************************************************************
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "LPROTOK\Start.txt") Then
    
        'Starte durch
        gcUserName = gcMASTERUSER
        gcPass = gcMASTER
        gcBedienerNr = "99"
        glLevel = 9
        frmWKL00!Label2.Visible = True
        frmWKL00!Label2.Caption = gcUserName & " angemeldet"
        frmWKL00!Label2.Refresh
        
        UpdateUSERSAFE gcBedienerNr, gcUserName
        
        Kill cPfad & "LPROTOK\Start.txt"
    Else
        If gbBEDKARTE Then
            fAnmeldung
        Else
            Do
                frmWKL99.Show 1
                If glLevel < 0 Then
                    iRet = MsgBox("Anmeldung fehlerhaft!" & vbCrLf & "Erneuter Versuch?", vbQuestion + vbYesNo, gsPname & " Anmeldung:")
                    If iRet = vbNo Then
                        Unload frmWKL00
                        End 'Ende
                    End If
                End If
            Loop While glLevel = -1
        End If
    End If
    


    'check ob im App pfad  Artikel100Top.zip vorliegt
    
    
    'check ob im DABapfad  Artikel100Top.zip vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    ctmp = "Die GDPdU-Datenbank fehlt, Winkiss wird beendet." & vbCrLf & vbCrLf
    ctmp = ctmp & "GDPdU: Grundsätze zum Datenzugriff und zur Prüfbarkeit digitaler Unterlagen" & vbCrLf & vbCrLf
    ctmp = ctmp & "Auf Verlangen der Finanzaufsicht wird dieses unveränderbare und maschinenlesbare Archiv benötigt." & vbCrLf & vbCrLf
    ctmp = ctmp & "Winkiss wird jetzt beendet."
    
    If FileExists(cPfad & "GDPdU\GDPdU.mdb") = False Then
        MsgBox ctmp, vbCritical, "Winkiss Hinweis:"
        End
    End If
    
    'GDPDUkltmp
    
    ctmp = "Fehler bei der letzten Komprimierung, Winkiss wird beendet." & vbCrLf & vbCrLf
    ctmp = ctmp & "Um diesen Fehler zu beheben, rufen Sie die Hotline (0511/9559110) an!" & vbCrLf & vbCrLf
    ctmp = ctmp & "Fehlerhinweis: Es existiert eine 'GDPDUkltmp.mdb' unter: " & cPfad & "GDPdU" & vbCrLf & vbCrLf
    ctmp = ctmp & "Winkiss wird jetzt beendet."

    If FileExists(cPfad & "GDPdU\GDPDUkltmp.mdb") = True Then
        MsgBox ctmp, vbCritical, "Winkiss Hinweis:"
        End
    End If
    
    'KASSBONkltmp
    
    ctmp = "Fehler bei der letzten Komprimierung, Winkiss wird beendet." & vbCrLf & vbCrLf
    ctmp = ctmp & "Um diesen Fehler zu beheben, rufen Sie die Hotline (0511/9559110) an!" & vbCrLf & vbCrLf
    ctmp = ctmp & "Fehlerhinweis: Es existiert eine 'KASSBONkltmp.mdb' unter: " & cPfad & "GDPdU" & vbCrLf & vbCrLf
    ctmp = ctmp & "Winkiss wird jetzt beendet."

    If FileExists(cPfad & "GDPdU\KASSBONkltmp.mdb") = True Then
        MsgBox ctmp, vbCritical, "Winkiss Hinweis:"
        End
    End If
    
    'kltmp
    Dim cPfadApp As String
    cPfadApp = gcPfad    'dabapfad
    If Right(cPfadApp, 1) <> "\" Then
        cPfadApp = cPfadApp & "\"
    End If
    
    ctmp = "Fehler bei der letzten Komprimierung, Winkiss wird beendet." & vbCrLf & vbCrLf
    ctmp = ctmp & "Um diesen Fehler zu beheben, rufen Sie die Hotline (0511/9559110) an!" & vbCrLf & vbCrLf
    ctmp = ctmp & "Fehlerhinweis: Es existiert eine 'kltmp.mdb'(kissapp) unter: " & cPfadApp & vbCrLf & vbCrLf
    ctmp = ctmp & "Winkiss wird jetzt beendet."

    If FileExists(cPfadApp & "kltmp.mdb") = True Then
        MsgBox ctmp, vbCritical, "Winkiss Hinweis:"
        End
    End If
    
   
    
    
    

''    If FileExists(cPfad & "IN\ArtikelAbgleich.mdb") Then
''        Dim dbSHOPART As Database
''
''        cPfadQ = gcDBPfad    'dabapfad
''        If Right(cPfadQ, 1) <> "\" Then
''            cPfadQ = cPfadQ & "\"
''        End If
''
''        Set dbSHOPART = OpenDatabase(cPfad & "IN\ArtikelAbgleich.mdb", False, False)
''
''        loeschNEW "leb_Artikel", gdBase
''        TransferTab dbSHOPART, cPfad & "kissdata.mdb", "leb_Artikel"
''
''        loeschNEW "leb_Artlief", gdBase
''        TransferTab dbSHOPART, cPfad & "kissdata.mdb", "leb_Artlief"
''
''
''        sSQL = "Create index Artnr on leb_artikel (artnr)"
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Update artikel set synstatus = 'D' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Update artikel a inner join leb_artikel l on a.artnr = l.artnr  "
''        sSQL = sSQL & " set a.synstatus = 'E' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Delete from artikel where synstatus = 'D' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Create index Artnr on leb_artlief (artnr)"
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Create index linr on leb_artlief (linr)"
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Update artlief set synstatus = 'D' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Update artlief a inner join leb_artlief l on a.artnr = l.artnr and a.linr = l.linr "
''        sSQL = sSQL & " set a.synstatus = 'E' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        sSQL = "Delete from artlief where synstatus = 'D' "
''        gdBase.Execute sSQL, dbFailOnError
''
''        dbSHOPART.Close
''
''        Kill cPfad & "IN\ArtikelAbgleich.mdb"
''    End If









    systembildcheck_all
    
    waehrungbildcheck_all

    
    'check im IN ob KeinBildm.jpg vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "KeinBildm.jpg") Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "KeinBildm.jpg"

        cZiel = cPfad & "\PICTURE\Kunden\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "KeinBildm.jpg"

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    'check im IN ob KeinBildw.jpg vorliegt

    'check im IN ob KeinBildw.jpg vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "KeinBildw.jpg") Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "KeinBildw.jpg"

        cZiel = cPfad & "\PICTURE\Kunden\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "KeinBildw.jpg"

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    'check im IN ob KeinBildw.jpg vorliegt
    
    
    'check im IN ob KeinBild.jpg vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "KeinBild.jpg") Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "KeinBild.jpg"

        cZiel = cPfad & "\PICTURE\Artikel\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "KeinBild.jpg"

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    'check im IN ob KeinBild.jpg vorliegt
    
    'check im IN ob neue Lizenz ZNEZIL.cfg vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "IN\ZNEZIL.cfg") Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "IN\ZNEZIL.cfg"

        cZiel = cPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "ZNEZIL.cfg"

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    'ende check im IN ob neue Lizenz ZNEZIL.cfg vorliegt
    
    'check im IN ob neue Lizenz ZNEZILINDI.cfg vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If FileExists(cPfad & "IN\ZNEZILINDI.cfg") Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "IN\ZNEZILINDI.cfg"

        cZiel = cPfad
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & "ZNEZILINDI.cfg"

        lRet = CopyFile(cQuelle, cZiel, lfail)
        Kill cQuelle
    End If
    'ende check im IN ob neue Lizenz ZNEZIL.cfg vorliegt





'''    ' GFK Wochenauswertung rausschicken
'''    Dim cAKTKW          As String 'Kalenderwoche
'''    Dim cGESpKW         As String
'''    Dim DateHeut        As Date
'''    Dim DateGespeich    As Date
'''
'''    If gbUnistatWeek Then 'Teilnahme an Wochenauswertung
'''        If gbStatweekperMail = True Then
'''            If gbDSL = True Then
'''                DateGespeich = DateValue(gdateStatlast): DateHeut = DateValue(Now): cAKTKW = DatePart("ww", DateHeut, vbMonday)
'''                If CInt(cAKTKW) = 1 Then
'''                    cAKTKW = "53"
'''                Else
'''                    cAKTKW = CInt(cAKTKW) - 1
'''                End If
'''
'''                cGESpKW = DatePart("ww", DateGespeich)
'''
'''                If CInt(cGESpKW) = 1 Then
'''                    cGESpKW = "53"
'''                Else
'''                    cGESpKW = CInt(cGESpKW) - 1
'''                End If
'''
'''                If CInt(cAKTKW) <> CInt(cGESpKW) Then 'Vergleich
'''                    If Trim(gsStatkundnr) = "" Then gsStatkundnr = "XXX" 'Kisskundennummer?
'''        '            If unistatweek Then Label1.Caption = DatumLastSuniW:
'''
'''                    If unistatweek_new(frmWKL00.txtStatus, frmWKL00.picprogress) Then lbl6(28).Caption = DatumLastSuniW:
'''
'''                End If
'''            End If
'''        Else
'''
'''        End If
'''    End If
'''    'Ende GFK Wochenauswertung rausschicken
'''
    
    
    
    
    'Coupon-Auswertung verschicken
    
    'gibt es eine BUDNINR
    Dim cBudniKundNr    As String
    Dim rsLi            As DAO.Recordset
    Dim slastAuswertTag As String
    Dim sBisAuswertTAG  As String
    Dim lminAuswerttag  As Long
    Dim lmaxAuswerttag  As Long
    Dim l               As Long

    cBudniKundNr = ""
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDIBUDNI' "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        cBudniKundNr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing

    If Val(cBudniKundNr) > 0 Then
        If NewTableSuchenDBKombi("COUPONSTAT", gdBase) = False Then
            CreateTableT2 "COUPONSTAT", gdBase

            slastAuswertTag = DateValue("31.01.2017")

            sSQL = "Insert into COUPONSTAT (LASTDATE) values ('" & slastAuswertTag & "')"
            gdBase.Execute sSQL, dbFailOnError
        End If

        slastAuswertTag = leseCouponStat("lastdate")

        If slastAuswertTag = "" Then
            slastAuswertTag = DateValue("31.01.2017")
        End If

        sBisAuswertTAG = DateValue(Now) - 1

        If CLng(DateValue(slastAuswertTag)) < CLng(DateValue(sBisAuswertTAG)) Then
            lminAuswerttag = CLng(DateValue(slastAuswertTag)) + 1
            lmaxAuswerttag = CLng(DateValue(sBisAuswertTAG))
            For l = lminAuswerttag To lmaxAuswerttag
                Couponeinloesung cBudniKundNr, l
            Next l

            COUPON_AUSW_uebertragen

            sSQL = "Update COUPONSTAT Set LASTDATE = " & lmaxAuswerttag & " "
            gdBase.Execute sSQL, dbFailOnError

        End If

    End If
    
    'Ende Coupon-Auswertung verschicken
    
    
    
    


''    'check im IN ob Banken.mdb vorliegt
''    Dim dbBanken As Database
''
''    cPfadQ = gcDBPfad    'dabapfad
''    If Right(cPfadQ, 1) <> "\" Then
''        cPfadQ = cPfadQ & "\"
''    End If
''
''    If FileExists(cPfadQ & "IN\Banken.mdb") Then
''
''        Set dbBanken = OpenDatabase(cPfadQ & "IN\Banken.mdb", False, False)
''
''        loeschNEW "BANKEN", gdBase
''
''        TransferTab dbBanken, cPfadQ & "kissdata.mdb", "BANKEN"
''        dbBanken.Close
''        Kill cPfadQ & "IN\Banken.mdb"
''
''    End If
''    'ende check im ZIN ob Banken.mdb vorliegt
    
    
    If gbGTBON = True Then
        If GTBONHEUTENOCHNICHT Then
            DruckenGutenTagBon
        End If
    End If
    
    '********************************************************************
    '* Liegt im DBPfad\IN noch ein Programm-Update?
    '********************************************************************

    iStep = 32
    txtStatus.Text = iStep * 2

    iRet = NEWfnCheck4UpdateDateiWKL00(True)
    If iRet <> 0 Then
    
        Dim lmerkeFS As Long
        Dim lmerkeH As Long
        Dim lmerketop As Long
        
        lmerkeFS = Label2.FontSize
        lmerketop = Label2.Top
        lmerkeH = Label2.Height
    
        Label2.FontSize = 42
        Label2.Height = lmerkeH + 500
        Label2.Top = lmerketop - 1000
        
        anzeige "rot2", "Programmupdate wird eingelesen...", Label2

        Label2.FontSize = lmerkeFS
        Label2.Height = lmerkeH
        Label2.Top = lmerketop
        
        
        Deaktiviere_alleSchaltflächen
        
        frmWKL53.cmdUpdEinlesen_Click
        anzeige "", "", Label2
        
    Else
        iStep = 33
        txtStatus.Text = iStep * 2
    End If
    gcPfad = App.Path
    If Right(gcPfad, 1) <> "\" Then
        gcPfad = gcPfad & "\"
    End If

    If gbLokalModus Then
    
        Command1(0).Enabled = False
        Command1(1).Enabled = True
        Command1(2).Enabled = False
        Command1(3).Enabled = False
        Command1(4).Enabled = False
        Command1(8).Enabled = False
        
        Command3(10).Enabled = False
        Command3(9).Enabled = False
        Command3(8).Enabled = False
        Command3(7).Enabled = False
        Command3(4).Enabled = False
        Command3(5).Enabled = False
        
        anzeige "Artikel", "", Label3
        anzeige "Artikel", "", Label3
        anzeige "Artikel", "", Label3
        anzeige "Artikel", "", Label3
        anzeige "Artikel", "", Label3
        
    Else
    
        Command1(0).Enabled = True
        Command1(1).Enabled = True
        Command1(2).Enabled = True
        Command1(3).Enabled = True
        Command1(4).Enabled = True
        Command1(8).Enabled = True
        
    End If
    
    If ermaktUmsatz(False) > 0 Then
        Command3(1).BackColor = vbRed
    Else
        Command3(1).BackColor = &H8000000F
    End If
    
    iStep = 34
    txtStatus.Text = iStep * 2
    
    If gbLocalSec Then
        If gbAutoLokalModus Then
            anzeige "normal", "Offline - Betrieb", Label2
            
            AbmeldungDabaNew
            ChkLM.value = vbChecked
        End If
    End If
    
    
    

    
    '********************************************************************
    '* Setzen verschiedener Konstanten:
    '* - MWSt
    '* - Wochentage
    '* - Monatsnamen
    '* - Zahlungsarten
    '********************************************************************
    
    iStep = 35
    txtStatus.Text = iStep * 2
    
    LeseMWStSaetzeWKL00
    
    leseKundendurchscnittswerte
    
    
    
    Timer1.Interval = 1000
    Timer2.Interval = 1000
    
    If gbREGEB = True Then
    
        ' ist es Sonntag dann fragen wir nicht. Grund: LCN Roth lässt danach automatisch die SMS-Erinnerungen raus und dann steht die MSGBOX im wege
        If Weekday(Now, vbMonday) <> "7" Then
            WerHatheuteGeburtstag giGebTage
        End If
    End If
    
    If gbTerminReminderSMS = True Then
    
        If gsTerminReminderstart = "" Then
            Termine_versenden_Frage
        End If
        
    End If
    
    glfarbe(0) = vbWhite  '&H8000000F
    glfarbe(1) = &HFFFF&
    
    glfarbe(2) = &HC000&
    glfarbe(3) = &HFF&
    glfarbe(4) = &HC0FFFF
    glfarbe(5) = &H80FF&
    glfarbe(6) = &HFF00FF
    glfarbe(7) = &HFFFF00
    glfarbe(8) = &HC0C0FF
    glfarbe(9) = &HFFC0C0
    
    glfarbe2(1) = &H8080FF
    glfarbe2(2) = &HC0FFC0
    glfarbe2(3) = &HFF8080
    glfarbe2(4) = &H40C0&
    glfarbe2(5) = &H800080
    glfarbe2(6) = &H80&
    glfarbe2(7) = &H808000
    glfarbe2(8) = &HC0C0&
    glfarbe2(9) = &HFF80FF
    
    gcWochentag(0) = ""
    gcWochentag(1) = "Montag"
    gcWochentag(2) = "Dienstag"
    gcWochentag(3) = "Mittwoch"
    gcWochentag(4) = "Donnerstag"
    gcWochentag(5) = "Freitag"
    gcWochentag(6) = "Samstag"
    gcWochentag(7) = "Sonntag"
    
    gcMonat(1) = "Januar"
    gcMonat(2) = "Februar"
    gcMonat(3) = "März"
    gcMonat(4) = "April"
    gcMonat(5) = "Mai"
    gcMonat(6) = "Juni"
    gcMonat(7) = "Juli"
    gcMonat(8) = "August"
    gcMonat(9) = "September"
    gcMonat(10) = "Oktober"
    gcMonat(11) = "November"
    gcMonat(12) = "Dezember"
    
    gcZahlArt(1) = "EC-Lastschrift:"
    gcZahlArt(5) = "Kredit:"
    gcZahlArt(6) = "Scheck:"
    gcZahlArt(8) = "Bar:"
    gcZahlArt(45) = "Ausz.:"
    gcZahlArt(46) = "Einz.:"
    gcZahlArt(17) = "Karte:"
    gcZahlArt(47) = "Kollege:"
    
    gcTag = Weekday(Now, vbMonday)
    
   
'    bestellvorschlagrechnen
    
    '********************************************************************
    '* Fülle Array Warengruppentasten mit Leerwerten
    '********************************************************************
        
    For lRet = 18 To 29
        gWarenGruppe(lRet).lWgNr = lRet - 17
        gWarenGruppe(lRet).dArtNr = -1
        gWarenGruppe(lRet).cBezeich = "WG " & Trim$(Str$(gWarenGruppe(lRet).lWgNr))
        gWarenGruppe(lRet).cFaktor = "+"
    Next lRet
    
    For lRet = 120 To 171
        gWarenGruppe(lRet).lWgNr = lRet - 107
        gWarenGruppe(lRet).dArtNr = -1
        gWarenGruppe(lRet).cBezeich = "WG " & Trim$(Str$(gWarenGruppe(lRet).lWgNr))
        gWarenGruppe(lRet).cFaktor = "+"
    Next lRet
    
    '********************************************************************
    '* Lese Kassenbontexte: 3 Kopfzeilen, 3 Fußzeilen
    '********************************************************************
    
    iStep = 36
    txtStatus.Text = iStep * 2
    LeseTexteKassenBonWKL00
    
    '********************************************************************
    '* Lese Zugriffsrechte für die einzelnen Dialoge
    '********************************************************************
    
'    Label2.Caption = "Zugriffsrechte werden ermittelt..."
'    Label2.Refresh

    iStep = 37
    txtStatus.Text = iStep * 2
    LeseZugriffsRechte
    
    '********************************************************************
    '* Lese Firmendaten (Name, Anschrift, ILN-Nummern)
    '********************************************************************
    
    iStep = 38
    txtStatus.Text = iStep * 2
    LeseFirmenDaten
    
    Do While gFirma.FirmaName = ""
        frmWKL16.Show 1
        LeseFirmenDaten
        End
    Loop
    
    '********************************************************************
    '* Setze Verbindungswerte für Betrieb eines Bondruckers über COM-Port
    '********************************************************************
'    MsgBox "12"

    gVerbindung.lBaudRate = 9600
    gVerbindung.sStopBits = 1
    gVerbindung.iDatenBits = 8
    gVerbindung.cParitaet = "E"
    gVerbindung.iComPort = 2
    gVerbindung.cSettings = "9600,E,8,1"
    
    '********************************************************************
    '* Lese Verbindungsdaten aus cfg-Datei
    '********************************************************************
    
    iStep = 39
    txtStatus.Text = iStep * 2
    LeseDatenVerbindung
    If gbLokalModus = False Then
        If gbDabakompfrueh Then
            zwangsoptimierung
        End If
    End If
    
    '*********************************************************
    '* gibt es Preis, die auf Termin geändert werden müssen?
    '*********************************************************
    
    iStep = 40
    txtStatus.Text = iStep * 2
    
    If gbLokalModus = False Then
        If gbSTADAP Then
            Check4PrsTerminWKL00
        End If
    End If
    
    If gbKUBONUS = False Then
        Tages_Bonus_zurückstellen
    End If
    

    
     If DatendrinSQL("select * from ARTAUSWAHL where art = 'Verleih' and Adate < clng(datevalue(now)-14)", gdBase) Then
        MsgBox "Eine oder mehrere Artikelauswahlen sind älter als 14 Tage.", vbOKOnly + vbInformation, "Winkiss Hinweis:"
    End If
    
    giAnzFil = 0
    '** Anzahl der vorhandene FILIALE(n) holen **
    cSQL = "Select count(*) as ANZ from FILIALEN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!anz) Then
            giAnzFil = rsrs!anz
        End If
    End If
    ReDim giFilNrS(giAnzFil)
    rsrs.Close: Set rsrs = Nothing
    
    j = 1
    cSQL = "Select * from FILIALEN order by filialnr "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
                giFilNrS(j) = rsrs!FILIALNR
                j = j + 1
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    
    
    
    '********************************************************************
    '* Ist der Stat - Ordner voll? dann Abbruch und Telekissstart
    '********************************************************************
    
    iStep = 41
    txtStatus.Text = iStep * 2
    VerzVorhanden "Stat", gcDBPfad & "\"
    File2.Path = gcDBPfad & "\STAT"
    File2.Refresh

    If File2.ListCount > 0 Then
        If gbFtpYes Then
            giKissFtpMode = 5 'FTPMODE= 5 , STAT - Ordner leeren abschicken
            frmWKL38.Show 1
'        Else
'            Screen.MousePointer = 0
'            frmWK21k.Show 1
        End If
    End If

   
    iStep = 42
    txtStatus.Text = iStep * 2
    '****** FTP - Server nach Programmupdates und Stammdaten überprüfen
    Dim bseekkasdat As Boolean
    Dim bmerke As Boolean
    bmerke = gbFTPautomatic
    
    bseekkasdat = False
    
    DateNow = DateValue(Now) 'Vergleichsdatum vom Rechner
    If Not gbLokalModus Then
        If gbFtpYes Then
            Select Case giStammFTPOFT 'Wie oft
                Case Is = 0                                         'FTP täglich
                    
                    iDayUnterschied = DateNow - gdateStammlastFTP
                    
                    If giStammFTPOFT < iDayUnterschied Then
                    
                        If gbNacht = False Then
                            If gbFtpZENT Then
                                If Val(gcFilNr) > 0 Then
                                    gbFTPautomatic = True
                                    giKissFtpMode = 10 'FTPMODE= 10 , Kombimodus Kassendateien holen und schicken
                                    frmWKL38.Show 1    ' Programmupdates,Stammdaten holen
                                    gbFTPautomatic = bmerke

                                    bseekkasdat = True
                                End If
                            End If
                        Else
                        
                        
                            If gbFtpZENT Then
                                If Val(gcFilNr) > 0 Then
                                    gbFTPautomatic = True
                                    giKissFtpMode = 10 'FTPMODE= 10 , Kombimodus Kassendateien holen und schicken
                                    frmWKL38.Show 1    ' Programmupdates,Stammdaten holen
                                    gbFTPautomatic = bmerke

                                    bseekkasdat = True
                                End If
                            Else
                                gbFTPautomatic = True
                                giKissFtpMode = 1 'FTPMODE= 1 , Programmupdates und Stammdaten holen
                                frmWKL38.Show 1
                                gbFTPautomatic = bmerke
                            
                            End If
                        End If
                    End If
                Case Is = 1                                         'FTP alle 3Tage
                    iDayUnterschied = DateNow - gdateStammlastFTP
                    If giStammFTPOFT < iDayUnterschied Then FTPprüfung
                Case Is = 2                                         'FTP alle 7 Tage
                    iDayUnterschied = DateNow - gdateStammlastFTP
                    If giStammFTPOFT < iDayUnterschied Then FTPprüfung
            End Select
        Else
            If gbSTADAP Then
                '****** Stammdaten prüfen normal
                iRet = fnCheck4MasterDateiWKL00()
                If iRet <> 0 Then
                    ctmp = "Achtung!" & vbCrLf & vbCrLf
                    ctmp = ctmp & "Sie haben neue Stammdaten in Ihrem Verzeichnis." & vbCrLf
                    ctmp = ctmp & "Möchten Sie diese jetzt einlesen?"
                    iRet = MsgBox(ctmp, vbYesNo + vbInformation, "Winkiss - neue Stammdaten:")
                End If
                If iRet = vbYes Then
                    If gbZugriffNew Then
                        byteZGNr = ermittleTag(Command4(7))
                        OpenProgrammTeil frmWKL11, ermittlezugriff(byteZGNr)
                    Else
                        If glLevel >= DlgZugriff(1).dZugriff Then
                            frmWKL11.Show 1
                        Else
                            MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    '************************************************
    '* Liegt im DBPfad\IN noch eine Kassen-Datei?
    '************************************************
    
    iStep = 43
    txtStatus.Text = iStep * 2
    DoEvents
    
    iRet = newfnCheck4KassenDateiWKL00()
    
    If gbLokalModus = False Then
    
        'Nach BudniStammdaten
        If gbSTADAP Then
            iRet = newfnCheck4BudniStammdaten()
        End If
        
        'Nach ReweStammdaten
        If gbSTADAP Then
            iRet = newfnCheck4ReweStammdaten()
        End If
        
        'Nach StreckenStammdaten
        If gbSTADAP Then
            iRet = newfnCheck4StreckenStammdaten()
        End If
        
        'Nach Coupon-daten
        If gbSTADAP Then
            iRet = newfnCheck4CouponDaten()
        End If

        'Nach LueningStammdaten
        
        
        If Val(gcFilNr) = 0 Then
            If gbSTADAP Then
            
                'check doch mal ob es Lüning gibt
                'wenn ja dann check mal ob tägliche Stammdaten vorliegen
                check_Lüning_Stammdaten
            
                iRet = newfnCheck4LueningStammdaten()
            End If
        End If
    End If
    
    
    
    
    If NewTableSuchenDBKombi("PICKLISTE_IN", gdBase) Then
    
        If Datendrin("PICKLISTE_IN", gdBase) Then
        
            ctmp = "Achtung!" & vbCrLf & vbCrLf
            ctmp = ctmp & "Sie haben neue Picklisten in Ihrem Verzeichnis." & vbCrLf
            ctmp = ctmp & "Möchten Sie diese jetzt auf dem Bondrucker ausdrucken?"
            iRet = MsgBox(ctmp, vbYesNo + vbInformation, "Winkiss - neue Picklisten:")
                    
            If iRet = vbYes Then
                Drucke_Picklisten
                loeschNEW "PICKLISTE_IN", gdBase
            End If
        
        Else
            loeschNEW "PICKLISTE_IN", gdBase
        End If
        
    End If
    
    
    
    
    
    If gbFTH = True Then
        fnCheck_Filialtausch "N*.mdb"
        fnCheck_Filialtausch "WV*.mdb"
    End If
    
    If Val(gcKasNum) > 0 And Val(gcKasNum) < 10 Then
        Command1(9).BackColorTo = glfarbe(gcKasNum)
        Command1(9).BackColorFrom = glfarbe(gcKasNum)
    End If
    
    Command1(9).Caption = gcKasNum
    If gbEDITKASSNR = False Then
        Command1(9).Enabled = False
    Else
        Command1(9).Enabled = True
    End If
    
    
    
    'wenn Bestandlive, dann hier starten
    '    If gsEPartner = "ZVT" Then
    
    
    
    
    If gbBestDateien = True And gsPfadBestandlive <> "" Then
    


        Dim bBLIvefound As Boolean
'        lese_ZVT_opt

        bBLIvefound = False

        'close anwendung
        Dim hwnd&
        Dim Y As String
        Dim result&
        Dim Title$

        Y = "BestandLive" 'gZVTPTitel

        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)

        Do
            result = GetWindowTextLength(hwnd) + 1
            Title = Space(result)
            result = GetWindowText(hwnd, Title, result)
            Title = Left$(Title, Len(Title) - 1)

            If InStr(1, Title, Y) Then
                bBLIvefound = True
                Exit Do
            End If

            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        Loop Until hwnd = 0

        If bBLIvefound = False Then
            'Starte anwendung
            Dim prev_dir As String
            ' Save the current directory.
            prev_dir = CurDir
            ' Go to the desired startup directory.
            ChDrive Left(gsPfadBestandlive, 1)
            ChDir gsPfadBestandlive
            ' Shell the application.
            lRet = Shell("BestandLive.exe", vbNormalFocus)
            ' Restore the saved directory.
            ChDir prev_dir
        End If
    End If
    
    
    
    
    
    Timer1.Interval = 1000
    Timer2.Interval = 1000
    
    schreibeProtokollNachtAblauf "Winkiss wird mit diesen Parametern gestartet:"
    schreibeProtokollNachtAblauf "Kassendateien sofort versenden = " & gbKSF
    schreibeProtokollNachtAblauf "Uhrzeit für den Tagesabschluss = " & gsKassDatstart
    
    If gbLokalModus = False Then
        AlleZugriffeLöschen
    End If
    
    If FileExists(App.Path & "\NoTimer.cfg") Then
        Timer1.Enabled = False
    End If
    
    gbEtiExArtikel = False

 
    Label1(3).Caption = Format$(Time, "HH:MM:SS")
    Label1(3).Refresh
    
    txtStatus.Text = 0
    picprogress.Visible = False
    

   'USB-Stick von TSE Initialisierung    START <------------------------------------------------------
   
         lbl_TSE.Width = 8000
         lbl_TSE.Height = 1500
         lbl_TSE.Left = (Me.ScaleWidth / 2) - (lbl_TSE.Width / 2)
         lbl_TSE.Top = (Me.ScaleHeight / 2) - (lbl_TSE.Height / 2)
         lbl_TSE.BackColor = vbGreen
         
         Check_TSE_Einstellugen
         
         If E_TSE_Aktiv Then
         
                If IstTseInstalliert Then
                     
                      'TSE zum Nuzen vorbereiten
                      TSE_Initialisieren
                     
                 Else
                 
                      iRet = MsgBox("EasyTSE ist nicht installiert, möchten Sie es installieren ?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                      If iRet = vbYes Then
                                  
                             TSE_SilentInstall
                             
                      End If
                     
                       
                End If
          Else
          TSE_Err = "TSE ist deaktiviert"
        End If
        
    'USB-Stick von TSE Initialisierung    ENDE <------------------------------------------------------
   
    
    If Not NewTableSuchenDB("GI_zum_EC_Fertig", gdBase) Then
    
        'GI(giro) auf EC setzen in allen Tabellen,die die Spalte KK_ART enthalten
         If GI_auf_EC_setzenFurAlleTabellenMitKK_ART = 1 Then
          gdBase.Execute "Create Table GI_zum_EC_Fertig(EsIstFertig bit)", dbFailOnError
         End If
     
    End If
    
    'prüf mal, ob Budni auf EDEKA schon geschafft ist (Budni auf EDEKA Umzug ist eine Anforderung, die im August 2021 geschafft wurde)
     BudniFtpUmzug
    
    
 
    'Leere Zeilen nach der Bon-Text <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
    
    'Hier ist eine Erweiterung, mit der die Kunden anpassen können, wie viele leere
    'Zeilen nach dem Bon-Text gedrückt werden sollen
        
            If Not NewTableSuchenDB("LeereZeilen", gdBase) Then
                 
                    gdBase.Execute "Create Table LeereZeilen (ZeilZahl NUMBER)", dbFailOnError
                    gdBase.Execute "INSERT INTO LeereZeilen(ZeilZahl) VALUES (9)", dbFailOnError
                    gbLeereZeil = 9
                    frmWKL52.txtLeereZeil.Text = gbLeereZeil
                    
            Else
                    
                    Dim rsrsleer As Recordset
                    Set rsrsleer = gdBase.OpenRecordset("SELECT ZeilZahl FROM LeereZeilen")
                         
                    If Not rsrsleer.EOF Then
                              
                      If Not IsNull(rsrsleer!ZeilZahl) Then
                      
                         frmWKL52.txtLeereZeil.Text = rsrsleer!ZeilZahl
                         gbLeereZeil = CInt(rsrsleer!ZeilZahl)
                         
                      End If
                              
                    End If
                     
            End If
        
    'Leere Zeilen nach der Bon-Text <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE


     Screen.MousePointer = 0
Exit Sub

LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 76 Then
        bPfad = False
        Resume Next
    ElseIf err.Number = 3024 Then
    
        giUmleitgrund = 1 'Datenbankpfad weg

        gcUmleittxt = "Es kann keine Verbindung zur Datenbank aufgebaut werden!" & vbCrLf
        gcUmleittxt = gcUmleittxt & "Starten Sie die Computer in der richtigen Reihenfolge neu!" & vbCrLf
        gcUmleittxt = gcUmleittxt & "Bleibt der Fehler bestehen, so rufen Sie unsere Kundendienstabteilung (0511 955910) an!" & vbCrLf
        gcUmleittxt = gcUmleittxt & vbCrLf
        gcUmleittxt = gcUmleittxt & "Wir helfen Ihnen gerne weiter."
        
        frmWKL60.Show 1
        
        Command12_Click 4


        iRet = fnLeseIniDateiWKL00()
        Resume
        
    ElseIf err.Number = 3343 Then '
        giUmleitgrund = 2 'Datenbankpfad muß rep
        
        gcUmleittxt = "Ihre Datenbank ist beschädigt. " & sErrDabapfad & vbCrLf
        gcUmleittxt = gcUmleittxt & vbCrLf
        gcUmleittxt = gcUmleittxt & "Stromausfall?,Stromschwankungen? oder den Computer unsachgemäß ausgeschaltet?" & vbCrLf
        
        gcUmleittxt = gcUmleittxt & "Beenden Sie alle anderen Programme auf allen anderen Computern! " & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Brauchen Sie Hilfe, so rufen Sie unsere Kundendienstabteilung (0511 955910) an!" & vbCrLf
        gcUmleittxt = gcUmleittxt & vbCrLf
        gcUmleittxt = gcUmleittxt & "Wir helfen Ihnen gerne weiter."
        
        frmWKL60.Show 1
        
        Resume
    ElseIf err.Number = 3045 Then
        Label1(1).Caption = ""
        Label2.ForeColor = vbYellow
        Label2.Caption = "An einem anderen Rechner wird vermutlich "
        Label2.Refresh
        
        Label3.ForeColor = vbYellow
        Label3.Visible = True
        Label3.Caption = "die Datenbank komprimiert. (eventuell DB-Reparatur erforderlich)"
        Label3.Refresh
        Pause (2)
        Label1(1).ForeColor = vbYellow
        For i = 60 To 1 Step -1
        
        Label1(1).Caption = "Nächster Versuch in:   " & i & " Sekunden "
        Label1(1).Refresh
        Pause (1)
        Next i
        Resume
        
    ElseIf err.Number = 429 Then
    
        giUmleitgrund = 8 'unbekannter Fehler

        gcUmleittxt = "Beim Starten von Winkiss ist ein Fehler aufgetreten!" & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Fehlerbeschreibung:" & vbCrLf
        gcUmleittxt = gcUmleittxt & err.Description & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Fehlernummer: " & err.Number & vbCrLf & vbCrLf
    
        frmWKL60.Show 1

    Else
        giUmleitgrund = 4 'unbekannter Fehler

        gcUmleittxt = "Beim Starten von Winkiss ist ein noch unbekannter Fehler aufgetreten!" & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Fehlerbeschreibung:" & vbCrLf
        gcUmleittxt = gcUmleittxt & err.Description & vbCrLf & vbCrLf
        gcUmleittxt = gcUmleittxt & "Fehlernummer: " & err.Number & vbCrLf & vbCrLf
        
            
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Load"
        Fehler.gsFehlertext = "Beim Starten von Winkiss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
      
'        frmWKL60.Show 1
   

    End If
End Sub

Public Sub BudniFtpUmzug()
 On Error GoTo LOKAL_ERROR
     
    Dim rsLi As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT Distinct FORMAT FROM LISRT WHERE FORMAT='EDIBUDNI' OR FORMAT='EDIBHSG'"
    Set rsLi = gdBase.OpenRecordset(sSQL)
    
     If Not rsLi.EOF Then
      
       If Not NewTableSuchenDB("FTPumzugFertig", gdBase) Then
          If Not NewTableSuchenDB("BudniEdekaDialogNichtZeigen", gdBase) Then
            gbBudniNeuesFtpVerfahren = False
            FTPwechsel.Left = (Me.ScaleWidth - FTPwechsel.Width) / 2
            FTPwechsel.Top = (Me.ScaleHeight - FTPwechsel.Height) / 2
            FTPwechsel.Show 1
          End If
         Else
          gbBudniNeuesFtpVerfahren = True
       End If
     
     End If
    
    rsLi.Close: Set rsLi = Nothing
 
 Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "frmWKL00"
    Fehler.gsFunktion = "BudniFtpUmzug"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Function GI_auf_EC_setzenFurAlleTabellenMitKK_ART() As Integer
On Error GoTo LOKAL_ERROR:

   GI_auf_EC_setzenFurAlleTabellenMitKK_ART = 0

   Dim i As Integer
   Dim lcount As Integer
   
   Dim j As Integer
   Dim colZahl As Integer
    
   Dim tabname As String
   
   Dim Colname As String

   gdBase.TableDefs.Refresh
   i = gdBase.TableDefs.Count
   
   ''''''' durch alle Tabellen schleifen  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
   For lcount = 0 To i - 1
       
       tabname = gdBase.TableDefs(lcount).name
       
       j = gdBase.TableDefs(tabname).Fields.Count
       
       ''''''' durch alle Columns schleifen  <<<< START
        For colZahl = 0 To j - 1
        
          Colname = gdBase.TableDefs(tabname).Fields(colZahl).name
          
          If InStr(1, Colname, "KK_ART") > 0 Then
             'GI auf EC ändern
             gdBase.Execute "UPDATE " & tabname & " SET KK_ART='EC' WHERE KK_ART='GI'", dbFailOnError
             Exit For
          End If
          
        Next colZahl
       ''''''' durch alle Columns schleifen  <<<< START
       
   Next lcount
 ''''''' durch alle Tabellen schleifen  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
 
  GI_auf_EC_setzenFurAlleTabellenMitKK_ART = 1

Exit Function

LOKAL_ERROR:

    GI_auf_EC_setzenFurAlleTabellenMitKK_ART = 0
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "frmWKL00"
    Fehler.gsFunktion = "GI_auf_EC_setzenFurAlleTabellenMitKK_ART"
    Fehler.gsFehlertext = "Im Programmteil Winkiss Starten ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function

