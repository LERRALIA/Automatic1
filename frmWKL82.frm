VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL82 
   BackColor       =   &H0080C0FF&
   Caption         =   "Termin"
   ClientHeight    =   8610
   ClientLeft      =   1710
   ClientTop       =   1845
   ClientWidth     =   11910
   Icon            =   "frmWKL82.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   -120
      TabIndex        =   115
      Top             =   7440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox Check2 
         Caption         =   "So"
         Height          =   195
         Index           =   6
         Left            =   8760
         TabIndex        =   224
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sa"
         Height          =   195
         Index           =   5
         Left            =   8160
         TabIndex        =   223
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fr"
         Height          =   195
         Index           =   4
         Left            =   7560
         TabIndex        =   222
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Do"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   221
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mi"
         Height          =   195
         Index           =   2
         Left            =   6360
         TabIndex        =   220
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Di"
         Height          =   195
         Index           =   1
         Left            =   5760
         TabIndex        =   219
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mo"
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   218
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ComboBox Combo14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         TabIndex        =   213
         Text            =   "Combo1"
         Top             =   180
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox List11 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         ItemData        =   "frmWKL82.frx":0442
         Left            =   5160
         List            =   "frmWKL82.frx":0444
         TabIndex        =   211
         Top             =   3360
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Index           =   5
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   209
         Text            =   "Text1"
         Top             =   5280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox List10 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   208
         Top             =   5520
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox Combo13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   204
         Text            =   "Combo13"
         Top             =   2880
         Visible         =   0   'False
         Width           =   3135
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   16
         Left            =   3480
         TabIndex        =   136
         Top             =   1680
         Width           =   1095
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Achtung"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   135
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   124
         Top             =   2040
         Width           =   1095
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
         Caption         =   "schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   120
         Top             =   240
         Width           =   1095
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   119
         Top             =   960
         Width           =   1095
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
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
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   118
         Top             =   1320
         Width           =   1095
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Info"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   117
         Top             =   600
         Width           =   1095
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorDisabled=   15398133
         BackColorFrom   =   16514300
         BackColorTo     =   15462640
         BackColorCheckedFrom=   15462640
         BackColorCheckedTo=   16514300
         BackColorDownFrom=   12700881
         BackColorDownTo =   15659506
         BackColorHoverFrom=   16514300
         BackColorHoverTo=   15462640
         BorderColor     =   7617536
         BorderColorDisabled=   12240841
         BorderColorFocus=   14986635
         BorderColorHover=   3913721
         ForeColorDisabled=   9609633
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   116
         Top             =   1440
         Width           =   3255
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   206
         Top             =   3360
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3201
         _Version        =   393216
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTPickerVon 
         Height          =   255
         Left            =   6600
         TabIndex        =   214
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   112852993
         CurrentDate     =   38457.8333333333
      End
      Begin MSComCtl2.DTPicker DTPickerBis 
         Height          =   255
         Left            =   8040
         TabIndex        =   215
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   112852993
         CurrentDate     =   38457.8333333333
      End
      Begin MSComctlLib.TreeView Tree11 
         Height          =   1815
         Left            =   5160
         TabIndex        =   216
         Top             =   960
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3201
         _Version        =   393217
         LabelEdit       =   1
         Scroll          =   0   'False
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzahl"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   6600
         TabIndex        =   226
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Bediener"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   7440
         TabIndex        =   225
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "ganzer Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   5160
         MouseIcon       =   "frmWKL82.frx":0446
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   217
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "freie Termine bei:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   5160
         TabIndex        =   212
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "entfernen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   3480
         MouseIcon       =   "frmWKL82.frx":0750
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   210
         Top             =   6480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "gewählte Behandlungen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   207
         Top             =   5280
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Behandlung wählen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   205
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "letzter Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   2040
         MouseIcon       =   "frmWKL82.frx":0A5A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   189
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "wurde bisher bedient:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   123
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Kunde wählen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   121
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00C0C000&
      Caption         =   "weitere Termine zum Kunden"
      Height          =   1335
      Left            =   600
      TabIndex        =   190
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
      Begin VB.ListBox List9 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   120
         TabIndex        =   192
         Top             =   1080
         Width           =   11415
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   7
         Left            =   9720
         TabIndex        =   193
         Top             =   7920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   8
         Left            =   9720
         TabIndex        =   194
         Top             =   7440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lblDatum 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         TabIndex        =   198
         Top             =   4920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblKunde 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         TabIndex        =   197
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Achtung! weitere Termine am gleichen Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   196
         Top             =   480
         Width           =   11415
      End
      Begin VB.Label lblDelgrund 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         TabIndex        =   195
         Top             =   4920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblBed 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         TabIndex        =   191
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0C000&
      Caption         =   "Notizen zum Kunden"
      Height          =   855
      Left            =   10200
      TabIndex        =   181
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   184
         Top             =   360
         Width           =   11295
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   10
         Left            =   9450
         TabIndex        =   182
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   1
         Left            =   10440
         TabIndex        =   183
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Index           =   26
         Left            =   120
         TabIndex        =   185
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame12 
      Height          =   1455
      Left            =   4800
      TabIndex        =   176
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox List8 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   120
         TabIndex        =   177
         Top             =   600
         Width           =   5415
      End
      Begin sevCommand3.Command Command4 
         Height          =   255
         Index           =   19
         Left            =   4440
         TabIndex        =   179
         Top             =   3120
         Width           =   1095
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
         Caption         =   "schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Termine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   178
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Height          =   5895
      Left            =   -1440
      TabIndex        =   149
      Top             =   7680
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   11280
         MaxLength       =   3
         TabIndex        =   172
         Top             =   7440
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "nur 14 tägig"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   171
         Top             =   6960
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   169
         Text            =   "Combo11"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   153
         Text            =   "Combo1"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   152
         Top             =   3120
         Width           =   4695
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   151
         Text            =   "Combo8"
         Top             =   5040
         Width           =   2175
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   150
         Text            =   "Combo7"
         Top             =   6360
         Width           =   4695
      End
      Begin sevCommand3.Command Command1 
         Height          =   400
         Index           =   3
         Left            =   3240
         TabIndex        =   154
         ToolTipText     =   "Kalender"
         Top             =   3680
         Width           =   495
         _ExtentX        =   873
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   14
         Left            =   5520
         TabIndex        =   155
         Top             =   1800
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1085
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
         ToolTip         =   "Runter"
         ToolTipTitle    =   "Runter"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   15
         Left            =   5520
         TabIndex        =   156
         Top             =   1080
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   1085
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
         ToolTip         =   "Rauf"
         ToolTipTitle    =   "Rauf"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   14
         Left            =   9360
         TabIndex        =   167
         Top             =   7920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   400
         Index           =   4
         Left            =   720
         TabIndex        =   168
         ToolTipText     =   "Kalender"
         Top             =   3680
         Width           =   500
         _ExtentX        =   873
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   18
         Left            =   9360
         TabIndex        =   175
         Top             =   6960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
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
         Caption         =   "Abbrechen/Zurück"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Eintrag/Bediener"
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
         Index           =   25
         Left            =   9360
         TabIndex        =   173
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "von (Uhrzeit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   720
         TabIndex        =   170
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Abwesenheit"
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
         Height          =   495
         Index           =   13
         Left            =   120
         TabIndex        =   166
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "ausgewählt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   720
         TabIndex        =   165
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
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
         Index           =   11
         Left            =   720
         TabIndex        =   164
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Grund der Abwesenheit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   720
         TabIndex        =   163
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "bis (Uhrzeit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   3240
         TabIndex        =   162
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Index           =   22
         Left            =   720
         TabIndex        =   161
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Index           =   21
         Left            =   3240
         TabIndex        =   160
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "von (Datum)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   1320
         TabIndex        =   159
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "bis (Datum)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   3840
         TabIndex        =   158
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "an allen Tagen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   720
         TabIndex        =   157
         Top             =   5880
         Width           =   2775
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "verfügbar ist"
      Height          =   7935
      Left            =   1560
      TabIndex        =   100
      Top             =   6720
      Width           =   2175
      Begin sevCommand3.Command Command3 
         Height          =   255
         Index           =   12
         Left            =   1800
         TabIndex        =   103
         Top             =   240
         Width           =   255
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
         Caption         =   "X"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComctlLib.TreeView List3 
         Height          =   6615
         Left            =   120
         TabIndex        =   102
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   11668
         _Version        =   393217
         LabelEdit       =   1
         Scroll          =   0   'False
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin sevCommand3.Command Command3 
         Height          =   360
         Index           =   16
         Left            =   180
         TabIndex        =   174
         Top             =   7240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
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
         Caption         =   "Abwesenheit"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "ausgewählt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   11895
      Begin VB.Frame Frame1 
         Caption         =   "Kunde wählen"
         Height          =   2295
         Left            =   6840
         TabIndex        =   106
         Top             =   0
         Width           =   4815
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   15
            Left            =   3480
            TabIndex        =   134
            Top             =   1650
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Achtung"
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   120
            TabIndex        =   133
            Top             =   960
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.ListBox List6 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   120
            TabIndex        =   111
            Top             =   960
            Width           =   3255
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   9
            Left            =   3480
            TabIndex        =   110
            Top             =   525
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Daten"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   109
            Top             =   1360
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   4
            Left            =   3480
            TabIndex        =   108
            Top             =   800
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   107
            Top             =   240
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "suchen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   21
            Left            =   3480
            TabIndex        =   180
            Top             =   1080
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   20
            Left            =   3480
            TabIndex        =   199
            Top             =   1920
            Width           =   1145
            _ExtentX        =   2011
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Info"
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   22
            Left            =   4240
            TabIndex        =   203
            Top             =   525
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "DS"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label lblUnter 
            BackColor       =   &H00C0C000&
            Caption         =   "Achtung: noch offene Kassiervorgänge"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   120
            TabIndex        =   200
            Top             =   2040
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   114
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "ausgewählt:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "wurde bisher bedient:"
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
            Index           =   11
            Left            =   120
            TabIndex        =   112
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "letzter Kunde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   14
            Left            =   1920
            MouseIcon       =   "frmWKL82.frx":0D64
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   188
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Termin"
         Height          =   2775
         Left            =   9480
         TabIndex        =   96
         Top             =   3600
         Width           =   2175
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   138
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
            Caption         =   "verschieben/ändern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   131
            Top             =   200
            Width           =   855
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   99
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   98
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
            Caption         =   "Druck auf Bon"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   97
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
            Caption         =   "Löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   228
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "gelöschte Termine"
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
            Index           =   19
            Left            =   240
            MouseIcon       =   "frmWKL82.frx":106E
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   145
            ToolTipText     =   "mit Doppelklick zur Auswertung"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
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
            Index           =   10
            Left            =   120
            TabIndex        =   132
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Behandlungsort"
         Height          =   3495
         Left            =   4080
         TabIndex        =   85
         Top             =   0
         Width           =   2655
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "ausgewählt:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mitarbeiter wählen"
         Height          =   3495
         Left            =   120
         TabIndex        =   78
         Top             =   0
         Width           =   3855
         Begin VB.CheckBox chk14t 
            Caption         =   "14tägig"
            Height          =   255
            Left            =   2880
            TabIndex        =   187
            Top             =   2760
            Width           =   855
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   140
            Text            =   "Combo5"
            Top             =   3000
            Width           =   2055
         End
         Begin VB.ComboBox Combo5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   130
            Text            =   "Combo5"
            Top             =   3000
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   82
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   79
            Text            =   "Combo1"
            Top             =   960
            Width           =   2535
         End
         Begin sevCommand3.Command Command1 
            Height          =   360
            Index           =   2
            Left            =   2280
            TabIndex        =   143
            ToolTipText     =   "Kalender"
            Top             =   2280
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
         Begin sevCommand3.Command Command3 
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   147
            Top             =   840
            Width           =   465
            _ExtentX        =   820
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
            ToolTip         =   "Runter"
            ToolTipTitle    =   "Runter"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   148
            Top             =   240
            Width           =   465
            _ExtentX        =   820
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
            ToolTip         =   "Rauf"
            ToolTipTitle    =   "Rauf"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "an allen Tagen"
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
            Left            =   1680
            TabIndex        =   141
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Grund"
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
            Index           =   15
            Left            =   120
            TabIndex        =   139
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "bis"
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
            Left            =   1200
            TabIndex        =   129
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "von"
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
            Left            =   120
            TabIndex        =   128
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label3 
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
            Index           =   12
            Left            =   1200
            TabIndex        =   127
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Index           =   5
            Left            =   120
            TabIndex        =   126
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "bis Uhrzeit:"
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
            Left            =   120
            TabIndex        =   84
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Abwesenheit"
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
            Index           =   9
            Left            =   120
            TabIndex        =   83
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "ausgewählt:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datum/Zeit/Dauer"
         Height          =   1095
         Left            =   6840
         TabIndex        =   71
         Top             =   2400
         Width           =   4815
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   120
            MaxLength       =   10
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   72
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin sevCommand3.Command Command1 
            Height          =   360
            Index           =   1
            Left            =   1080
            TabIndex        =   144
            ToolTipText     =   "Kalender"
            Top             =   240
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
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Datum:"
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
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Uhrzeit:"
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
            Index           =   1
            Left            =   1560
            TabIndex        =   76
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Dauer(in Min.):"
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
            Index           =   2
            Left            =   3000
            TabIndex        =   75
            Top             =   360
            Width           =   1335
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   360
         Index           =   21
         Left            =   11160
         TabIndex        =   66
         Top             =   6480
         Width           =   380
         _ExtentX        =   661
         _ExtentY        =   635
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
         Caption         =   ""
         Image           =   9401
         PictureAlign    =   3
         UseDefaultMaskColor=   -1  'True
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame0 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'Kein
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   6480
         Width           =   9135
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   40
            Left            =   5760
            TabIndex        =   48
            Top             =   1440
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "LÖSCHEN"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   39
            Left            =   5160
            TabIndex        =   47
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   " "
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   41
            Left            =   4560
            TabIndex        =   45
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "."
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   38
            Left            =   3960
            TabIndex        =   44
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "M"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   37
            Left            =   3360
            TabIndex        =   43
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "N"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   36
            Left            =   2760
            TabIndex        =   42
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   35
            Left            =   2160
            TabIndex        =   41
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "V"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   34
            Left            =   1560
            TabIndex        =   40
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   33
            Left            =   960
            TabIndex        =   39
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "X"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   32
            Left            =   360
            TabIndex        =   38
            Top             =   1440
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Y"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   31
            Left            =   6240
            TabIndex        =   37
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Ä"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   30
            Left            =   5640
            TabIndex        =   36
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Ö"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   29
            Left            =   5040
            TabIndex        =   35
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "L"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   28
            Left            =   4440
            TabIndex        =   34
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "K"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   27
            Left            =   3840
            TabIndex        =   33
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "J"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   26
            Left            =   3240
            TabIndex        =   32
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "H"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   25
            Left            =   2640
            TabIndex        =   31
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "G"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   24
            Left            =   2040
            TabIndex        =   30
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "F"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   23
            Left            =   1440
            TabIndex        =   29
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "D"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   22
            Left            =   840
            TabIndex        =   28
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "S"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   21
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "A"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   20
            Left            =   6120
            TabIndex        =   26
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Ü"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   19
            Left            =   5520
            TabIndex        =   25
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   18
            Left            =   4920
            TabIndex        =   24
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "O"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   17
            Left            =   4320
            TabIndex        =   23
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "I"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   16
            Left            =   3720
            TabIndex        =   22
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "U"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   15
            Left            =   3120
            TabIndex        =   21
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Z"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   14
            Left            =   2520
            TabIndex        =   20
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "T"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   13
            Left            =   1920
            TabIndex        =   19
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "R"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   12
            Left            =   1320
            TabIndex        =   18
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   11
            Left            =   720
            TabIndex        =   17
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "W"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
            MenuBackColor   =   16448250
            MenuBackColorChecked=   7323903
            MenuBackColorHover=   10935807
            MenuBorderColor =   8388608
            MenuCheckMarkColorFrom=   16514300
            MenuCheckMarkColorTo=   15462640
            MenuForeColor   =   -2147483640
            MenuForeColorHover=   -2147483640
            ButtonStyle     =   2
            Caption         =   "Q"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   495
            Index           =   9
            Left            =   5400
            TabIndex        =   15
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   8
            Left            =   4800
            TabIndex        =   14
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   7
            Left            =   4200
            TabIndex        =   13
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   6
            Left            =   3600
            TabIndex        =   12
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   5
            Left            =   3000
            TabIndex        =   11
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   4
            Left            =   2400
            TabIndex        =   10
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   3
            Left            =   1800
            TabIndex        =   9
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   2
            Left            =   1200
            TabIndex        =   8
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   7
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   615
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDisabled=   15398133
            BackColorFrom   =   16514300
            BackColorTo     =   15462640
            BackColorCheckedFrom=   15462640
            BackColorCheckedTo=   16514300
            BackColorDownFrom=   12700881
            BackColorDownTo =   15659506
            BackColorHoverFrom=   16514300
            BackColorHoverTo=   15462640
            BorderColor     =   7617536
            BorderColorDisabled=   12240841
            BorderColorFocus=   14986635
            BorderColorHover=   3913721
            ForeColorDisabled=   9609633
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
         Begin VB.Label Label0 
            BackColor       =   &H00808000&
            Caption         =   "Label2"
            Height          =   255
            Left            =   7320
            TabIndex        =   46
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   2
         Left            =   9720
         TabIndex        =   49
         Top             =   7920
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
      Begin VB.Frame Frame7 
         Caption         =   "Behandlung"
         Height          =   2775
         Left            =   120
         TabIndex        =   89
         Top             =   3600
         Width           =   9255
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Index           =   3
            Left            =   5040
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   94
            Top             =   600
            Width           =   4095
         End
         Begin sevCommand3.Command Command4 
            Height          =   255
            Index           =   6
            Left            =   8160
            TabIndex        =   93
            Top             =   240
            Width           =   975
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
            Caption         =   "Leeren"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.ListBox List7 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   120
            TabIndex        =   92
            Top             =   600
            Width           =   4215
         End
         Begin sevCommand3.Command Command6 
            Height          =   495
            Index           =   0
            Left            =   4440
            TabIndex        =   91
            Top             =   1200
            Width           =   495
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   ">"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.ComboBox Combo4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            TabIndex        =   90
            Text            =   "Combo4"
            Top             =   160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "zur Kasse"
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
            Index           =   0
            Left            =   7680
            MouseIcon       =   "frmWKL82.frx":1378
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   146
            ToolTipText     =   "direkt mit Kundendaten und Artikel zur Kasse"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            Caption         =   "Gliederung:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label4 
         Caption         =   "-1"
         Height          =   255
         Left            =   9480
         TabIndex        =   50
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'Kein
      Caption         =   "Termine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   -120
      MouseIcon       =   "frmWKL82.frx":1682
      TabIndex        =   1
      Top             =   360
      Width           =   9855
      Begin sevCommand3.Command Command3 
         Height          =   360
         Index           =   13
         Left            =   120
         TabIndex        =   105
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Caption         =   "Kunde"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   360
         Index           =   10
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   1335
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
         Caption         =   "Heute"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   360
         Index           =   6
         Left            =   7680
         TabIndex        =   65
         Top             =   500
         Width           =   735
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
         Caption         =   "+ Mo"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   360
         Index           =   5
         Left            =   8520
         TabIndex        =   64
         Top             =   500
         Width           =   735
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
         Caption         =   "- Mo"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   360
         Index           =   4
         Left            =   6840
         TabIndex        =   63
         Top             =   500
         Width           =   735
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
         Caption         =   "- Wo"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   360
         Index           =   2
         Left            =   6000
         TabIndex        =   62
         Top             =   500
         Width           =   735
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
         Caption         =   "+ Wo"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   360
         Index           =   9
         Left            =   7800
         TabIndex        =   61
         Top             =   120
         Width           =   380
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
         Caption         =   "<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   360
         Index           =   8
         Left            =   8280
         TabIndex        =   60
         Top             =   120
         Width           =   380
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
         Caption         =   ">"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   120
         Width           =   1695
      End
      Begin sevCommand3.Command Command6 
         Height          =   360
         Index           =   3
         Left            =   9360
         TabIndex        =   59
         Top             =   120
         Width           =   375
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7455
         Left            =   0
         TabIndex        =   2
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   13150
         _Version        =   393216
         Rows            =   57
         Cols            =   6
         FixedCols       =   2
         ForeColor       =   0
         ForeColorFixed  =   0
         ForeColorSel    =   16777215
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   0
         Left            =   8760
         TabIndex        =   142
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "alle anzeigen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   4800
         MouseIcon       =   "frmWKL82.frx":198C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   137
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmWKL82.frx":1C96
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   125
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "verfügbar ist"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   4800
         MouseIcon       =   "frmWKL82.frx":1FA0
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   104
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Index           =   7
         Left            =   3360
         TabIndex        =   70
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Index           =   6
         Left            =   1560
         TabIndex        =   68
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox Text3 
      Height          =   1335
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   55
      Top             =   4320
      Width           =   1815
   End
   Begin sevCommand3.Command Command3 
      Height          =   255
      Index           =   3
      Left            =   9840
      TabIndex        =   51
      Top             =   6480
      Width           =   1815
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
      Caption         =   "neuer Kunde"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   255
      Index           =   6
      Left            =   9840
      TabIndex        =   57
      Top             =   6120
      Width           =   1815
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
      Caption         =   "Info speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   0
      Left            =   9840
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
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
   Begin sevCommand3.Command Command3 
      Height          =   255
      Index           =   11
      Left            =   9840
      TabIndex        =   69
      Top             =   5760
      Width           =   1815
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
      Caption         =   "Info löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ComboBox Combo12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9840
      TabIndex        =   186
      Text            =   "Combo12"
      Top             =   7200
      Width           =   1815
   End
   Begin sevCommand3.Command Command3 
      Height          =   255
      Index           =   5
      Left            =   9840
      TabIndex        =   52
      Top             =   7560
      Width           =   1815
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   201
      Top             =   3240
      Width           =   1815
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Termin kopieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   227
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   9840
      TabIndex        =   202
      Top             =   3555
      Width           =   1815
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   9840
      TabIndex        =   54
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "1"
      Height          =   255
      Left            =   9600
      TabIndex        =   53
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9840
      TabIndex        =   56
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Info"
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
      Left            =   11040
      TabIndex        =   58
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmWKL82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iclick As Integer
Dim dWidth  As Double
Dim gsLastKunde As String

Dim globBuchNr As Long
Dim globRow As Long
Dim globCol As Long

Private Sub AktualisiereTerminTabelleWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim lWoTag As Long
    Dim lMaxIndex As Long
    Dim lMinIndex As Long
    Dim lrow As Long
    Dim lcount As Long
    Dim bgefunden As Boolean
    Dim ctmp As String
    Dim cbednu As String
    Dim cDatum As String
    Dim lDatum As Long
    Dim lDatumVon As Long
    Dim lDatumBis As Long
    Dim lZeitArt As Long
    Dim lZeitGuelt As Long
    Dim cZeitVon As String
    Dim cZeitBis As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cBEGRUENDTX As String
    
    AktualisiereMATerminTabelleWKL82
    
    Label3(7).Caption = ""
    Label3(7).Refresh

    ctmp = Left(Combo2.Text, 2)
    ctmp = UCase$(ctmp)
    
    Select Case ctmp
        Case Is = "MO"
            lWoTag = 1
        Case Is = "DI"
            lWoTag = 2
        Case Is = "MI"
            lWoTag = 3
        Case Is = "DO"
            lWoTag = 4
        Case Is = "FR"
            lWoTag = 5
        Case Is = "SA"
            lWoTag = 6
        Case Is = "SO"
            lWoTag = 7
    End Select
    
    cDatum = Combo2.Text
    cDatum = Right(cDatum, 8)
    
    lDatum = DateValue(cDatum)
    
    'Feiertage 2013
    Dim bFeiertag As Boolean
    bFeiertag = False
    
    If IsThis_EinFeiertag(Format(cDatum, "DD.MM.YYYY")) Then
        bFeiertag = True
    End If
    
    lMinIndex = ((lWoTag - 1) * 3) + 1
    lMaxIndex = ((lWoTag - 1) * 3) + 3
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Col = 0
    
    For lrow = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lrow
        bgefunden = False
        
        For lcount = lMinIndex To lMaxIndex
            ctmp = MSFlexGrid1.Text
            If ctmp >= gZeiten(lcount).Von And ctmp < gZeiten(lcount).Bis Then
                MSFlexGrid1.Col = 1
                MSFlexGrid1.Text = ""
                MSFlexGrid1.Col = 0
                bgefunden = True
                Exit For
            End If
        Next lcount
        
        If Not bgefunden Then
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = "geschl."
            MSFlexGrid1.Col = 0
        End If
        
        If bFeiertag Then
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = "F.Tag"
            MSFlexGrid1.Col = 0
        End If
    Next lrow
    
    cbednu = Combo1.Text
    cbednu = Trim$(Left(cbednu, 3))
    If cbednu <> "" And cbednu <> "Com" Then
       
        cSQL = "Select * from FEHLZEIT where BEDNU = " & cbednu & " "
        cSQL = cSQL & "and DATUM_VON <= " & Trim$(Str$(lDatum)) & " "
        cSQL = cSQL & "and DATUM_BIS >= " & Trim$(Str$(lDatum)) & " "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!DATUM_VON) Then
                    lDatumVon = rsrs!DATUM_VON
                Else
                    lDatumVon = -1
                End If
                If Not IsNull(rsrs!DATUM_BIS) Then
                    lDatumBis = rsrs!DATUM_BIS
                Else
                    lDatumBis = -1
                End If
                If Not IsNull(rsrs!ZEIT_VON) Then
                    cZeitVon = rsrs!ZEIT_VON
                Else
                    cZeitVon = ""
                End If
                If Not IsNull(rsrs!ZEIT_bis) Then
                    cZeitBis = rsrs!ZEIT_bis
                Else
                    cZeitBis = ""
                End If
                If Not IsNull(rsrs!ZEIT_ART) Then
                    lZeitArt = rsrs!ZEIT_ART
                Else
                    lZeitArt = -1
                End If
                If Not IsNull(rsrs!ZEIT_GUELT) Then
                    lZeitGuelt = rsrs!ZEIT_GUELT
                Else
                    lZeitGuelt = -1
                End If
                
                If Not IsNull(rsrs!BEGRUENDTX) Then
                    cBEGRUENDTX = rsrs!BEGRUENDTX
                Else
                    cBEGRUENDTX = ""
                End If
                rsrs.MoveNext
                
                For lrow = 1 To MSFlexGrid1.Rows - 1
                MSFlexGrid1.Row = lrow
                If lZeitArt = 1 Then
                    'Fehlzeit gilt den ganzen Tag über
                    MSFlexGrid1.Col = 1
                    If MSFlexGrid1.Text = "" Then
                        MSFlexGrid1.Text = cBEGRUENDTX
                    End If
                Else
                    'Fehlzeit weist Uhrzeiten auf
                    If lZeitGuelt = 1 Then
                        'Uhrzeiten gelten nur am ersten und letzten Tag
                        If lDatum = lDatumVon Then
                            MSFlexGrid1.Col = 0
                            If MSFlexGrid1.Text >= cZeitVon Then
                            
                                If MSFlexGrid1.Text < cZeitBis Then
                                    MSFlexGrid1.Col = 1
                                    If MSFlexGrid1.Text = "" Then
                                        MSFlexGrid1.Text = cBEGRUENDTX
                                    End If
                                End If
                            End If
                        
                        ElseIf lDatum = lDatumBis Then
                            MSFlexGrid1.Col = 0
                            If MSFlexGrid1.Text < cZeitBis Then
                                MSFlexGrid1.Col = 1
                                If MSFlexGrid1.Text = "" Then
                                    MSFlexGrid1.Text = cBEGRUENDTX
                                End If
                            End If
                        Else
                            MSFlexGrid1.Col = 1
                            If MSFlexGrid1.Text = "" Then
                                MSFlexGrid1.Text = cBEGRUENDTX
                            End If
                        End If
                    Else
                        'Uhrzeiten gelten an allen Fehltagen
                        MSFlexGrid1.Col = 0
                        If MSFlexGrid1.Text >= cZeitVon Then
                            If MSFlexGrid1.Text < cZeitBis Then
                                MSFlexGrid1.Col = 1
                                If MSFlexGrid1.Text = "" Then
                                    MSFlexGrid1.Text = cBEGRUENDTX
                                End If
                            End If
                        End If
                    End If
                End If
            Next lrow
                
            Loop
                            
        Else
            cZeitVon = ""
            cZeitBis = ""
        End If
        
        rsrs.Close: Set rsrs = Nothing
    End If

    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktualisiereTerminTabelleWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Function IsThis_EinFeiertag(sdat As String) As Boolean
On Error GoTo LOKAL_ERROR

Dim sTag As String
Dim sJahr As String
Dim sSQL As String
Dim rsrs As DAO.Recordset

IsThis_EinFeiertag = False

'Menswch wenn die Tabelle nicht da ist
If NewTableSuchenDBKombi("FEIERTAGE", gdBase) = False Then
    CreateTableT2 "FEIERTAGE", gdBase

    'bundesweit fix
    insert_Feiertag "01.01.", "", "Neujahr", 1, 1
    insert_Feiertag "01.05.", "", "Tag der Arbeit", 1, 1
    insert_Feiertag "03.10.", "", "Tag der Deutschen Einheit", 1, 1
    insert_Feiertag "25.12.", "", "1. Weihnachtstag", 1, 1
    insert_Feiertag "26.12.", "", "2. Weihnachtstag", 1, 1
    
    'bundesweit beweglich
    insert_Feiertag "18.04.", "2014", "Karfreitag", 1, 1
    insert_Feiertag "04.04.", "2015", "Karfreitag", 1, 1
    insert_Feiertag "25.03.", "2016", "Karfreitag", 1, 1
    
    
    insert_Feiertag "21.04.", "2014", "Ostermontag", 1, 1
    insert_Feiertag "06.04.", "2015", "Ostermontag", 1, 1
    insert_Feiertag "28.03.", "2016", "Ostermontag", 1, 1
    
    insert_Feiertag "29.05.", "2014", "Christi Himmelfahrt", 1, 1
    insert_Feiertag "14.05.", "2015", "Christi Himmelfahrt", 1, 1
    insert_Feiertag "05.05.", "2016", "Christi Himmelfahrt", 1, 1
    
    insert_Feiertag "09.06.", "2014", "Pfingstmontag", 1, 1
    insert_Feiertag "25.05.", "2015", "Pfingstmontag", 1, 1
    insert_Feiertag "16.05.", "2016", "Pfingstmontag", 1, 1
    
    
    'Nicht bundesweit fix
    insert_Feiertag "06.01.", "", "Heilige Drei Könige", 0, 0
    insert_Feiertag "15.08.", "", "Mariä Himmelfahrt", 0, 0
    insert_Feiertag "31.10.", "", "Reformationstag", 0, 0
    insert_Feiertag "01.11.", "", "Allerheiligen", 0, 0
    
    'Nicht bundesweit beweglich
    insert_Feiertag "19.06.", "2014", "Fronleichnam", 0, 0
    insert_Feiertag "04.06.", "2015", "Fronleichnam", 0, 0
    insert_Feiertag "26.05.", "2016", "Fronleichnam", 0, 0
    
    insert_Feiertag "20.11.", "2013", "Buß- und Bettag", 0, 0
    insert_Feiertag "19.11.", "2014", "Buß- und Bettag", 0, 0
    insert_Feiertag "18.11.", "2015", "Buß- und Bettag", 0, 0
    insert_Feiertag "16.11.", "2016", "Buß- und Bettag", 0, 0

End If


























sTag = Left(sdat, 6)
sJahr = Right(sdat, 4)





' fix
sSQL = "Select * from Feiertage where "
sSQL = sSQL & " FDAT = '" & sTag & "'"
sSQL = sSQL & " and FDATJAHR = ''"
sSQL = sSQL & " and ANWENDEN = -1 "
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    IsThis_EinFeiertag = True
End If
rsrs.Close: Set rsrs = Nothing

If IsThis_EinFeiertag = False Then

    ' beweglich
    sSQL = "Select * from Feiertage where "
    sSQL = sSQL & " FDAT = '" & sTag & "'"
    sSQL = sSQL & " and FDATJAHR = '" & sJahr & "'"
    sSQL = sSQL & " and ANWENDEN = -1 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        IsThis_EinFeiertag = True
    End If
    rsrs.Close: Set rsrs = Nothing

End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "IsThis_EinFeiertag"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub AktualisiereMATerminTabelleWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cSQL3       As String
    Dim rsRs2       As Recordset
    Dim rsRs3       As Recordset
    Dim lDatum      As Long
    Dim cDatum      As String
    Dim dUhrzeit    As Double
    Dim lcount      As Long
    Dim dSprung     As Double
    Dim cbednu      As String
    Dim czeit       As String
    Dim cKabine     As String
    Dim cKuerzel    As String
    Dim cBehandlung As String
    Dim lrow        As Long
    Dim lcol        As Long
    Dim lBuchnr     As Long
    Dim lFarbcode   As Long
    Dim ctmpbeh     As String
    Dim lPos        As Long
    Dim bSpezi      As Boolean
    Dim glBuchnr    As Long
    
    dSprung = TimeValue(gcZeitBlock)
        
    cDatum = Right(Combo2.Text, 8)
    lDatum = DateValue(cDatum)
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Visible = False
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Rows = (TimeValue(gcEndeZeit) - TimeValue(gcStartZeit)) / TimeValue(gcZeitBlock)
    
    cSQL = "Select * from BEDTERM order by bednu desc "
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    If Not rsRs2.EOF Then
        rsRs2.MoveFirst
        Do While Not rsRs2.EOF
        
            If Not IsNull(rsRs2!FARBCODE) Then
                lFarbcode = rsRs2!FARBCODE
            Else
                lFarbcode = 0
            End If
            
            If Not IsNull(rsRs2!BEDNU) Then
                cbednu = rsRs2!BEDNU
            Else
                cbednu = ""
            End If
            
            ZeigeMitarbeiterInFarbeWKL82 cbednu, lFarbcode
            rsRs2.MoveNext
        Loop
    End If
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
    
    dUhrzeit = TimeValue(gcStartZeit)
    
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = 0
        MSFlexGrid1.RowHeight(lcount) = 250
        MSFlexGrid1.Text = Format$(dUhrzeit + (TimeValue(gcZeitBlock) * lcount), "HH:MM")
    Next lcount
    
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktualisiereMATerminTabelleWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub ZeigeMitarbeiterInFarbeWKL82(cbednu As String, lFarbe As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cSQL3       As String
    Dim rsRs3       As Recordset
    Dim lDatum      As Long
    Dim cDatum      As String
    Dim dUhrzeit    As Double
    Dim lcount      As Long
    Dim dSprung     As Double
    
    Dim czeit       As String
    Dim cKabine     As String
    Dim cKuerzel    As String
    Dim cBehandlung As String
    Dim lrow        As Long
    Dim lcol        As Long
    Dim lBuchnr     As Long
    Dim ctmpbeh     As String
    Dim lPos        As Long
    Dim bSpezi      As Boolean
    Dim bOhneVK     As Boolean
    Dim lKUNDNR     As Long
    Dim glBuchnr    As Long
    Dim cbedname    As String
    Dim sKUNDNR     As String
    Dim sKundennameText As String
    
    dSprung = TimeValue(gcZeitBlock)
       
    cDatum = Right(Combo2.Text, 8)
    lDatum = DateValue(cDatum)
    
    cSQL = "Select  "
    cSQL = cSQL & " t.BUCHUNGSNR "
    cSQL = cSQL & " ,t.Uhrzeit "
    cSQL = cSQL & " ,t.KABINE "
    cSQL = cSQL & " ,t.Kuerzel "
    cSQL = cSQL & " ,t.BEDNU "
    cSQL = cSQL & " ,t.bedeintrag "
    cSQL = cSQL & " ,t.bedname "
    cSQL = cSQL & " ,t.Behandlung "
    cSQL = cSQL & " ,t.Kundnr "
    
    cSQL = cSQL & " from TERMINE t "
    cSQL = cSQL & " where t.BEDNU = " & cbednu & " and t.DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & " order by t.UHRZEIT "
    
    Dim lMerkKundnr As Long
    lMerkKundnr = 0
    lKUNDNR = 0
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
hier:
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!BUCHUNGSNR) Then
                lBuchnr = rsrs!BUCHUNGSNR
            End If
            
            bSpezi = False
            cSQL3 = "Select * from SPEZIINFO where BUCHUNGSNR = " & lBuchnr
            Set rsRs3 = gdBase.OpenRecordset(cSQL3)
            If Not rsRs3.EOF Then
                rsRs3.MoveFirst
                If Not IsNull(rsRs3!gesehen) Then
                    If rsRs3!gesehen = False Then
                        bSpezi = True
                    End If
                End If
            End If
            rsRs3.Close: Set rsRs3 = Nothing
            
            
            
            If Not IsNull(rsrs!Kundnr) Then
                lKUNDNR = rsrs!Kundnr
            End If
            
            If lMerkKundnr <> lKUNDNR Then
                lMerkKundnr = lKUNDNR
            
            
            
            
                bOhneVK = False 'UMSKDJ
                cSQL3 = "Select count(*) as max from Kassjour where KUNDNR = " & lMerkKundnr
                Set rsRs3 = gdBase.OpenRecordset(cSQL3)
                If Not rsRs3.EOF Then
                    If Not IsNull(rsRs3!Max) Then
                        If rsRs3!Max = 0 Then
                            bOhneVK = True
                        End If
                    End If
                End If
                rsRs3.Close: Set rsRs3 = Nothing
            
            End If
                
            If Not IsNull(rsrs!Uhrzeit) Then
                czeit = rsrs!Uhrzeit
            Else
                czeit = ""
            End If
            If Not IsNull(rsrs!Kabine) Then
                cKabine = rsrs!Kabine
            Else
                cKabine = ""
            End If
            
            If Label2(10).Caption = "alle anzeigen" Then
                If WirdDieKabineAngezeigt(cKabine) = False Then
                    rsrs.MoveNext
                    GoTo hier
                End If
            End If
                
            
            If Not IsNull(rsrs!Kuerzel) Then
                cKuerzel = rsrs!Kuerzel
            Else
                cKuerzel = ""
            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                sKUNDNR = rsrs!Kundnr
            Else
                sKUNDNR = ""
            End If
            
            cbedname = ""
            Dim cbednua As String
            cbednua = ""
            If Not IsNull(rsrs!BEDNU) Then
                cbednua = Space$(3 - Len(rsrs!BEDNU)) & rsrs!BEDNU
            End If
            
            Dim cbedEintrag As String
            cbedEintrag = ""
            If Not IsNull(rsrs!bedeintrag) Then
                cbedEintrag = rsrs!bedeintrag
            End If
            
            If Not IsNull(rsrs!bedname) Then
                cbedname = cbednua & " " & Trim(rsrs!bedname)
            End If
            
            If glBuchnr <> lBuchnr Then
                glBuchnr = lBuchnr
                If Not IsNull(rsrs!Behandlung) Then
                    cBehandlung = SwapStr(rsrs!Behandlung, Chr(13), ".")
                    cBehandlung = SwapStr(cBehandlung, Chr(10), ".")
                    cBehandlung = SwapStr(cBehandlung, "..", ".")
                    cBehandlung = cBehandlung & "."
                Else
                    cBehandlung = ""
                End If
            End If
            
            dUhrzeit = TimeValue(czeit)
            lrow = (dUhrzeit - TimeValue(gcStartZeit)) / dSprung
            
            MSFlexGrid1.Row = 0
            For lcol = 2 To MSFlexGrid1.Cols - 1
                MSFlexGrid1.Col = lcol
                If Trim$(UCase$(MSFlexGrid1.Text)) = Trim$(UCase$(cKabine)) Then
                    Exit For
                End If
            Next lcol
            
            If lcol = MSFlexGrid1.Cols Then
                lcol = lcol - 1
            End If
            
            If lrow < 0 Then
                Exit Sub
            ElseIf lrow < MSFlexGrid1.Rows Then
                MSFlexGrid1.Row = lrow
            Else
                Exit Sub
            End If
            MSFlexGrid1.Col = lcol
             
            MSFlexGrid1.CellBackColor = FarbeBackColor(lFarbe)
            MSFlexGrid1.CellForeColor = FarbeForeColor(lFarbe)
            
            
            lPos = InStr(1, cBehandlung, ".")
            If lPos <> 0 Then
                ctmpbeh = Left(cBehandlung, lPos - 1)
                cBehandlung = Right(cBehandlung, Len(cBehandlung) - lPos)
            Else
            
                ctmpbeh = ""
            End If
            
            MSFlexGrid1.Text = ""
            
            If bOhneVK = True Then
                MSFlexGrid1.CellFontItalic = True
'                MSFlexGrid1.CellFontBold = True
            Else
                MSFlexGrid1.CellFontItalic = False
            End If
            
            If bSpezi = True Then
                MSFlexGrid1.Text = "! "
            End If
            
            If gbTerm_Name = True Then
                sKundennameText = lookingForKundendaten(sKUNDNR).nachname
                MSFlexGrid1.Text = MSFlexGrid1.Text & sKundennameText & " / " & ctmpbeh
            Else
                MSFlexGrid1.Text = MSFlexGrid1.Text & cKuerzel & " / " & ctmpbeh
            End If
            
            If Len(MSFlexGrid1.Text) < 50 Then
                MSFlexGrid1.Text = MSFlexGrid1.Text & Space(50 - Len(MSFlexGrid1.Text)) & cbedname
            Else
                MSFlexGrid1.Text = Left(MSFlexGrid1.Text, 49) & Space(1) & cbedname
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
    Fehler.gsFunktion = "ZeigeMitarbeiterInFarbeWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub AusrichtenTabelleWKL82(sWelche As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim dUhrzeit As Double
    
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iSichtbar As Integer
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 700
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 800
    MSFlexGrid1.Text = "Bemerk"
    
    For lcount = 2 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Col = lcount
        MSFlexGrid1.ColWidth(lcount) = 1550
'        MSFlexGrid1.ColWidth(lcount) = 1900
        
        'neu
        If sWelche = "alle" Then
            cSQL = "Select anzeigeN from PFLEGORT "
            cSQL = cSQL & " where Ucase(BEZEICH) =  '" & UCase(Trim(MSFlexGrid1.TextMatrix(0, lcount))) & "'"
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
            
                iSichtbar = -1
                If Not IsNull(rsrs!anzeigeN) Then
                    iSichtbar = rsrs!anzeigeN
                    If iSichtbar = 0 Then
                        MSFlexGrid1.ColWidth(lcount) = 400
                    End If
                End If
            End If
            rsrs.Close
        End If
        'neu Ende
        
        
        
    Next lcount
    
    MSFlexGrid1.Rows = (TimeValue(gcEndeZeit) - TimeValue(gcStartZeit)) / TimeValue(gcZeitBlock)
        
    dUhrzeit = TimeValue(gcStartZeit)
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = 0
        MSFlexGrid1.RowHeight(lcount) = 250
        MSFlexGrid1.Text = Format$(dUhrzeit + (TimeValue(gcZeitBlock) * lcount), "HH:MM")
    Next lcount

    MSFlexGrid1.FixedCols = 2
    
    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AusrichtenTabelleWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub DruckeEinsatzPlanWKL82(cBediener As String, cDatum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim rsRs3 As Recordset
    
    Dim czeit As String
    Dim dZeit As Double
    Dim dViertelStunde As Double
    
    Dim cKdnr As String
    Dim cKundenName As String
    Dim cbednu As String
    
    dViertelStunde = TimeValue(gcZeitBlock)
    
    

    loeschNEW "Dru_term", gdBase
    
    cSQL = "Create Table DRU_TERM "
    cSQL = cSQL & "( BEDNAME Text(32)"
    cSQL = cSQL & ", BUCHUNGSNR Long"
    cSQL = cSQL & ", DATUM Text(10)"
    cSQL = cSQL & ", TERMIN Text(5)"
    cSQL = cSQL & ", ENDE Text(5)"
    cSQL = cSQL & ", KABINE Text(35)"
    cSQL = cSQL & ", KUNDNR LONG"
    cSQL = cSQL & ", KUERZEL Text(5)"
    cSQL = cSQL & ", KDNAME Text(250)"
    cSQL = cSQL & ", BEHANDLUNG Text(250)"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError
    
    cDatum = Right(cDatum, 8)
    
    lDatum = DateValue(cDatum)
    cbednu = Left(cBediener, 3)
    cbednu = Trim(cbednu)
    cBediener = Trim(cBediener)
    
    cSQL = "Select BUCHUNGSNR, DATUM, MIN(UHRZEIT) as TERMIN, KABINE, "
    cSQL = cSQL & "KUNDNR, KUERZEL, BEHANDLUNG "
    cSQL = cSQL & "from TERMINE "
    cSQL = cSQL & "where bednu = " & cbednu
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "group by BUCHUNGSNR, DATUM, KABINE, KUNDNR, KUERZEL, BEHANDLUNG"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    
    cSQL = "Select BUCHUNGSNR, MAX(UHRZEIT) as ENDE "
    cSQL = cSQL & "from TERMINE "
    cSQL = cSQL & "where BEDNAME like '*" & cBediener & "*' "
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "group by BUCHUNGSNR"
    Set rsRs3 = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from DRU_TERM "
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsRs2.AddNew
            rsRs2!bedname = cBediener
            rsRs2!Datum = cDatum
            rsRs2!BUCHUNGSNR = rsrs!BUCHUNGSNR
            rsRs2!TERMIN = rsrs!TERMIN
            If Not rsRs3.EOF Then
                rsRs3.MoveFirst
                Do While Not rsRs3.EOF
                    If rsRs3!BUCHUNGSNR = rsrs!BUCHUNGSNR Then
                        czeit = rsRs3!ENDE
                        dZeit = TimeValue(czeit)
                        dZeit = dZeit + dViertelStunde
                        czeit = Format$(dZeit, "HH:MM")
                        rsRs2!ENDE = czeit
                    End If
                    rsRs3.MoveNext
                Loop
                rsRs3.MoveFirst

            End If
            rsRs2!Kabine = rsrs!Kabine
            rsRs2!Kundnr = rsrs!Kundnr
            If Not IsNull(rsrs!Kundnr) Then
                cKdnr = rsrs!Kundnr
                cKundenName = fnHoleKundenNameVollWKL82(cKdnr)
            Else
                cKundenName = ""
            End If
            rsRs2!Kuerzel = rsrs!Kuerzel
            rsRs2!KdName = cKundenName
            rsRs2!Behandlung = rsrs!Behandlung
            rsRs2.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
    rsrs.Close: Set rsrs = Nothing
    rsRs3.Close: Set rsRs3 = Nothing
    reportbildschirm "", "aWKL008"

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeEinsatzPlanWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub DruckeTagesPlanWKL82(cDatum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim rsRs3 As Recordset
    
    Dim czeit As String
    Dim dZeit As Double
    Dim dViertelStunde As Double
    
    Dim cKdnr As String
    Dim cKundenName As String
    
    dViertelStunde = TimeValue(gcZeitBlock)
    
    loeschNEW "DRU_TERM", gdBase
    cSQL = "Create Table DRU_TERM "
    cSQL = cSQL & "( BEDNAME Text(32)"
    cSQL = cSQL & ", BUCHUNGSNR Long"
    
    cSQL = cSQL & ", DATUM Text(10)"
    cSQL = cSQL & ", TERMIN Text(5)"
    cSQL = cSQL & ", ENDE Text(5)"
    cSQL = cSQL & ", KABINE Text(35)"
    cSQL = cSQL & ", KUNDNR LONG"
    cSQL = cSQL & ", KUERZEL Text(5)"
    cSQL = cSQL & ", KDNAME Text(250)"
    cSQL = cSQL & ", BEHANDLUNG Text(250)"
    cSQL = cSQL & ")"

    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cDatum = Right(cDatum, 8)
    
    lDatum = DateValue(cDatum)
    
    cSQL = "Select BUCHUNGSNR, BEDNAME, DATUM, MIN(UHRZEIT) as TERMIN, KABINE, "
    cSQL = cSQL & "KUNDNR, KUERZEL, BEHANDLUNG "
    cSQL = cSQL & "from TERMINE "
    cSQL = cSQL & "where DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "group by BEDNAME, BUCHUNGSNR, DATUM, KABINE, KUNDNR, KUERZEL, BEHANDLUNG"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    
    cSQL = "Select BUCHUNGSNR, MAX(UHRZEIT) as ENDE "
    cSQL = cSQL & "from TERMINE "
    cSQL = cSQL & "where DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "group by BUCHUNGSNR"
    
    Set rsRs3 = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from DRU_TERM "
    
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsRs2.AddNew
            rsRs2!bedname = rsrs!bedname
            rsRs2!Datum = cDatum
            rsRs2!BUCHUNGSNR = rsrs!BUCHUNGSNR
            rsRs2!TERMIN = rsrs!TERMIN
            If Not rsRs3.EOF Then
                rsRs3.MoveFirst
                Do While Not rsRs3.EOF
                    If rsRs3!BUCHUNGSNR = rsrs!BUCHUNGSNR Then
                        czeit = rsRs3!ENDE
                        dZeit = TimeValue(czeit)
                        dZeit = dZeit + dViertelStunde
                        czeit = Format$(dZeit, "HH:MM")
                        rsRs2!ENDE = czeit
                    End If
                    rsRs3.MoveNext
                Loop
                rsRs3.MoveFirst

            End If
            rsRs2!Kabine = rsrs!Kabine
            rsRs2!Kundnr = Val(rsrs!Kundnr)
            If Not IsNull(rsrs!Kundnr) Then
                cKdnr = Val(rsrs!Kundnr)
                cKundenName = fnHoleKundenNameVollWKL82(cKdnr)
            Else
                cKundenName = ""
            End If
            rsRs2!Kuerzel = rsrs!Kuerzel
            rsRs2!KdName = cKundenName
            rsRs2!Behandlung = rsrs!Behandlung
            rsRs2.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
    rsrs.Close: Set rsrs = Nothing
    rsRs3.Close: Set rsRs3 = Nothing
    
    reportbildschirm "", "aWKL019"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTagesPlanWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
   
End Sub
Private Sub DruckeTagesPlanNeuFarbe(cDatum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim cKabineText As String
    
    loeschNEW "TPLAN", gdBase
    CreateTableT2 "TPLAN", gdBase
    
    MSFlexGrid1.Redraw = False

    MSFlexGrid1.Row = 0
    For i = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = i
        
        sSQL = "Insert into TPLAN ( Zeit "
        k = 0
        For j = 2 To MSFlexGrid1.Cols - 1
            k = k + 1
            sSQL = sSQL & ", ORT" & k
            sSQL = sSQL & ", FarbeORT" & k
            sSQL = sSQL & ", FarbeSORT" & k
        Next j
        sSQL = sSQL & " ) values ( "
        
        MSFlexGrid1.Col = 0
        sSQL = sSQL & "  '" & MSFlexGrid1.Text & "'  "
        
        For j = 2 To MSFlexGrid1.Cols - 1
        
            MSFlexGrid1.Col = j
            cKabineText = ""
            cKabineText = MSFlexGrid1.Text
            For l = 0 To 10
                cKabineText = SwapStr(cKabineText, "  ", " ")
            Next l
            
            sSQL = sSQL & ",  '" & cKabineText & "'  "
            sSQL = sSQL & ",  " & MSFlexGrid1.CellBackColor & "  "
            
            If i = 0 Then
                sSQL = sSQL & ",  " & vbWhite & "  "
            Else
                sSQL = sSQL & ",  " & MSFlexGrid1.CellForeColor & "  "
            End If
        Next j
        sSQL = sSQL & " )"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Redraw = True
    
    sSQL = " Update TPLAN set Datum = '" & cDatum & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    If k > 7 Then
        reportbildschirm "", "aWKL082b"
    ElseIf k > 3 Then
        reportbildschirm "", "aWKL082"
    Else
        reportbildschirm "", "aWKL082a"
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTagesPlanNeuFarbe"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeWochenPlanWKL82(cKW As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cVon As Date
    Dim cBis As Date
    Dim cFeld As String
    Dim cChKw As String
    Dim DateHeut As Date
            
    DateHeut = DateValue(Right(Combo2.Text, 8))
    
    Do
        DateHeut = DateHeut - 1
        cChKw = DatePart("ww", DateHeut, vbMonday)
    Loop While cKW < cChKw
    
    
    cVon = DateHeut + 1
    cBis = cVon + 6
    
    loeschNEW "TERMPRINT", gdBase
    CreateTable "TERMPRINT", gdBase
    
    cSQL = "Insert into TERMPRINT select "
    cSQL = cSQL & " BEDNAME "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & ", BEHANDLUNG "
    cSQL = cSQL & ", BUCHUNGSNR "
    cSQL = cSQL & ", DATUM "
    cSQL = cSQL & ", KABINE "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", UHRZEIT "
    cSQL = cSQL & ", BEDEINTRAG "
    cSQL = cSQL & " from termine where datum between " & CLng(cVon) & " "
    cSQL = cSQL & "  and " & CLng(cBis) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT set von = " & CLng(cVon)
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT set bis = " & CLng(cBis)
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT set adate = datum "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT inner join kunden on TERMPRINT.KUNDNR = Kunden.Kundnr "
    cSQL = cSQL & " set TERMPRINT.Name = Kunden.Name "
    cSQL = cSQL & " , TERMPRINT.TEL = Kunden.TEL "
    cSQL = cSQL & " , TERMPRINT.MOBILTEL = Kunden.MOBILTEL "
    cSQL = cSQL & " , TERMPRINT.VORNAME = Kunden.VORNAME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TERMPRINT "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Behandlung) Then
                cFeld = rsrs!Behandlung
            Else
                cFeld = ""
            End If
            
            cFeld = SwapStr(cFeld, Chr(13), " ")
            cFeld = SwapStr(cFeld, Chr(10), " ")
            
            rsrs.Edit
            rsrs!Behandlung = cFeld
            rsrs.Update
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    reportbildschirm "", "aWKL019c"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeWochenPlanWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeMitarbeiterEinsatzPlanWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cVon As Date
    Dim cFeld As String
    Dim DateHeut As Date
    Dim cDauer As String
            
    DateHeut = DateValue(Right(Combo2.Text, 8))
    cVon = DateHeut
    
    loeschNEW "TERMPRINT_EP", gdBase
    CreateTableT2 "TERMPRINT_EP", gdBase
    
    cSQL = "Insert into TERMPRINT_EP select "
    cSQL = cSQL & " BEDNAME "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & ", BEHANDLUNG "
    cSQL = cSQL & ", BUCHUNGSNR "
    cSQL = cSQL & ", DATUM "
    cSQL = cSQL & ", KABINE "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", UHRZEIT "
    cSQL = cSQL & ", BEDEINTRAG "
    cSQL = cSQL & " from termine where datum = " & CLng(cVon) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP set von = " & CLng(cVon)
    gdBase.Execute cSQL, dbFailOnError
    
  
    cSQL = "Update TERMPRINT_EP set adate = datum "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP inner join kunden on TERMPRINT_EP.KUNDNR = Kunden.Kundnr "
    cSQL = cSQL & " set TERMPRINT_EP.Name = Kunden.Name "
    cSQL = cSQL & " , TERMPRINT_EP.TEL = Kunden.TEL "
    cSQL = cSQL & " , TERMPRINT_EP.MOBILTEL = Kunden.MOBILTEL "
    cSQL = cSQL & " , TERMPRINT_EP.VORNAME = Kunden.VORNAME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TERMPRINT_EP "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Behandlung) Then
                cFeld = rsrs!Behandlung
            Else
                cFeld = ""
            End If
            
            cFeld = SwapStr(cFeld, Chr(13), " ")
            cFeld = SwapStr(cFeld, Chr(10), " ")
            
            rsrs.Edit
            rsrs!Behandlung = Trim(cFeld)

            rsrs.Update
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Dim cStartzeit As String
    Dim cEndZeit As String
    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim lBuchnr As Long
    
    loeschNEW "TERMPRINT_MEP", gdBase
    CreateTableT2 "TERMPRINT_MEP", gdBase
    
    cSQL = "Select BUCHUNGSNR, max(Uhrzeit) as maxizeit, min(Uhrzeit) as minizeit from TERMPRINT_EP group by Buchungsnr "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!BUCHUNGSNR) Then
                lBuchnr = rsrs!BUCHUNGSNR
            End If
            
            If Not IsNull(rsrs!maxizeit) Then
                cEndZeit = rsrs!maxizeit
            Else
                cEndZeit = ""
            End If
            
            If Not IsNull(rsrs!minizeit) Then
                cStartzeit = rsrs!minizeit
            Else
                cStartzeit = ""
            End If
            
            dStart = TimeValue(cStartzeit)
            dEnde = TimeValue(cEndZeit)
            dEnde = dEnde + TimeValue(gcZeitBlock)
    
            dDauer = dEnde - dStart
            cDauer = Format$(dDauer, "HH:MM")
            
            
            cSQL = "Insert into TERMPRINT_MEP (buchnr,Dauer,Uhrzeit_ende,UHRZEIT) values ("
            cSQL = cSQL & " " & lBuchnr & ",'" & cDauer & "',  '" & Format$(dEnde, "HH:MM") & "',  '" & Format$(dStart, "HH:MM") & "')"
            gdBase.Execute cSQL, dbFailOnError
        
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Update TERMPRINT_MEP M inner join TERMPRINT_EP E on M.buchnr = E.BUCHUNGSNR "
    cSQL = cSQL & " set M.Name = E.Name "
    cSQL = cSQL & ",M.BEDNAME = E.BEDNAME "
    cSQL = cSQL & ",M.BEDNU = E.BEDNU  "
    cSQL = cSQL & ",M.BEHANDLUNG = E.BEHANDLUNG "
    cSQL = cSQL & ",M.DATUM = E.DATUM  "
    cSQL = cSQL & ",M.KABINE = E.KABINE  "
    cSQL = cSQL & ",M.KUERZEL = E.KUERZEL  "
    cSQL = cSQL & ",M.KUNDNR = E.KUNDNR  "
    
    cSQL = cSQL & ",M.TEL = E.TEL  "
    cSQL = cSQL & ",M.MOBILTEL = E.MOBILTEL  "
    cSQL = cSQL & ",M.VORNAME = E.VORNAME  "

    
'    cSQL = cSQL & ",M.UHRZEIT = E.UHRZEIT  "
    cSQL = cSQL & ",M.adate = E.adate  "
    cSQL = cSQL & ",M.von = E.von  "
    cSQL = cSQL & ",M.bis = E.bis  "
    cSQL = cSQL & ",M.BEDEINTRAG = E.BEDEINTRAG   "
    
    gdBase.Execute cSQL, dbFailOnError
    
        
    reportbildschirm "", "aWKL019d"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeMitarbeiterEinsatzPlanWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeMitarbeiterVerfügbarkeit()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsBed As DAO.Recordset
    Dim rsBuch As DAO.Recordset
    
    Dim sZeitblock As String
    Dim dateDat As Date
    
    Dim cDauer As String
    Dim lZeit1 As Long
    Dim ldynStartzeit As Long

    Dim ibednu As Integer
    Dim lFarbe As Long
    Dim lBuchungsnr As Long
    
    
    
    
    
    Dim lStartzeit As Long
    Dim lEndzeit As Long
    Dim lzeitblock As Long
    Dim cFeld As String
    
    cFeld = gcZeitBlock
    cFeld = SwapStr(cFeld, ":", "")
    lzeitblock = CLng(cFeld)
    
    cFeld = gcStartZeit
    cFeld = SwapStr(cFeld, ":", "")
    lStartzeit = CLng(cFeld)
    
    lStartzeit = lStartzeit + lzeitblock
                    
    If Right(CStr(lStartzeit), 2) = "60" Then
        lStartzeit = lStartzeit + 40
    End If
    
    cFeld = gcEndeZeit
    cFeld = SwapStr(cFeld, ":", "")
    lEndzeit = CLng(cFeld)
    
    
    
    dateDat = DateValue(Right(Combo2.Text, 8))
    
    
    loeschNEW "TERM_VERFUEGBAR", gdBase
    CreateTableT2 "TERM_VERFUEGBAR", gdBase
    
    Screen.MousePointer = 11

    cSQL = "Select * from BEDTERM order by bednu asc "
    Set rsBed = gdBase.OpenRecordset(cSQL)
    If Not rsBed.EOF Then
        rsBed.MoveFirst
        Do While Not rsBed.EOF
            If Not IsNull(rsBed!BEDNU) Then
                ibednu = rsBed!BEDNU
            Else
                ibednu = 0
            End If
            
            If Not IsNull(rsBed!FARBCODE) Then
                lFarbe = rsBed!FARBCODE
            Else
                lFarbe = 0
            End If
            
            ldynStartzeit = lStartzeit
            
            Dim lDatumBeginn As Long
            Dim lDatumEnde As Long
            Dim l As Long
            lDatumBeginn = CLng(dateDat)
            lDatumEnde = lDatumBeginn + 30
            
            For l = lDatumBeginn To lDatumEnde
            
                cSQL = "Select max(Buchungsnr) as maxBuch,min(uhrzeit) as mini from Termine where datum = " & l
                cSQL = cSQL & " and bednu = " & ibednu
                cSQL = cSQL & " group by Buchungsnr order by min(uhrzeit) asc"
                Set rsBuch = gdBase.OpenRecordset(cSQL)
                If Not rsBuch.EOF Then
                    rsBuch.MoveFirst
                    Do While Not rsBuch.EOF
                        If Not IsNull(rsBuch!maxBuch) Then
                            lBuchungsnr = rsBuch!maxBuch
                        Else
                            lBuchungsnr = 0
                        End If
    
                        lZeit1 = ermMinZeitperBuchung(lBuchungsnr)
                        
                        If lZeit1 > ldynStartzeit Then
                                            
                            sZeitblock = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lZeit1)
                            
                            cSQL = "Insert into TERM_VERFUEGBAR (Datum,Bediener,Zeitblock) values ("
                            cSQL = cSQL & " " & l & ",'" & ermBEDbez(CLng(ibednu)) & "',  '" & sZeitblock & "')"
                            gdBase.Execute cSQL, dbFailOnError
                        End If
                        
                        ldynStartzeit = ermMaxZeitperBuchung(lBuchungsnr)
                        ldynStartzeit = ldynStartzeit + lzeitblock
                        
                        If Right(CStr(ldynStartzeit), 2) = "60" Then
                            ldynStartzeit = ldynStartzeit + 40
                        End If
                        
                    rsBuch.MoveNext
                    Loop
                Else
                    sZeitblock = "von " & zeitanz(lStartzeit) & " bis " & gcEndeZeit
                    
                    cSQL = "Insert into TERM_VERFUEGBAR (Datum,Bediener,Zeitblock) values ("
                    cSQL = cSQL & " " & l & ",'" & ermBEDbez(CLng(ibednu)) & "',  '" & sZeitblock & "')"
                    gdBase.Execute cSQL, dbFailOnError
                End If
                
                If ldynStartzeit <> lStartzeit Then
                    If ldynStartzeit < lEndzeit Then
            
                        sZeitblock = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lEndzeit)
                        
                        cSQL = "Insert into TERM_VERFUEGBAR (Datum,Bediener,Zeitblock) values ("
                        cSQL = cSQL & " " & l & ",'" & ermBEDbez(CLng(ibednu)) & "',  '" & sZeitblock & "')"
                        gdBase.Execute cSQL, dbFailOnError
                    End If
                End If
            
                rsBuch.Close: Set rsBuch = Nothing
            
            
            Next l
            
            
            rsBed.MoveNext
        Loop
    End If
    rsBed.Close: Set rsBed = Nothing
    
    
    Screen.MousePointer = 0
        
    
            
    
        
    reportbildschirm "", "aWKL019h"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeMitarbeiterVerfügbarkeit"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeTermineSMS()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cVon As Date
    Dim cFeld As String
    Dim DateHeut As Date
    Dim cDauer As String
            
    DateHeut = DateValue(Right(Combo2.Text, 8))
    cVon = DateHeut
    
    loeschNEW "TERMPRINT_EP", gdBase
    CreateTableT2 "TERMPRINT_EP", gdBase
    
    cSQL = "Insert into TERMPRINT_EP select "
    cSQL = cSQL & " BEDNAME "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & ", BEHANDLUNG "
    cSQL = cSQL & ", BUCHUNGSNR "
    cSQL = cSQL & ", DATUM "
    cSQL = cSQL & ", KABINE "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", UHRZEIT "
    cSQL = cSQL & ", BEDEINTRAG "
    cSQL = cSQL & " from termine where datum = " & CLng(cVon) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP set von = " & CLng(cVon)
    gdBase.Execute cSQL, dbFailOnError
    
  
    cSQL = "Update TERMPRINT_EP set adate = datum "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP inner join kunden on TERMPRINT_EP.KUNDNR = Kunden.Kundnr "
    cSQL = cSQL & " set TERMPRINT_EP.Name = Kunden.Name "
    cSQL = cSQL & " , TERMPRINT_EP.TEL = Kunden.TEL "
    cSQL = cSQL & " , TERMPRINT_EP.MOBILTEL = Kunden.MOBILTEL "
    cSQL = cSQL & " , TERMPRINT_EP.VORNAME = Kunden.VORNAME "
    cSQL = cSQL & " , TERMPRINT_EP.EMAIL = Kunden.EMAIL "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TERMPRINT_EP "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Behandlung) Then
                cFeld = rsrs!Behandlung
            Else
                cFeld = ""
            End If
            
            cFeld = SwapStr(cFeld, Chr(13), " ")
            cFeld = SwapStr(cFeld, Chr(10), " ")
            
            rsrs.Edit
            rsrs!Behandlung = Trim(cFeld)

            rsrs.Update
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Dim cStartzeit As String
    Dim cEndZeit As String
    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim lBuchnr As Long
    
    loeschNEW "TERMPRINT_MEP", gdBase
    CreateTableT2 "TERMPRINT_MEP", gdBase
    
    cSQL = "Select BUCHUNGSNR, max(Uhrzeit) as maxizeit, min(Uhrzeit) as minizeit from TERMPRINT_EP group by Buchungsnr "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!BUCHUNGSNR) Then
                lBuchnr = rsrs!BUCHUNGSNR
            End If
            
            If Not IsNull(rsrs!maxizeit) Then
                cEndZeit = rsrs!maxizeit
            Else
                cEndZeit = ""
            End If
            
            If Not IsNull(rsrs!minizeit) Then
                cStartzeit = rsrs!minizeit
            Else
                cStartzeit = ""
            End If
            
            dStart = TimeValue(cStartzeit)
            dEnde = TimeValue(cEndZeit)
            dEnde = dEnde + TimeValue(gcZeitBlock)
    
            dDauer = dEnde - dStart
            cDauer = Format$(dDauer, "HH:MM")
            
            
            cSQL = "Insert into TERMPRINT_MEP (buchnr,Dauer,Uhrzeit_ende,UHRZEIT) values ("
            cSQL = cSQL & " " & lBuchnr & ",'" & cDauer & "',  '" & Format$(dEnde, "HH:MM") & "',  '" & Format$(dStart, "HH:MM") & "')"
            gdBase.Execute cSQL, dbFailOnError
        
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Update TERMPRINT_MEP M inner join TERMPRINT_EP E on M.buchnr = E.BUCHUNGSNR "
    cSQL = cSQL & " set M.Name = E.Name "
    cSQL = cSQL & ",M.BEDNAME = E.BEDNAME "
    cSQL = cSQL & ",M.BEDNU = E.BEDNU  "
    cSQL = cSQL & ",M.BEHANDLUNG = E.BEHANDLUNG "
    cSQL = cSQL & ",M.DATUM = E.DATUM  "
    cSQL = cSQL & ",M.KABINE = E.KABINE  "
    cSQL = cSQL & ",M.KUERZEL = E.KUERZEL  "
    cSQL = cSQL & ",M.KUNDNR = E.KUNDNR  "
    
    cSQL = cSQL & ",M.TEL = E.TEL  "
    cSQL = cSQL & ",M.MOBILTEL = E.MOBILTEL  "
    cSQL = cSQL & ",M.VORNAME = E.VORNAME  "

    
    cSQL = cSQL & ",M.EMAIL = E.EMAIL  "
    cSQL = cSQL & ",M.adate = E.adate  "
    cSQL = cSQL & ",M.von = E.von  "
    cSQL = cSQL & ",M.bis = E.bis  "
    cSQL = cSQL & ",M.BEDEINTRAG = E.BEDEINTRAG   "
    
    gdBase.Execute cSQL, dbFailOnError
    
        
    reportbildschirm "", "aWKL019e"
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTermineSMS"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub speicherSpeziInfo(cSInfo As String, cBuchnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    If IsNumeric(cBuchnr) Then
    
        cSQL = "Delete from Speziinfo where Buchungsnr = " & cBuchnr
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Insert Into Speziinfo (Buchungsnr,Spezitext,gesehen) values (" & cBuchnr & ",'" & cSInfo & "', False)"
        gdBase.Execute cSQL, dbFailOnError
        
   End If
   
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherSpeziInfo"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub dupliziere_Termin(lOLDBuchnr As Long, lrow As Long, lcol As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cTerminDauer As String
    cTerminDauer = zeigeTerminDauer(CStr(lOLDBuchnr))
    
    If cTerminDauer = "" Then
        cTerminDauer = zeigeRasterDauer(CStr(lOLDBuchnr))
    End If
    
    speicher_Dupli_Termin lOLDBuchnr, lrow, lcol, cTerminDauer
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dupliziere_Termin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loeschenSpeziInfo(cBuchnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    If IsNumeric(cBuchnr) Then
    
        cSQL = "Delete from Speziinfo where Buchungsnr = " & cBuchnr
        gdBase.Execute cSQL, dbFailOnError
        
   End If
   
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loeschenSpeziInfo"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub DruckeTerminBonWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim lAnz As Long
    Dim bReturn As Boolean
    Dim cDrucker As String

    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim cDaten As String
    
    Dim iLenZeile As Integer
    Dim lAnzZeile As Long
    
    Dim cZeichen As String
    Dim cValid As String
    Dim cZiel As String
    Dim iAktCopy                As Integer
    
'    ReDim cDruckZeile(1 To 1) As String
    
    Dim cKundenName As String
    Dim cKdnr As String
    Dim sSQL As String
    
    cValid = "1234567890"
    
    cKdnr = Label2(2).Caption
    cKdnr = Trim$(cKdnr)
    
    If cKdnr = "" Then
        MsgBox "Keinen Kunden ausgewählt! Bon-Druck nicht möglich!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    cZiel = ""
    For lcount = 1 To Len(cKdnr)
        cZeichen = Mid(cKdnr, lcount, 1)
        If cZeichen = " " Then
            Exit For
        Else
            If InStr(cValid, cZeichen) > 0 Then
                cZiel = cZiel & cZeichen
            End If
        End If
    Next lcount
    
    cKdnr = cZiel
    
    cKundenName = fnHoleKundenNameWKL82(cKdnr)
    
    '********************************************
    '*** 1.Schritt: Umschalten auf BonDrucker ***
    '********************************************
    
    setzedrucker gcBonDrucker

    '********************************************************
    '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
    '********************************************************

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz
    
StartPunkt:

    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1


    iLenZeile = 32
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker

    '******************************************************************
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
        ReDim Preserve cDruckZeile(1 To 1) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    '******************************************************************
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
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO-VERSION!"
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
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = "I H R   T E R M I N"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    '******************************************************************
    
    cDaten = "Kundendaten"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "Name:    " & cKundenName
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = "KundNr:  " & cKdnr
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    
    
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    Dim cMobil As String
    cMobil = lookingForKundendaten(cKdnr).Mobiltel
    
    cDaten = "MobilNr: " & cMobil
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    
    
    
    
    
    
    
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    '******************************************************************
    
    Dim sWochentag As String
    sWochentag = WeekdayName(Weekday(DateValue(Text1(0).Text), vbMonday))
        
        
    cDaten = " am: " & sWochentag & ", " & Text1(0).Text
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '******************************************************************
    
    cDaten = " um: " & Text1(1).Text & " Uhr"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '******************************************************************
    cDaten = Label2(0).Caption
    cDaten = Trim$(cDaten)
    
    For lcount = 1 To Len(cDaten)
        cZeichen = Mid(cDaten, lcount, 1)
        If cZeichen = " " Then
            lcount = lcount + 1
            Exit For
        End If
    Next lcount
    
    If lcount < Len(cDaten) Then
        cDaten = Mid(cDaten, lcount, Len(cDaten) - (lcount - 1))
    End If
    
    cDaten = "bei: " & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '******************************************************************
    'alle Behandlungen zum Termin
    
    Dim rsBEH As DAO.Recordset
    Dim cFeld As String
    
    sSQL = " Select distinct(Behandlung) from Termine where buchungsnr = " & Label4.Caption & " "
    
    Set rsBEH = gdBase.OpenRecordset(sSQL)
    If Not rsBEH.EOF Then
        
        rsBEH.MoveFirst
        Do While Not rsBEH.EOF
        
            cFeld = ""
            If Not IsNull(rsBEH!Behandlung) Then
                cFeld = Trim(rsBEH!Behandlung)
            End If
            
            If cFeld <> "" Then
                cDaten = cFeld
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            End If
            
    
        rsBEH.MoveNext
        Loop
    Else
        cFeld = ""
        cFeld = Trim(Text1(3).Text)
        
        If cFeld <> "" Then
            cDaten = cFeld
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        End If
    End If
    rsBEH.Close: Set rsBEH = Nothing
    '******************************************************************
    
    
    
    
    
    
    
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
    
    '******************************************************************
        
    cDaten = "Bitte planen Sie für Ihren"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = "Aufenthalt in unserem Hause"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    Dim lMinuten        As Long
    Dim lStunden        As Long
    
'    'los rechne mal die Zeit für alle Behandlungen aus!
'    Dim cBehandlungszeit As String
'    cBehandlungszeit = ermBehZeit
    
    
    
    lStunden = Fix((CLng(Text1(2).Text) / 60))
    lMinuten = CLng(Text1(2).Text) - (lStunden * 60)
    
    If lStunden = 0 Then
        cDaten = "ca. " & Text1(2).Text & " Minuten ein!"
    Else
        cDaten = "ca. " & lStunden & ":" & Format(CStr(lMinuten), "00") & " h ein!"
    End If
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
    
    '******************************************************************
    
    lese_Storno_Text_in_Array_T1
        
    For lcount = LBound(sStornoTextT1) To UBound(sStornoTextT1)
        
        
        cDaten = sStornoTextT1(lcount)
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    

        
    If gbTerminNoWarn = False Then
    
        lese_Storno_Text_in_Array
        
        For lcount = LBound(sStornoText) To UBound(sStornoText)
            
            
            cDaten = sStornoText(lcount)
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
        
        

    
    End If
    '******************************************************************
    

    
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = Format$(Now, "DD.MM.YYYY") & "                 " & Format$(Now, "HH:MM")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
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
    
    '******************************************************************
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
    
    '******************************************************************
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
    
    '******************************************************************
    
    For lcount = 1 To 9
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
'        OpenDrawer aDeviceName, cEscapeSequenz
    Next lcount
    
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    Erase cDruckZeile
    
BON_SCHNEIDEN:

    'Kassenbon abschneiden
    If gbAPI Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(105)
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
ENDE:

    '...und tschüß!

    '*******************************************************
    '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
    '*******************************************************
    
    If gb2BONTermin Then
        If iAktCopy < 2 Then
            GoTo StartPunkt
        End If
    End If
    
    setzedrucker gcListenDrucker

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTerminBonWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Function ermBehZeit() As String
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim lcount          As Long
    Dim cLBSatz         As String
    Dim cBezeich        As String
    Dim cDauer          As String
    Dim rsrs            As DAO.Recordset
    Dim sSQL            As String
    
    ermBehZeit = "0"
    
    If Text1(3).Text <> "" Then
        Dim sArray() As String
        sArray = Split(Text1(3).Text, Chr$(13) & Chr$(10))
        
        For i = 0 To UBound(sArray) - 1
        
            sSQL = " Select * from TERM_STD where BEZEICH = '" & sArray(i) & "'"
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
            
                If Not IsNull(rsrs!DAUER) Then
                    cDauer = Trim(rsrs!DAUER)
                    ermBehZeit = Trim$(Str$(Val(ermBehZeit) + Val(cDauer)))
                End If
            Else
                ermBehZeit = "0"
                rsrs.Close: Set rsrs = Nothing
                Exit Function
                
            End If
            rsrs.Close: Set rsrs = Nothing
            
        Next i
    End If
    
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBehZeit"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub fuelle_Artikel_Array()
On Error GoTo LOKAL_ERROR

    Dim i   As Integer
    Dim j   As Integer
    
    For i = 0 To 9
        gsArtikelArray(i) = ""
    Next i
    
    j = 0
    
    If Text1(3).Text <> "" Then
        Dim sArray() As String
        sArray = Split(Text1(3).Text, Chr$(13) & Chr$(10))
        
        For i = 0 To UBound(sArray) - 1
            gsArtikelArray(j) = erm_Artnr_aus_TERM_STD(sArray(i))
            j = j + 1
        Next i
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_Artikel_Array"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function erm_Artnr_aus_TERM_STD(sBehandlung) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As DAO.Recordset
    
    erm_Artnr_aus_TERM_STD = ""
    
    sSQL = "Select ARTNR from TERM_STD where BEZEICH = '" & sBehandlung & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!artnr) Then
            erm_Artnr_aus_TERM_STD = rsrs!artnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erm_Artnr_aus_TERM_STD"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnHoleKundenNameWKL82(cKdnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKdName As String
    Dim cKdVorname As String
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!vorname) Then
            cKdVorname = rsrs!vorname
        Else
            cKdVorname = ""
        End If
        If Not IsNull(rsrs!name) Then
            cKdName = rsrs!name
        Else
            cKdName = ""
        End If
        
        
'        cKdVorname = Trim$(cKdVorname)
'        If Len(cKdVorname) > 0 Then
'            cKdVorname = UCase$(Left(cKdVorname, 1)) & ". "
'        End If
'        cKdName = Trim$(cKdName)


        fnHoleKundenNameWKL82 = cKdVorname & " " & cKdName
    Else
        fnHoleKundenNameWKL82 = ""
    End If
    rsrs.Close: Set rsrs = Nothing
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleKundenNameWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Function fnHoleKundenNameVollWKL82(cKdnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKdName As String
    Dim cKdVorname As String
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!vorname) Then
            cKdVorname = rsrs!vorname
        Else
            cKdVorname = ""
        End If
        If Not IsNull(rsrs!name) Then
            cKdName = rsrs!name
        Else
            cKdName = ""
        End If
        cKdVorname = Trim$(cKdVorname)
        cKdName = Trim$(cKdName)
        fnHoleKundenNameVollWKL82 = cKdVorname & " " & cKdName
    Else
        fnHoleKundenNameVollWKL82 = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleKundenNameVollWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub EinzelDatenMitarbeiterWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim lcol As Long
    Dim lrow As Long
    Dim cbednu As String
    Dim cDatum As String
    Dim czeit As String
    Dim cOrt As String
    Dim cAnzeige As String
    Dim cBuchnr As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = lcol
    
    cAnzeige = MSFlexGrid1.Text
    cAnzeige = Trim$(cAnzeige)
    
    HoleAllePflegeOrteWKL82
    fuellecboAbwesend
    
    'Mitarbeiter
    If Len(MSFlexGrid1.Text) > 0 Then

        Label2(0).Caption = Mid(MSFlexGrid1.Text, 51, Len(MSFlexGrid1.Text) - 50)
        cbednu = Mid(MSFlexGrid1.Text, 51, Len(MSFlexGrid1.Text) - 50)
        
        Faerbebed Trim$(Left(Label2(0).Caption, 3)), Label2(0)
        Combo1.Text = cbednu
    Else
        Label2(0).Caption = Combo1.Text
        cbednu = Combo1.Text
    End If
    
    List6.Clear
    
    'Datum
    Text1(0).Text = Right(Combo2.Text, 8)
    cDatum = Right(Combo2.Text, 8)
    
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    czeit = MSFlexGrid1.Text
    Text1(1).Text = czeit
    Text1(2).Text = Hour(gcZeitBlock) * 60 + Minute(gcZeitBlock)
    Text1(3).Text = ""
    
    'Ort
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = lcol
    cOrt = MSFlexGrid1.Text
    Label2(1).Caption = cOrt
    
    If cAnzeige <> "" Then
        cBuchnr = fnHoleBuchNr(cDatum, czeit, cbednu, cOrt)
        If cBuchnr <> "" Then
            Label4.Caption = cBuchnr
        Else
            Label4.Caption = "-1"
        End If
    Else
        Label4.Caption = "-1"
    End If
    
    'Mitarbeiter, der den Eintrag vorgenommen hatte
    If Label4.Caption <> "-1" Then
        cSQL = "Select bedeintrag from TERMINE where BUCHUNGSNR = " & Label4.Caption & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                If Not IsNull(rsrs!bedeintrag) Then
                    Text1(4).Text = Trim(rsrs!bedeintrag)
                    
                End If
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        cSQL = "Select * from TERMINE_ANL where BUCHUNGSNR = " & Label4.Caption & " "
        Label12.Caption = ""
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!bedeintrag) Then
                Label12.Caption = Label12.Caption & "erstellt von: " & rsrs!bedname & " am: " & rsrs!ANLAGE_DATUM
            End If
            Label12.Refresh
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        
        
        
        
    End If
    
    If Label4.Caption <> "-1" Then
        LeseTerminBuchungWKL82
    End If
    
    If cAnzeige = "" Then
        If Label3(4).Caption <> "" Then
            Label2(2).Caption = Label3(4).Caption
            Label2(2).BackColor = Label3(4).BackColor
'            MsgBox Label2(2).Caption
            gckundnr = Left(Label2(2).Caption, InStr(1, Label2(2).Caption, " "))
            gckundnr = Trim$(gckundnr)
            
            
            
            SucheBediener gckundnr, List6
            SucheGelöschteTermine gckundnr
            DS_Unterschrieben gckundnr
            
            gckundnr = ""
            
        End If
    Else
        gckundnr = Left(Label2(2).Caption, InStr(1, Label2(2).Caption, " "))
        gckundnr = Trim$(gckundnr)
            
        lblUnter.Visible = False
                
        If SucheUnter(gckundnr) Then
            lblUnter.ForeColor = glWarn
            lblUnter.Visible = True
        End If
        
        DS_Unterschrieben gckundnr
    
    End If
    
    Frame4.Visible = True
    Frame9.Visible = False
    Frame3.Visible = False
    Command3(0).Visible = False
    Command3(3).Visible = False
    Command3(5).Visible = False
    Combo12.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EinzelDatenMitarbeiterWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function fnHoleBuchNr(cDatum As String, czeit As String, cbednu As String, cOrt As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim lDatum As Long
    Dim cBed As String
    Dim ctmp As String
    Dim SFarbe As String
    Dim cKürzel As String
    Dim cName As String
    Dim cVorname As String
    
    fnHoleBuchNr = ""
    
    lDatum = DateValue(cDatum)
    cBed = Left(cbednu, 3)
    cBed = Trim$(cBed)
    
    Label2(2).Caption = ""
    Label2(2).BackColor = glH1
    lblUnter.Visible = False
    
    cSQL = "Select * from TERMINE where DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & "and UHRZEIT = '" & czeit & "' "
    cSQL = cSQL & "and BEDNU = " & cBed & " "
    cSQL = cSQL & "and KABINE = '" & cOrt & "' "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        ctmp = ""
        If Not IsNull(rsrs!Kundnr) Then
            ctmp = rsrs!Kundnr
        Else
            ctmp = ""
        End If
        
        If IsNumeric(ctmp) Then
            SFarbe = ermFarbe(Trim(ctmp))
            If Trim(SFarbe) = "0" Then
                Label2(2).BackColor = glH1
            Else
                Label2(2).BackColor = glfarbe(SFarbe)
            End If
        End If
        
        
    
        cKürzel = lookingForKundendaten(ctmp).Kuerzel
        cName = lookingForKundendaten(ctmp).nachname
        cVorname = lookingForKundendaten(ctmp).vorname
    
        Label2(2).Caption = ctmp & "  "
        Label2(2).Caption = Label2(2).Caption & Space$(5 - Len(cKürzel)) & cKürzel & "  "
        Label2(2).Caption = Label2(2).Caption & cName & ", "
        Label2(2).Caption = Label2(2).Caption & cVorname
    
'        If Not IsNull(rsrs!Kuerzel) Then
'            ctmp = ctmp & " " & rsrs!Kuerzel
'        Else
'            ctmp = ctmp & ""
'        End If
'        ctmp = ctmp & " " & lookingForKundendaten(ctmp).nachname

        

'        Label2(2).Caption = ctmp
        Label2(2).Refresh
        
        If Not IsNull(rsrs!BUCHUNGSNR) Then
            ctmp = rsrs!BUCHUNGSNR
        Else
            ctmp = ""
        End If
    Else
        ctmp = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    fnHoleBuchNr = ctmp
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleBuchNr"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub fnHoleKundendaten(cDatum As String, czeit As String, cbednu As String, cOrt As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lDatum As Long
    Dim cBed As String
    Dim ctmp As String
    Dim lBuchnr As Long
    Dim SFarbe As String
    
    lDatum = DateValue(cDatum)
    cBed = Left(cbednu, 3)
    cBed = Trim$(cBed)
    
    If Not IsNumeric(cBed) Then
        Exit Sub
    End If
    
    cSQL = "Select * from TERMINE where DATUM = " & Trim$(Str$(lDatum)) & " "
    cSQL = cSQL & " and UHRZEIT = '" & czeit & "' "
    cSQL = cSQL & " and BEDNU = " & cBed & " "
    cSQL = cSQL & " and KABINE = '" & cOrt & "' "
    
    Label8.Caption = ""
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        ctmp = ""
        
        If Not IsNull(rsrs!BUCHUNGSNR) Then
            lBuchnr = rsrs!BUCHUNGSNR
            Label8.Caption = rsrs!BUCHUNGSNR
            Command3(4).Caption = "Termin kopieren"
            Command3(4).ForeColor = vbBlack
        End If
        If Not IsNull(rsrs!Kundnr) Then
            Label6.Caption = rsrs!Kundnr
            
            If IsNumeric(rsrs!Kundnr) Then
                SFarbe = ermFarbe(Trim(rsrs!Kundnr))
                If Trim(SFarbe) = "0" Then
                    Label6.BackColor = glH1
                Else
                    Label6.BackColor = glfarbe(SFarbe)
                End If
            End If
            
            If Not IsNull(rsrs!Kuerzel) Then
                Label6.Caption = Label6.Caption & " " & rsrs!Kuerzel & vbCrLf
            End If
            Label6.Caption = Label6.Caption & " " & lookingForKundendaten(Trim(rsrs!Kundnr)).vorname
            Label6.Caption = Label6.Caption & " " & lookingForKundendaten(Trim(rsrs!Kundnr)).nachname & vbCrLf
            Label6.Caption = Label6.Caption & " " & lookingForKundendaten(Trim(rsrs!Kundnr)).Plz
            Label6.Caption = Label6.Caption & " " & lookingForKundendaten(Trim(rsrs!Kundnr)).Ort & vbCrLf
            Label6.Caption = Label6.Caption & " " & lookingForKundendaten(Trim(rsrs!Kundnr)).strasse & vbCrLf & vbCrLf
            Label6.Caption = Label6.Caption & " Telefon: " & lookingForKundendaten(Trim(rsrs!Kundnr)).telefon & vbCrLf
            Label6.Caption = Label6.Caption & " Mobil: " & lookingForKundendaten(Trim(rsrs!Kundnr)).Mobiltel & vbCrLf
            Label6.Caption = Label6.Caption & " Email: " & lookingForKundendaten(Trim(rsrs!Kundnr)).Email & vbCrLf
            Label6.Caption = Label6.Caption & " Geb: " & lookingForKundendaten(Trim(rsrs!Kundnr)).GEBDATUM & vbCrLf
            Label6.Caption = Label6.Caption & " Beruf: " & lookingForKundendaten(Trim(rsrs!Kundnr)).KTEXT2 & vbCrLf
            
            

        End If
        
        
        Label6.Refresh
    Else
        ctmp = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from SPEZIINFO where BUCHUNGSNR = " & lBuchnr
    Text3.Text = ""
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SPEZITEXT) Then
            Text3.Text = rsrs!SPEZITEXT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    

'    cSQL = cSQL & " BUCHUNGSNR float"
'    cSQL = cSQL & ", ANLAGE_DATUM Datetime"
'    cSQL = cSQL & ", UHRZEIT varchar(10)"
'    cSQL = cSQL & ", BEDEINTRAG smallint "
'    cSQL = cSQL & ", BEDNAME varchar(32) "
    
    
    cSQL = "Select * from TERMINE_ANL where BUCHUNGSNR = " & lBuchnr
    Label11.Caption = ""
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!bedeintrag) Then
            Label11.Caption = Label11.Caption & "erstellt von: " & rsrs!bedname & " am: " & rsrs!ANLAGE_DATUM
        End If
        Label11.Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    
    
'    lese_Termin_Optionen
            
    If gbTerm_InfoDauerh = False Then
        cSQL = "Update SPEZIINFO set gesehen = True where BUCHUNGSNR = " & lBuchnr
        gdBase.Execute cSQL, dbFailOnError
    End If
            
            
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleKundendaten"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeDialogWKL82() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim dWert As Double
    Dim cFeld1 As String
    
    fnPruefeEingabeDialogWKL82 = 0
    
    If Trim$(Label2(0).Caption) = "" Then
        fnPruefeEingabeDialogWKL82 = 1
        Exit Function
    End If
    If Trim$(Label2(1).Caption) = "" Then
        fnPruefeEingabeDialogWKL82 = 2
        Exit Function
    End If
    
    
    If Trim$(Text1(0).Text) = "" Then
        fnPruefeEingabeDialogWKL82 = 4
        Exit Function
    End If
    
    If Not IsDate(Trim$(Text1(0).Text)) Then
        fnPruefeEingabeDialogWKL82 = 4
        Exit Function
    Else
        dWert = DateValue(Text1(0).Text)
        Text1(0).Text = Format$(dWert, "DD.MM.YYYY")
    End If
    
    If Trim$(Text1(1).Text) = "" Then
        fnPruefeEingabeDialogWKL82 = 5
        Exit Function
    End If
    
    If Left(Text1(1).Text, 2) > 23 Then
        fnPruefeEingabeDialogWKL82 = 5
        Exit Function
    End If
    If Right(Text1(1).Text, 2) > 59 Then
        fnPruefeEingabeDialogWKL82 = 5
        Exit Function
    End If
    
    cFeld1 = Text1(1).Text
    cFeld1 = Trim$(cFeld1)
    dWert = TimeValue(cFeld1)
    Text1(1).Text = Format$(dWert, "HH:MM")
    
    Dim lCheck As Long
    lCheck = Val(Right(gcZeitBlock, 2))
    
    If Val(Right(Text1(1).Text, 2)) / lCheck <> Fix(Val(Right(Text1(1).Text, 2)) / lCheck) Then
        fnPruefeEingabeDialogWKL82 = 5
        Exit Function
    End If
    
    'sind wir über dauer + start
    
    
    
    
    
    'uhrzeit abwesend
    
    If Combo5.Text <> "" Then
        If Left(Combo5.Text, 2) > 23 Then
            fnPruefeEingabeDialogWKL82 = 15
            Exit Function
        End If
        If Right(Combo5.Text, 2) > 59 Then
            fnPruefeEingabeDialogWKL82 = 15
            Exit Function
        End If
        
        cFeld1 = Combo5.Text
        cFeld1 = Trim$(cFeld1)
        dWert = TimeValue(cFeld1)
        Combo5.Text = Format$(dWert, "HH:MM")
        
        If Val(Right(Combo5.Text, 2)) / lCheck <> Fix(Val(Right(Combo5.Text, 2)) / lCheck) Then
            fnPruefeEingabeDialogWKL82 = 15
            Exit Function
        End If
        
        Dim cBlock As String
        Dim cMaxZeit As String
        Dim ccheckZeit As String
        Dim Lend As Long
        
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
        cMaxZeit = MSFlexGrid1.Text
        cMaxZeit = SwapStr(cMaxZeit, ":", "")
        
        cBlock = gcZeitBlock
        cBlock = SwapStr(cBlock, ":", "")
        
        ccheckZeit = Combo5.Text
        ccheckZeit = SwapStr(ccheckZeit, ":", "")
        
        Dim lKontrolltime As Long
        Dim cKontrolltime As String
        lKontrolltime = Val(cMaxZeit) + Val(cBlock)
        If Right(lKontrolltime, 2) > 59 Then
            cKontrolltime = Left(CStr(lKontrolltime), 2) + 1 & "00"
            lKontrolltime = Val(cKontrolltime)
        End If
        
        If Val(ccheckZeit) > lKontrolltime Then
            Combo5.Text = Left(CStr(lKontrolltime), 2) & ":" & Right(CStr(lKontrolltime), 2)

        End If
    End If
    
    'kunde
    If Trim$(Label2(2).Caption) = "" And Combo5.Text = "" Then
        fnPruefeEingabeDialogWKL82 = 3
        Exit Function
    End If
    
    'kunde oder abwesenheit
    If Trim$(Label2(2).Caption) = "" And Combo3.Text = "" Then
        fnPruefeEingabeDialogWKL82 = 7
        Exit Function
    End If
    
    
    If Trim$(Text1(2).Text) = "" Then
        fnPruefeEingabeDialogWKL82 = 6
        Exit Function
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function fnPruefeEingabeDialog_abwesend() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim dWert As Double
    Dim cFeld1 As String
    
    fnPruefeEingabeDialog_abwesend = 0
    
    If Trim$(Label2(11).Caption) = "" Then
        fnPruefeEingabeDialog_abwesend = 1
        Exit Function
    End If
    
    
    
    
    If Trim$(Label3(22).Caption) = "" Then
        fnPruefeEingabeDialog_abwesend = 4
        Exit Function
    End If
    
    If Not IsDate(Trim$(Label3(22).Caption)) Then
        fnPruefeEingabeDialog_abwesend = 4
        Exit Function
    Else
        dWert = DateValue(Label3(22).Caption)
        Label3(22).Caption = Format$(dWert, "DD.MM.YYYY")
    End If
    
    If Trim$(Combo11.Text) = "" Then
        fnPruefeEingabeDialog_abwesend = 5
        Exit Function
    End If
    
    If Left(Combo11.Text, 2) > 23 Then
        fnPruefeEingabeDialog_abwesend = 5
        Exit Function
    End If
    If Right(Combo11.Text, 2) > 59 Then
        fnPruefeEingabeDialog_abwesend = 5
        Exit Function
    End If
    
    cFeld1 = Combo11.Text
    cFeld1 = Trim$(cFeld1)
    dWert = TimeValue(cFeld1)
    Combo11.Text = Format$(dWert, "HH:MM")
    
    Dim lCheck As Long
    lCheck = Val(Right(gcZeitBlock, 2))
    
    If Val(Right(Combo11.Text, 2)) / lCheck <> Fix(Val(Right(Combo11.Text, 2)) / lCheck) Then
        fnPruefeEingabeDialog_abwesend = 5
        Exit Function
    End If
    
    'sind wir über dauer + start
    
    
    
    
    
    'uhrzeit abwesend
    
    If Combo8.Text <> "" Then
        If Left(Combo8.Text, 2) > 23 Then
            fnPruefeEingabeDialog_abwesend = 15
            Exit Function
        End If
        If Right(Combo8.Text, 2) > 59 Then
            fnPruefeEingabeDialog_abwesend = 15
            Exit Function
        End If
        
        cFeld1 = Combo8.Text
        cFeld1 = Trim$(cFeld1)
        dWert = TimeValue(cFeld1)
        Combo8.Text = Format$(dWert, "HH:MM")
        
        If Val(Right(Combo8.Text, 2)) / lCheck <> Fix(Val(Right(Combo8.Text, 2)) / lCheck) Then
            fnPruefeEingabeDialog_abwesend = 15
            Exit Function
        End If
        
        Dim cBlock As String
        Dim cMaxZeit As String
        Dim ccheckZeit As String
        Dim Lend As Long
        
        
        cMaxZeit = gcEndeZeit
        cMaxZeit = SwapStr(cMaxZeit, ":", "")
        
        cBlock = gcZeitBlock
        cBlock = SwapStr(cBlock, ":", "")
        
        ccheckZeit = Combo8.Text
        ccheckZeit = SwapStr(ccheckZeit, ":", "")
        
        Dim lKontrolltime As Long
        Dim cKontrolltime As String
        lKontrolltime = Val(cMaxZeit) + Val(cBlock)
        If Right(lKontrolltime, 2) > 59 Then
            cKontrolltime = Left(CStr(lKontrolltime), 2) + 1 & "00"
            lKontrolltime = Val(cKontrolltime)
        End If
        
        If Val(ccheckZeit) > lKontrolltime Then
            Combo8.Text = Left(CStr(lKontrolltime), 2) & ":" & Right(CStr(lKontrolltime), 2)

        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialog_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function fnPruefeEingabeMitarbeiter_abwesend() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim lHeute As Long
    Dim lDatum As Long
    Dim lDauer As Long
    
    fnPruefeEingabeMitarbeiter_abwesend = 0
    
    'Datum
    ctmp = Label3(22).Caption
    ctmp = Trim$(ctmp)
    If ctmp = "" Or Not IsDate(ctmp) Then
        fnPruefeEingabeMitarbeiter_abwesend = 1
        Exit Function
    End If
    
    lHeute = DateValue(Now)
    lDatum = DateValue(ctmp)
    
    If lDatum < lHeute Then
        fnPruefeEingabeMitarbeiter_abwesend = 99
        Exit Function
    End If
    
    'Uhrzeit
    ctmp = Combo11.Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeMitarbeiter_abwesend = 2
        Exit Function
    Else
        If Len(ctmp) <> 5 Then
            fnPruefeEingabeMitarbeiter_abwesend = 2
            Exit Function
        End If
        
        For lcount = 1 To 5
            cZeichen = Mid(ctmp, lcount, 1)
            Select Case lcount
                Case Is = 1
                    If cZeichen < "0" Or cZeichen > "2" Then
                        fnPruefeEingabeMitarbeiter_abwesend = 2
                        Exit Function
                    End If
                Case Is = 2
                    If cZeichen < "0" Or cZeichen > "9" Then
                        fnPruefeEingabeMitarbeiter_abwesend = 2
                        Exit Function
                    End If
                Case Is = 3
                    If cZeichen <> ":" Then
                        fnPruefeEingabeMitarbeiter_abwesend = 2
                        Exit Function
                    End If
                Case Is = 4
                    If cZeichen < "0" Or cZeichen > "5" Then
                        fnPruefeEingabeMitarbeiter_abwesend = 2
                        Exit Function
                    End If
                Case Is = 5
                    If cZeichen < "0" Or cZeichen > "9" Then
                        fnPruefeEingabeMitarbeiter_abwesend = 2
                        Exit Function
                    End If
            End Select
        Next lcount
        
        If TimeValue(ctmp) < TimeValue(gcStartZeit) Then
            fnPruefeEingabeMitarbeiter_abwesend = 21
            Exit Function
        End If
        
        If TimeValue(ctmp) >= TimeValue(gcEndeZeit) Then
            fnPruefeEingabeMitarbeiter_abwesend = 22
            Exit Function
        End If
    End If
    
    ctmp = Right(ctmp, 2)
    lcount = Val(ctmp)
    
    
    Dim lCheck As Long
    lCheck = Val(Right(gcZeitBlock, 2))
    
    If lcount / lCheck <> Fix(lcount / lCheck) Then
        fnPruefeEingabeMitarbeiter_abwesend = 2
        Exit Function
    End If
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeMitarbeiter_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function PruefeDatumZeit(lrow As Long) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim lHeute As Long
    Dim lDatum As Long
    Dim lDauer As Long
    
    PruefeDatumZeit = 0
    
    'Datum
    ctmp = Right(Combo2.Text, 8)
    ctmp = Trim$(ctmp)
    If ctmp = "" Or Not IsDate(ctmp) Then
        PruefeDatumZeit = 1
        Exit Function
    End If
    
    lHeute = DateValue(Now)
    lDatum = DateValue(ctmp)
    
    If lDatum < lHeute Then
        PruefeDatumZeit = 99
        Exit Function
    End If
    
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    ctmp = Trim$(MSFlexGrid1.Text)
    If ctmp = "" Then
        PruefeDatumZeit = 2
        Exit Function
    Else
        If Len(ctmp) <> 5 Then
            PruefeDatumZeit = 2
            Exit Function
        End If
        
        For lcount = 1 To 5
            cZeichen = Mid(ctmp, lcount, 1)
            Select Case lcount
                Case Is = 1
                    If cZeichen < "0" Or cZeichen > "2" Then
                        PruefeDatumZeit = 2
                        Exit Function
                    End If
                    
                Case Is = 2
                    If cZeichen < "0" Or cZeichen > "9" Then
                        PruefeDatumZeit = 2
                        Exit Function
                    End If
                    
                Case Is = 3
                    If cZeichen <> ":" Then
                        PruefeDatumZeit = 2
                        Exit Function
                    End If
                    
                Case Is = 4
                    If cZeichen < "0" Or cZeichen > "5" Then
                        PruefeDatumZeit = 2
                        Exit Function
                    End If
                    
                Case Is = 5
                    If cZeichen < "0" Or cZeichen > "9" Then
                        PruefeDatumZeit = 2
                        Exit Function
                    End If
            End Select
        Next lcount
        
        If TimeValue(ctmp) < TimeValue(gcStartZeit) Then
            PruefeDatumZeit = 21
            Exit Function
        End If
        
        If TimeValue(ctmp) >= TimeValue(gcEndeZeit) Then
            PruefeDatumZeit = 22
            Exit Function
        End If
    End If
    
    ctmp = Right(ctmp, 2)
    lcount = Val(ctmp)
    
    Dim lCheck As Long
    lCheck = Val(Right(gcZeitBlock, 2))
    
    If lcount / lCheck <> Fix(lcount / lCheck) Then
        PruefeDatumZeit = 2
        Exit Function
    End If
    
    'Dauer
    lDauer = Hour(gcZeitBlock) * 60 + Minute(gcZeitBlock)
    ctmp = Hour(gcZeitBlock) * 60 + Minute(gcZeitBlock)
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        PruefeDatumZeit = 3
        Exit Function
    End If
    
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr("1234567890", cZeichen) = 0 Then
            PruefeDatumZeit = 3
            Exit Function
        End If
    Next lcount
    
    lcount = Val(ctmp)
    If lcount < lDauer Or lcount / lDauer <> Fix(lcount / lDauer) Then
    
        'Dauer der Behandlung passt nicht zum Zeitblock
        'also aufrunden
        
        
        
'        fnPruefeEingabeMitarbeiterWKL82 = 3
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PruefeDatumZeit"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function fnPruefeEingabeMitarbeiterWKL82() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim lHeute As Long
    Dim lDatum As Long
    Dim lDauer As Long
    
    fnPruefeEingabeMitarbeiterWKL82 = 0
    
    
    
    
    
    
    
    
    
    
    
    Dim cKund As String
    Dim cKundnr As String
    Dim cbednu As String
    
    cbednu = Trim(Left(Label2(0).Caption, 3))
    
    If Label2(2).Caption = Label2(0).Caption Then
        cKundnr = ermBedKundnr(cbednu)

    Else
        cKund = Label2(2).Caption
        cKundnr = ""
        For lcount = 1 To Len(cKund)
            If InStr("1234567890", Mid(cKund, lcount, 1)) > 0 Then
                cKundnr = cKundnr & Mid(cKund, lcount, 1)
            End If
        Next lcount
    
        lcount = InStr(1, cKund, " ")
    
    End If
    
    If cKundnr = "" Or Not IsNumeric(cKundnr) Then
        fnPruefeEingabeMitarbeiterWKL82 = 44
        Exit Function
    End If
    
    
    
    
    
    
    
    
    
    
    
    'Datum
    ctmp = Text1(0).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Or Not IsDate(ctmp) Then
        fnPruefeEingabeMitarbeiterWKL82 = 1
        Exit Function
    End If
    
    lHeute = DateValue(Now)
    lDatum = DateValue(ctmp)
    
    If lDatum < lHeute Then
        fnPruefeEingabeMitarbeiterWKL82 = 99
        Exit Function
    End If
    
    'Uhrzeit
    ctmp = Text1(1).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeMitarbeiterWKL82 = 2
        Exit Function
    Else
        If Len(ctmp) <> 5 Then
            fnPruefeEingabeMitarbeiterWKL82 = 2
            Exit Function
        End If
        
        For lcount = 1 To 5
            cZeichen = Mid(ctmp, lcount, 1)
            Select Case lcount
                Case Is = 1
                    If cZeichen < "0" Or cZeichen > "2" Then
                        fnPruefeEingabeMitarbeiterWKL82 = 2
                        Exit Function
                    End If
                    
                Case Is = 2
                    If cZeichen < "0" Or cZeichen > "9" Then
                        fnPruefeEingabeMitarbeiterWKL82 = 2
                        Exit Function
                    End If
                    
                Case Is = 3
                    If cZeichen <> ":" Then
                        fnPruefeEingabeMitarbeiterWKL82 = 2
                        Exit Function
                    End If
                    
                Case Is = 4
                    If cZeichen < "0" Or cZeichen > "5" Then
                        fnPruefeEingabeMitarbeiterWKL82 = 2
                        Exit Function
                    End If
                    
                Case Is = 5
                    If cZeichen < "0" Or cZeichen > "9" Then
                        fnPruefeEingabeMitarbeiterWKL82 = 2
                        Exit Function
                    End If
            End Select
        Next lcount
        
        If TimeValue(ctmp) < TimeValue(gcStartZeit) Then
            fnPruefeEingabeMitarbeiterWKL82 = 21
            Exit Function
        End If
        
        If TimeValue(ctmp) >= TimeValue(gcEndeZeit) Then
            fnPruefeEingabeMitarbeiterWKL82 = 22
            Exit Function
        End If
    End If
    
    ctmp = Right(ctmp, 2)
    lcount = Val(ctmp)
    
    
    Dim lCheck As Long
    lCheck = Val(Right(gcZeitBlock, 2))
    
    If lcount / lCheck <> Fix(lcount / lCheck) Then
        fnPruefeEingabeMitarbeiterWKL82 = 2
        Exit Function
    End If
    
    'Dauer
    
    lDauer = Hour(gcZeitBlock) * 60 + Minute(gcZeitBlock)
    
    ctmp = Text1(2).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeMitarbeiterWKL82 = 3
        Exit Function
    End If
    
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr("1234567890", cZeichen) = 0 Then
            fnPruefeEingabeMitarbeiterWKL82 = 3
            Exit Function
        End If
    Next lcount
    
    lcount = Val(ctmp)
    If lcount < lDauer Or lcount / lDauer <> Fix(lcount / lDauer) Then
    
        'Dauer der Behandlung passt nicht zum Zeitblock
        'also aufrunden
        
        
        
'        fnPruefeEingabeMitarbeiterWKL82 = 3
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeMitarbeiterWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function fnPruefeVakanzMitarbeiterWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cUhrZeit As String
    Dim cDauer As String
    Dim cZeitSpanne As String
    Dim cbednu As String
    Dim dStartzeit As Double
    Dim dEndeZeit As Double
    
    Dim iAlleTage As Integer
    Dim i As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    fnPruefeVakanzMitarbeiterWKL82 = 0
    
    cbednu = Left(Label2(0).Caption, 3)
    cbednu = Trim$(Str$(Val(cbednu)))
    
    cDatum = Text1(0).Text
    lDatum = DateValue(cDatum)

    
    'Zeitspanne prüfen begin
    cUhrZeit = Text1(1).Text
    If Label2(2).Caption = Label2(0).Caption Then
        cZeitSpanne = Combo5.Text
    Else
        cDauer = Text1(2).Text
        cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
        dStartzeit = TimeValue(cUhrZeit)
        dEndeZeit = dStartzeit + TimeValue(cZeitSpanne)
        cZeitSpanne = Format$(dEndeZeit, "HH:MM")
    End If
    'Zeitspanne ende
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(5).Caption))
    lDatBis = CLng(DateValue(Label3(12).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo6.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
    
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
    
        cSQL = "Select * from TERMINE where BEDNU = " & cbednu & " and DATUM = " & i & " "
        cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            fnPruefeVakanzMitarbeiterWKL82 = 1
        End If
        rsrs.Close: Set rsrs = Nothing
sprung:
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeVakanzMitarbeiterWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnPruefeVakanzKundeWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum      As String
    Dim lDatum      As Long
    Dim cUhrZeit    As String
    Dim cDauer      As String
    Dim cZeitSpanne As String
    Dim cKundnr     As String
    Dim dStartzeit  As Double
    Dim dEndeZeit   As Double
    
    Dim iAlleTage   As Integer
    Dim i           As Long
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim lcount      As Long
    
    fnPruefeVakanzKundeWKL82 = 0
    
    Dim cKund As String
    Dim cbednu As String
    
    cbednu = Trim(Left(Label2(0).Caption, 3))
    
    
    If Label2(2).Caption = Label2(0).Caption Then
        cKundnr = ermBedKundnr(cbednu)
'        cKuerzel = ermBedKuerzel(cKundnr)
    Else
        cKund = Label2(2).Caption
        cKundnr = ""
        For lcount = 1 To Len(cKund)
            If InStr("1234567890", Mid(cKund, lcount, 1)) > 0 Then
                cKundnr = cKundnr & Mid(cKund, lcount, 1)
            End If
        Next lcount
    
        lcount = InStr(1, cKund, " ")
    
    End If
    
    
    
    
    
    
    
    
    
'    cKundnr = Left(Label2(0).Caption, 3)
'    cKundnr = Trim$(Str$(Val(cKundnr)))
    
    cDatum = Text1(0).Text
    lDatum = DateValue(cDatum)

    
    'Zeitspanne prüfen begin
    cUhrZeit = Text1(1).Text
    If Label2(2).Caption = Label2(0).Caption Then
        cZeitSpanne = Combo5.Text
    Else
        cDauer = Text1(2).Text
        cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
        dStartzeit = TimeValue(cUhrZeit)
        dEndeZeit = dStartzeit + TimeValue(cZeitSpanne)
        cZeitSpanne = Format$(dEndeZeit, "HH:MM")
    End If
    'Zeitspanne ende
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(5).Caption))
    lDatBis = CLng(DateValue(Label3(12).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo6.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
    
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
    
        cSQL = "Select * from TERMINE where KUNDNR = " & cKundnr & " and DATUM = " & i & " "
        cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            fnPruefeVakanzKundeWKL82 = 1
        End If
        rsrs.Close: Set rsrs = Nothing
sprung:
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeVakanzKundeWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function PruefeMitarbeiter(lOLDBuchnr As Long, lrow As Long, lcol As Long, cDauer As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cUhrZeit As String
    Dim cZeitSpanne As String
    Dim cbednu As String
    Dim dStartzeit As Double
    Dim dEndeZeit As Double
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    PruefeMitarbeiter = 0
        
    cbednu = ermMitarbeiterausTermin(lOLDBuchnr)
    
    lDatum = DateValue(Right(Combo2.Text, 8))

    'Zeitspanne prüfen begin
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    cUhrZeit = MSFlexGrid1.Text
    
    cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
    dStartzeit = TimeValue(cUhrZeit)
    dEndeZeit = dStartzeit + TimeValue(cZeitSpanne)
    cZeitSpanne = Format$(dEndeZeit, "HH:MM")
    'Zeitspanne ende

    cSQL = "Select * from TERMINE where BEDNU = " & cbednu & " and DATUM = " & lDatum & " "
    cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        PruefeMitarbeiter = 1
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PruefeMitarbeiter"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermMitarbeiterausTermin(lOLDBuchnr As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    ermMitarbeiterausTermin = 0
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select BEDNU from TERMINE where BUCHUNGSNR = " & lOLDBuchnr & ""
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNU) Then
            ermMitarbeiterausTermin = rsrs!BEDNU
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermMitarbeiterausTermin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermBednameausTermin(lOLDBuchnr As Long) As String
    On Error GoTo LOKAL_ERROR
    
    ermBednameausTermin = ""
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select Bedname from TERMINE where BUCHUNGSNR = " & lOLDBuchnr & ""
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!bedname) Then
            ermBednameausTermin = rsrs!bedname
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBednameausTermin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermKundeausTermin(lOLDBuchnr As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    ermKundeausTermin = 0
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select KUNDNR from TERMINE where BUCHUNGSNR = " & lOLDBuchnr & ""
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Kundnr) Then
            ermKundeausTermin = rsrs!Kundnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKundeausTermin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnPruefeVakanzMitarbeiter_abwesend()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cUhrZeit As String
    Dim cZeitSpanne As String
    Dim cbednu As String
    Dim dStartzeit As Double
    Dim dEndeZeit As Double
    
    Dim iAlleTage As Integer
    Dim i As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    fnPruefeVakanzMitarbeiter_abwesend = 0
    
    cbednu = Left(Label2(11).Caption, 3)
    cbednu = Trim$(Str$(Val(cbednu)))
    
    lDatum = DateValue(Label3(22).Caption)

    'Zeitspanne prüfen begin
    cUhrZeit = Combo11.Text
    cZeitSpanne = Combo8.Text
    'Zeitspanne ende
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(22).Caption))
    lDatBis = CLng(DateValue(Label3(21).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo7.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
    
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
    
        cSQL = "Select * from TERMINE where BEDNU = " & cbednu & " and DATUM = " & i & " "
        cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            fnPruefeVakanzMitarbeiter_abwesend = 1
        End If
        rsrs.Close: Set rsrs = Nothing
sprung:
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeVakanzMitarbeiter_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnPruefeVakanzOrtWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum          As String
    Dim lDatum          As Long
    Dim lDatVon         As Long
    Dim lDatBis         As Long
    Dim i               As Long
    Dim iAlleTage       As Integer
    Dim cUhrZeit        As String
    Dim cDauer          As String
    Dim cZeitSpanne     As String
    Dim cKabine         As String
    Dim dStartzeit      As Double
    Dim dEndeZeit       As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    fnPruefeVakanzOrtWKL82 = 0
    
    cKabine = Label2(1).Caption
    cKabine = Trim$(cKabine)
    
    cDatum = Text1(0).Text
    lDatum = DateValue(cDatum)
    
    'Zeitspanne prüfen begin
    cUhrZeit = Text1(1).Text
    If Label2(2).Caption = Label2(0).Caption Then
        cZeitSpanne = Combo5.Text
    Else
        cDauer = Text1(2).Text
        cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
        dStartzeit = TimeValue(cUhrZeit)
        dEndeZeit = dStartzeit + TimeValue(cZeitSpanne)
        cZeitSpanne = Format$(dEndeZeit, "HH:MM")
    End If
    'Zeitspanne ende
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(5).Caption))
    lDatBis = CLng(DateValue(Label3(12).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo6.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
    
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
    
        cSQL = "Select * from TERMINE where KABINE = '" & cKabine & "' and DATUM = " & i & " "
        cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            fnPruefeVakanzOrtWKL82 = 1
        End If
        rsrs.Close: Set rsrs = Nothing
sprung:
    Next i
    

    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeVakanzOrtWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function PruefeOrt(lOLDBuchnr As Long, lrow As Long, lcol As Long, cDauer As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum          As Long
    Dim cUhrZeit        As String
    Dim cZeitSpanne     As String
    Dim cKabine         As String
    Dim dStartzeit      As Double
    Dim dEndeZeit       As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    PruefeOrt = 0
    
     'Ort
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = lcol
    cKabine = Trim$(MSFlexGrid1.Text)
    
    'Datum
    lDatum = DateValue(Right(Combo2.Text, 8))
    
    'Zeitspanne prüfen begin
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    cUhrZeit = MSFlexGrid1.Text
        

    
    cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
    dStartzeit = TimeValue(cUhrZeit)
    dEndeZeit = dStartzeit + TimeValue(cZeitSpanne)
    cZeitSpanne = Format$(dEndeZeit, "HH:MM")
   
    'Zeitspanne ende
    
    cSQL = "Select * from TERMINE where KABINE = '" & cKabine & "' and DATUM = " & lDatum & " "
    cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        PruefeOrt = 1
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PruefeOrt"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function fnPruefeVakanzOrt_abwesend()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum          As String
    Dim lDatum          As Long
    Dim lDatVon         As Long
    Dim lDatBis         As Long
    Dim i               As Long
    Dim iAlleTage       As Integer
    Dim cUhrZeit        As String
    Dim cDauer          As String
    Dim cZeitSpanne     As String
    Dim cKabine         As String
    Dim dStartzeit      As Double
    Dim dEndeZeit       As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    fnPruefeVakanzOrt_abwesend = 0
    
    If Label2(11).Caption <> "" Then
        Dim sArray() As String
        sArray = Split(Label2(11).Caption, " ")
        
        For i = 3 To 3
            cKabine = UCase(Trim(sArray(i)))
        Next i
    End If
    
    lDatum = DateValue(Label3(22).Caption)
    
    'Zeitspanne prüfen begin
    cUhrZeit = Combo11.Text
    cZeitSpanne = Combo8.Text
    'Zeitspanne ende
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(22).Caption))
    lDatBis = CLng(DateValue(Label3(21).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo7.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
    
        cSQL = "Select * from TERMINE where KABINE = '" & cKabine & "' and DATUM = " & i & " "
        cSQL = cSQL & "and UHRZEIT >= '" & cUhrZeit & "' and UHRZEIT < '" & cZeitSpanne & "' "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            fnPruefeVakanzOrt_abwesend = 1
        End If
        rsrs.Close: Set rsrs = Nothing
sprung:
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeVakanzOrt_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub HoleLokalitaetenWKL82(sWelche As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lAnzSatz As Long
    Dim lcount As Long
    Dim cFeld As String
    
    cSQL = "Select * from PFLEGORT "
    If sWelche = "alle" Then
    
    Else
        cSQL = cSQL & " where anzeigen = true "
    End If
    cSQL = cSQL & " order by Anzeigen desc, BEZEICH "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzSatz = rsrs.RecordCount
        MSFlexGrid1.Cols = lAnzSatz + 2
        rsrs.MoveFirst
        MSFlexGrid1.Row = 0
        lcount = 1
        Do While Not rsrs.EOF
            lcount = lcount + 1
            MSFlexGrid1.Col = lcount
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(UCase$(cFeld))
            MSFlexGrid1.Text = cFeld
            
            rsrs.MoveNext
        Loop
    Else
        MSFlexGrid1.Cols = 1
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleLokalitaetenWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleAllePflegeOrteWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
        
    List4.Clear
    
    cSQL = "Select * from PFLEGORT order by BEZEICH "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(UCase$(cFeld))
            List4.AddItem cFeld
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleAllePflegeOrteWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboBedienerWKL82()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cZiel As String
    
    Combo1.Clear
    
    cSQL = "Select * from BEDTERM order by bednu desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = Space$(3 - Len(cFeld)) & cFeld
            
            If Not IsNull(rsrs!bedname) Then
                cFeld = rsrs!bedname
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = cZiel & " " & cFeld
            
            Combo1.AddItem cZiel
            
            If Combo1.Text = "" Then
                Combo1.Text = cZiel
                Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
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
    Fehler.gsFunktion = "fuellecboBedienerWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboDruckansicht()
    On Error GoTo LOKAL_ERROR

    Combo12.Clear
    
    Combo12.AddItem "Verfügbar"
    Combo12.AddItem "Wochenansicht"
    Combo12.AddItem "Tagesansicht"
    Combo12.AddItem "Detail Woche"
    Combo12.AddItem "Einsatzplan"
    Combo12.AddItem "Termine SMS"
    
    Combo12.Text = "Termine SMS"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecboDruckansicht"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboBediener_abwesend()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cZiel As String
    
    Combo10.Clear
    
    cSQL = "Select * from BEDTERM order by bednu desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = Space$(3 - Len(cFeld)) & cFeld
            
            If Not IsNull(rsrs!bedname) Then
                cFeld = rsrs!bedname
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = cZiel & " " & cFeld
            
            Combo10.AddItem cZiel
            
            If Combo10.Text = "" Then
                Combo10.Text = cZiel
                Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
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
    Fehler.gsFunktion = "fuellecboBediener_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboAbwesend()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    Combo3.Clear
    
    cSQL = "Select * from Abwesend order by Abwesend "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Abwesend) Then
                cFeld = rsrs!Abwesend
            End If
            cFeld = Trim(cFeld)
            
            Combo3.AddItem cFeld
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecboAbwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboAbwesend_abwesend()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    
    Combo9.Clear
    
    cSQL = "Select * from Abwesend order by Abwesend"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Abwesend) Then
                cFeld = rsrs!Abwesend
            End If
            cFeld = Trim(cFeld)
            
            Combo9.AddItem cFeld
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecboAbwesend_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub insertAbwesend(cAbwesend)
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String

    cSQL = "Delete from Abwesend where abwesend = '" & cAbwesend & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Abwesend (abwesend) values ('" & cAbwesend & "')"
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insertAbwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboDatum()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Combo2.Clear
    Combo2.Text = Left(WeekdayName(Weekday(DateValue(Now), vbMonday)), 2) & " " & Format(DateValue(Now), "DD.MM.YY")
    
    For i = -180 To 0
        Combo2.AddItem Left(WeekdayName(Weekday(DateValue(Now + i), vbMonday)), 2) & " " & Format(DateValue(Now + i), "DD.MM.YY")
    Next i
    
    For i = 1 To 180
        Combo2.AddItem Left(WeekdayName(Weekday(DateValue(Now + i), vbMonday)), 2) & " " & Format(DateValue(Now + i), "DD.MM.YY")
    Next i

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecboDatum"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleTerminBediener2WKL82()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lcount As Long
    Dim cFeld As String
    Dim cZiel As String
    
    cSQL = "Select * from BEDTERM order by BEDNAME"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        MSFlexGrid1.Cols = lcount + 2
        rsrs.MoveFirst
        lcount = 1
        MSFlexGrid1.Row = 0
        Do While Not rsrs.EOF
            lcount = lcount + 1
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = Space$(3 - Len(cFeld)) & cFeld
            
            If Not IsNull(rsrs!bedname) Then
                cFeld = rsrs!bedname
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = cZiel & " " & cFeld
            
            MSFlexGrid1.Col = lcount
            MSFlexGrid1.Text = cZiel
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleTerminBediener2WKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub entferneoverandunder()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim czeit As String
    Dim lDatum As Long

    cSQL = "Select * from TERMINE  order by UHRZEIT "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Uhrzeit) Then
                czeit = rsrs!Uhrzeit
            Else
                czeit = ""
            End If
            
            If TimeValue(czeit) >= TimeValue(gcEndeZeit) Then
                rsrs.delete
            End If
            
            If TimeValue(czeit) < TimeValue(gcStartZeit) Then
                rsrs.delete
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
    Fehler.gsFunktion = "entferneoverandunder"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseStandardTexteWKL82(sKrit As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim ctmp As String
    Dim lWert As Long
    
    List7.Clear
    
    If sKrit = "" Then
        cSQL = "Select * from TERM_STD order by NR "
    Else
        cSQL = "Select * from TERM_STD where Gliederung = '" & sKrit & "' order by NR "
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!nr) Then
                lWert = rsrs!nr
            Else
                lWert = 0
            End If
            ctmp = Format$(lWert, "#####0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(6 - Len(ctmp)) & ctmp

            cLBSatz = ctmp & " "

            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(30 - Len(ctmp))
            
            cLBSatz = cLBSatz & ctmp & Space$(10)
            
            If Not IsNull(rsrs!DAUER) Then
                lWert = rsrs!DAUER
            Else
                lWert = 0
            End If
            ctmp = Format$(lWert, "##0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(3 - Len(ctmp)) & ctmp
            
            cLBSatz = cLBSatz & ctmp
            
            List7.AddItem cLBSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseStandardTexteWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub formatgrid(msflexGridX As MSFlexGrid)
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    With msflexGridX
        .Clear
        
        .Rows = 5
        .Cols = 3
        byAnzahlSpalten = .Cols
         ReDim aBreite(.Cols)
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 500
        .Text = "Nr"
        
        .Col = 1
        .ColWidth(1) = 3000
        .Text = "Behandlung"
        
        .Col = 2
        .ColWidth(2) = 500
        .Text = "Dauer"
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "formatgrid"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseStandardTexte_inGrid(sKrit As String, msflexGridX As MSFlexGrid)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sWert As String
    Dim lWert As Long
    Dim lrow As Long
    lrow = 1
    
    formatgrid msflexGridX
    
    msflexGridX.Clear

    If sKrit = "" Then
        cSQL = "Select * from TERM_STD order by NR "
    Else
        cSQL = "Select * from TERM_STD where Gliederung = '" & sKrit & "' order by NR "
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            msflexGridX.Rows = lrow + 1
            msflexGridX.Row = lrow
            msflexGridX.Col = 0
            
            If Not IsNull(rsrs!nr) Then
                lWert = rsrs!nr
            Else
                lWert = 0
            End If
            msflexGridX.Text = Format$(lWert, "#####0")
            
            If Not IsNull(rsrs!BEZEICH) Then
                sWert = rsrs!BEZEICH
            Else
                sWert = ""
            End If
            msflexGridX.Col = 1
            msflexGridX.Text = sWert
            
            If Not IsNull(rsrs!DAUER) Then
                lWert = rsrs!DAUER
            Else
                lWert = 0
            End If
            msflexGridX.Col = 2
            msflexGridX.Text = Format$(lWert, "##0")
            
            rsrs.MoveNext
        
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    msflexGridX.RowHeight(1) = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseStandardTexte_inGrid"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseTerminBuchungWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cStartzeit As String
    Dim cEndZeit As String
    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim ctmp As String
    Dim cAllebeh As String
    Dim lPos As Long
    
    cSQL = "Select * from TERMINE where BUCHUNGSNR = " & Label4.Caption & " order by UHRZEIT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        cStartzeit = ""
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Uhrzeit) Then
                If cStartzeit = "" Then
                    cStartzeit = rsrs!Uhrzeit
                    cEndZeit = rsrs!Uhrzeit
                Else
                    cEndZeit = rsrs!Uhrzeit
                End If
            End If
            If Not IsNull(rsrs!Behandlung) Then
                Text1(3).Text = rsrs!Behandlung
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dStart = TimeValue(cStartzeit)
    dEnde = TimeValue(cEndZeit)
    dEnde = dEnde + TimeValue(gcZeitBlock)
    
    dDauer = dEnde - dStart
    
    ctmp = Format$(dDauer, "HH:MM")
    
    ctmp = fnBerechneMinuten(ctmp)
    
    Text1(1).Text = cStartzeit

    Text1(2).Text = zeigeTerminDauer(Label4.Caption)
    
    If Text1(2).Text = "" Then
        Text1(2).Text = ermBehZeit
    
        If Text1(2).Text = "0" Then
            Text1(2).Text = ctmp
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseTerminBuchungWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SchreibeTerminBuchungWKL82() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cBuchnr As String
    Dim iRet As Integer
    Dim cBed As String
    Dim cbednu As String
    Dim cbedname As String
    Dim cOrt As String
    Dim cKund As String
    Dim cKundnr As String
    Dim cKuerzel As String
    Dim cDatum As String
    Dim czeit As String
    Dim cDauer As String
    Dim cBehandlung As String
    Dim lcount As Long
    Dim lDatum As Long
    Dim ldatumDB  As Long
    Dim dStart As Double
    Dim dEnde As Double
    Dim dSprung As Double
    Dim dCounter As Double
    Dim cZeitSpanne As String
    Dim i As Long
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    Dim iAlleTage As Integer
    
    SchreibeTerminBuchungWKL82 = False
    
    If Combo3.Text <> "" Then
        insertAbwesend Left(Combo3.Text, 20)
    End If
    
    iRet = fnPruefeEingabeDialogWKL82()
    If iRet <> 0 Then
        Select Case iRet
            Case Is = 1
                MsgBox "Bitte eine/n Mitarbeiter/in bestimmen!", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 2
                MsgBox "Bitte einen Behandlungsort bestimmen!", vbInformation, "Winkiss Hinweis:"
                List4.SetFocus
            Case Is = 3
                MsgBox "Bitte einen Kunden bestimmen!", vbInformation, "Winkiss Hinweis:"
                List6.SetFocus
            Case Is = 4
                MsgBox "Bitte ein gültiges Datum angeben!", vbInformation, "Winkiss Hinweis:"
                Text1(0).SetFocus
            Case Is = 5
                MsgBox "Bitte eine gültige Uhrzeit angeben!", vbInformation, "Winkiss Hinweis:"
                Text1(1).SetFocus
            Case Is = 15
                MsgBox "Bitte eine gültige Uhrzeit angeben!", vbInformation, "Winkiss Hinweis:"
                Combo5.SetFocus
            Case Is = 13
                MsgBox "Die Öffnungszeiten sind überschritten!", vbInformation, "Winkiss Hinweis:"
                Combo5.SetFocus
            Case Is = 6
                MsgBox "Bitte eine gültige Dauer (in Minuten) angeben!", vbInformation, "Winkiss Hinweis:"
                Text1(2).SetFocus
            Case Is = 7
                MsgBox "Bitte Kunde wählen oder Abwesenheit definieren!", vbInformation, "Winkiss Hinweis:"
'                Text1(2).SetFocus
        End Select
        Exit Function
    End If
    
    If Trim$(Label2(2).Caption) = "" And Combo5.Text <> "" Then
        'trage den ausgewählten bediener ein - abwesend
        Label2(2).Caption = Label2(0).Caption
    End If
    
    cBuchnr = Label4.Caption
    cBed = Label2(0).Caption
    cbednu = ""
    cbednu = Trim(Left(Label2(0).Caption, 3))
    cbedname = Mid(Label2(0).Caption, 4, Len(Label2(0).Caption) - 3)
    cOrt = Label2(1).Caption
    
    If Label2(2).Caption = Label2(0).Caption Then
        cKundnr = ermBedKundnr(cbednu)
        cKuerzel = ermBedKuerzel(cKundnr)
    Else
        cKund = Label2(2).Caption
        cKundnr = ""
        For lcount = 1 To Len(cKund)
            If InStr("1234567890", Mid(cKund, lcount, 1)) > 0 Then
                cKundnr = cKundnr & Mid(cKund, lcount, 1)
            End If
        Next lcount
    
        lcount = InStr(1, cKund, " ")
    
        cKuerzel = Mid(cKund, lcount + 1, 5)
        cKuerzel = Trim$(UCase$(cKuerzel))
    End If
    
    gsLastKunde = cKundnr
    
    
    
    czeit = Text1(1).Text
    If Label2(2).Caption = Label2(0).Caption Then
        'rechne dauer aus
        Dim lDauer As Long
        lDauer = DateDiff("n", TimeValue(Text1(1).Text), TimeValue(Combo5.Text))
        cZeitSpanne = TimeSerial(0, lDauer - 1, 0)
        cBehandlung = Combo3.Text
    Else
        cDauer = Text1(2).Text
        cZeitSpanne = TimeSerial(0, Val(cDauer) - 1, 0)
        cBehandlung = Text1(3).Text
    End If
    dStart = TimeValue(czeit)
    dEnde = dStart + TimeValue(cZeitSpanne)

    lDatum = CLng(DateValue(Text1(0).Text))
    
    'Hier von bis übertragen
    
    lDatVon = CLng(DateValue(Label3(5).Caption))
    lDatBis = CLng(DateValue(Label3(12).Caption))
    
    If lDatum <> lDatVon Then
        lDatVon = lDatum
        lDatBis = lDatum
    End If
    
    iAlleTage = 0
    Select Case Combo6.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    Dim b14t As Boolean
    b14t = True
    
    For i = lDatVon To lDatBis
    
        
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            
                'im Prinzip kann hier alles gespeichert werden
                'aber
                If chk14t.value = vbChecked Then
                
                    If b14t = True Then
                        b14t = False
                    
                    Else
                        b14t = True
                        GoTo sprung
                    End If
                
                End If
                
            
            Else
                GoTo sprung
            End If
        End If
        
        ldatumDB = i
                            
        If Label4.Caption = "-1" Then
            cSQL = " Select max(BUCHUNGSNR) as MAXBUCHNR from TERMINE"
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                If Not IsNull(rsrs!MAXBUCHNR) Then
                    cBuchnr = rsrs!MAXBUCHNR + 1
                Else
                    cBuchnr = "1"
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            cBuchnr = Label4.Caption
        End If
        
        cSQL = "Select * from TERMINE where BUCHUNGSNR = " & Label4.Caption & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                rsrs.delete
                rsrs.MoveNext
            Loop
        End If
        
        dSprung = TimeValue(gcZeitBlock)
        
        schreibeProtokollUNITXT "Eintrag für " & cbednu & ", BuchNr: " & cBuchnr & ", KundNr: " & cKundnr & ", Termin am: " & Format(ldatumDB, "dd.mm.yyyy") & " um: " & Format$(dStart, "HH:MM") & ", Dauer: " & cDauer & ", Eintrag von: " & Text1(4).Text & ", Beh: " & cBehandlung, "Termin_Eintrag"

        
        For dCounter = dStart To dEnde Step dSprung
            czeit = Format$(dCounter, "HH:MM")
            rsrs.AddNew
            rsrs!BUCHUNGSNR = cBuchnr
            rsrs!Datum = ldatumDB
            rsrs!Uhrzeit = czeit
            rsrs!BEDNU = Val(cbednu)
            rsrs!bedname = Left(cbedname, 32)
            rsrs!Kabine = Left(cOrt, 35)
            rsrs!Kundnr = cKundnr
            rsrs!Kuerzel = Left(cKuerzel, 5)
            rsrs!Behandlung = Left(cBehandlung, 250)
            rsrs!bedeintrag = Val(Text1(4).Text)
            rsrs!neu = True
            rsrs.Update
            
            SchreibeTerminDauer cBuchnr, cDauer
            
            SchreibeTerminAnlage cBuchnr, ldatumDB, czeit, Val(Text1(4).Text)
            
            
            
            
        Next dCounter
        rsrs.Close: Set rsrs = Nothing
        
sprung:
    Next i
    
    Label2(2).Caption = ""
    Label3(5).Caption = ""
    Label3(12).Caption = ""
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = "15"
    Text1(3).Text = ""
    Combo5.Text = ""
    Combo6.Text = "alle Tage"
    Text1(0).SetFocus
    lblUnter.Visible = False
    
    SchreibeTerminBuchungWKL82 = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTerminBuchungWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SchreibeTerminBuchung_DUPLI(lOLDBuchnr As Long, lrow As Long, lcol As Long, cDauer As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cBuchnr As String
    Dim cOrt As String
    Dim czeit As String
    Dim lCounter As Long
    Dim lDatum As Long
    Dim dStart As Double
    Dim dEnde As Double
    Dim dSprung As Double
    Dim dCounter As Double
    Dim cZeitSpanne As String
    
    SchreibeTerminBuchung_DUPLI = False
    
    Dim cbednu As String
    cbednu = ermMitarbeiterausTermin(lOLDBuchnr)
    
    Dim cKundnr As String
    cKundnr = ermKundeausTermin(lOLDBuchnr)
    
    Dim cbedname As String
    cbedname = ermBednameausTermin(lOLDBuchnr)
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = lcol
    cOrt = Trim$(MSFlexGrid1.Text)
    
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    czeit = MSFlexGrid1.Text
    
    cZeitSpanne = TimeSerial(0, Val(cDauer) - 1, 0)
    
    dStart = TimeValue(czeit)
    dEnde = dStart + TimeValue(cZeitSpanne)

    lDatum = DateValue(Right(Combo2.Text, 8))
    
    cSQL = " Select max(BUCHUNGSNR) as MAXBUCHNR from TERMINE"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!MAXBUCHNR) Then
            cBuchnr = rsrs!MAXBUCHNR + 1
        Else
            cBuchnr = "1"
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Insert into Termine Select "
    cSQL = cSQL & cBuchnr & " as BUCHUNGSNR"
    cSQL = cSQL & "," & lDatum & " as Datum "
    cSQL = cSQL & ", Uhrzeit "
    cSQL = cSQL & "," & cbednu & " as Bednu "
    cSQL = cSQL & ",'" & cbedname & "' as Bedname "
    cSQL = cSQL & ",'" & cOrt & "' as Kabine "
    cSQL = cSQL & ", Kundnr "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", BEHANDLUNG "
    cSQL = cSQL & ", BEDEINTRAG "
    cSQL = cSQL & " from Termine where BUCHUNGSNR = " & lOLDBuchnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Termine_ANL Select "
    cSQL = cSQL & cBuchnr & " as BUCHUNGSNR"
    cSQL = cSQL & ", ANLAGE_DATUM "
    cSQL = cSQL & ", UHRZEIT "
    cSQL = cSQL & ", BEDEINTRAG  "
    cSQL = cSQL & ", BEDNAME  "
    cSQL = cSQL & " from Termine_ANL where BUCHUNGSNR = " & lOLDBuchnr
    gdBase.Execute cSQL, dbFailOnError
    
    SchreibeTerminDauer cBuchnr, cDauer
    
    cZeitSpanne = TimeSerial(0, Val(cDauer), 0)
    dStart = TimeValue(czeit)
    dEnde = dStart + TimeValue(cZeitSpanne)
    
    
    lCounter = 0
    cSQL = "Select * from TERMINE where BUCHUNGSNR = " & cBuchnr & " order by Uhrzeit asc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            rsrs.Edit
            
            rsrs!Uhrzeit = lCounter
            lCounter = lCounter + 1
            
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dSprung = TimeValue(gcZeitBlock)
    
    lCounter = 0
        
    For dCounter = dStart To dEnde Step dSprung
        czeit = Format$(dCounter, "HH:MM")
        
        cSQL = "Update TERMINE set Uhrzeit = '" & czeit & "' where BUCHUNGSNR = " & cBuchnr & " "
        cSQL = cSQL & " and Uhrzeit = '" & lCounter & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        lCounter = lCounter + 1

    Next dCounter
    
    
    
    SchreibeTerminBuchung_DUPLI = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTerminBuchung_DUPLI"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SchreibeTermin_abwesend() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cBuchnr As String
    Dim iRet As Integer
    Dim cBed As String
    Dim cbednu As String
    Dim cbedname As String
    Dim cOrt As String
    Dim cKund As String
    Dim cKundnr As String
    Dim cKuerzel As String
    Dim cDatum As String
    Dim czeit As String
    Dim cDauer As String
    Dim cBehandlung As String
    Dim lcount As Long
    Dim ldatumDB  As Long
    Dim dStart As Double
    Dim dEnde As Double
    Dim dSprung As Double
    Dim dCounter As Double
    Dim cZeitSpanne As String
    Dim i As Long
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    Dim iAlleTage As Integer
    
    SchreibeTermin_abwesend = False
    
    If Combo9.Text <> "" Then
        insertAbwesend Left(Combo9.Text, 20)
    End If
    
    iRet = fnPruefeEingabeDialog_abwesend()
    If iRet <> 0 Then
        Select Case iRet
            Case Is = 1
                MsgBox "Bitte eine/n Mitarbeiter/in bestimmen!", vbInformation, "Winkiss Hinweis:"
            Case Is = 4
                MsgBox "Bitte ein gültiges Datum angeben!", vbInformation, "Winkiss Hinweis:"
            Case Is = 5
                MsgBox "Bitte eine gültige Uhrzeit angeben!", vbInformation, "Winkiss Hinweis:"
            Case Is = 15
                MsgBox "Bitte eine gültige Uhrzeit angeben!", vbInformation, "Winkiss Hinweis:"
            Case Is = 13
                MsgBox "Die Öffnungszeiten sind überschritten!", vbInformation, "Winkiss Hinweis:"
        End Select
        Exit Function
    End If
    
    'check bis hier
    
    cBuchnr = "1"
    cSQL = " Select max(BUCHUNGSNR) as MAXBUCHNR from TERMINE"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!MAXBUCHNR) Then
            cBuchnr = rsrs!MAXBUCHNR + 1
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from TERMINE where BUCHUNGSNR = " & cBuchnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    cBed = Label2(11).Caption
    cbednu = ""
    cbednu = Trim(Left(Label2(11).Caption, 3))
    cbedname = Mid(Label2(11).Caption, 4, Len(Label2(11).Caption) - 3)
    
    If Label2(11).Caption <> "" Then
        Dim sArray() As String
        sArray = Split(Label2(11).Caption, " ")
        
        For i = 3 To 3
            cOrt = UCase(Trim(sArray(i)))
        Next i
    End If
    
    
    
    cKundnr = ermBedKundnr(cbednu)
    cKuerzel = ermBedKuerzel(cKundnr)
    
    
    czeit = Combo11.Text
    
    'rechne dauer aus
    Dim lDauer As Long
    lDauer = DateDiff("n", TimeValue(Combo11.Text), TimeValue(Combo8.Text))
    cZeitSpanne = TimeSerial(0, lDauer - 1, 0)
    
    cBehandlung = Combo9.Text
    
    dStart = TimeValue(czeit)
    dEnde = dStart + TimeValue(cZeitSpanne)

    'Hier von bis übertragen
    lDatVon = CLng(DateValue(Label3(22).Caption))
    lDatBis = CLng(DateValue(Label3(21).Caption))
    
    iAlleTage = 0
    Select Case Combo7.Text
        Case "alle Tage"
            iAlleTage = 0
        Case "nur montags"
            iAlleTage = 1
        Case "nur dienstags"
            iAlleTage = 2
        Case "nur mittwochs"
            iAlleTage = 3
        Case "nur donnerstags"
            iAlleTage = 4
        Case "nur freitags"
            iAlleTage = 5
        Case "nur samstags"
            iAlleTage = 6
        Case "nur sonntags"
            iAlleTage = 7
    End Select
    
    For i = lDatVon To lDatBis
        If iAlleTage > 0 Then
            If Weekday(i, vbMonday) = iAlleTage Then
            
            Else
                GoTo sprung
            End If
        End If
        
        ldatumDB = i
                            
        dSprung = TimeValue(gcZeitBlock)
        
        cSQL = "Select * from TERMINE where BUCHUNGSNR = -1 "
        Set rsrs = gdBase.OpenRecordset(cSQL)

        
        For dCounter = dStart To dEnde Step dSprung
            czeit = Format$(dCounter, "HH:MM")
            rsrs.AddNew
            rsrs!BUCHUNGSNR = cBuchnr
            rsrs!Datum = ldatumDB
            rsrs!Uhrzeit = czeit
            rsrs!BEDNU = Val(cbednu)
            rsrs!bedname = Left(cbedname, 32)
            rsrs!Kabine = Left(cOrt, 35)
            rsrs!Kundnr = cKundnr
            rsrs!Kuerzel = Left(cKuerzel, 5)
            rsrs!Behandlung = Left(cBehandlung, 250)
            rsrs!bedeintrag = Val(Text1(6).Text)
            rsrs!neu = True
            rsrs.Update
        Next dCounter
        rsrs.Close: Set rsrs = Nothing
        
sprung:
    Next i
    
    speicherSpeziInfo "Abwesenheit: " & vbCrLf & Combo9.Text, cBuchnr
    
    Label2(11).Caption = ""
    Label3(22).Caption = ""
    Label3(21).Caption = ""
    
    Combo11.Clear
    Combo8.Clear
    Combo7.Text = "alle Tage"
    
    
    SchreibeTermin_abwesend = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTermin_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Function
Private Sub SchreibeTerminDauer(cBuchnr As String, cDauer As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete * from TERMINDAUER where BUCHUNGSNR = " & cBuchnr & " "
    gdBase.Execute cSQL, dbFailOnError
   
    cSQL = "Insert into TERMINDAUER (BUCHUNGSNR,DAUER) values (" & cBuchnr & ",'" & cDauer & "') "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTerminDauer"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SchreibeTerminAnlage(cBuchnr As String, lDat As Long, sUhrzeit As String, sBednr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
  
    cSQL = "Delete * from TERMINE_ANL where BUCHUNGSNR = " & cBuchnr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    Dim sBedname As String
    sBedname = ermBEDbez(CLng(sBednr))
   
    cSQL = "Insert into TERMINE_ANL (BUCHUNGSNR,ANLAGE_DATUM,UHRZEIT,BEDEINTRAG,BEDNAME) values (" & cBuchnr & "," & lDat & ",'" & sUhrzeit & "','" & sBednr & "','" & sBedname & "') "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeTerminAnlage"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function zeigeTerminDauer(cBuchnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    zeigeTerminDauer = ""
    
    cSQL = "select dauer from TERMINDAUER where BUCHUNGSNR = " & cBuchnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!DAUER) Then
            zeigeTerminDauer = rsrs!DAUER
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigeTerminDauer"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function zeigeRasterDauer(cBuchnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim lcount As Long
    
    lcount = 0
    
    zeigeRasterDauer = ""
    
    cSQL = "select * from TERMINE where BUCHUNGSNR = " & cBuchnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lcount = lcount + 1
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
   
    zeigeRasterDauer = CStr(Minute(gcZeitBlock) * lcount)

    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigeRasterDauer"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub ZeigeTermineJeMitarbeiterWKL82(sWelcheKab As String)
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Cols = 1
    MSFlexGrid1.Rows = 1
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MA"
  
    HoleLokalitaetenWKL82 sWelcheKab
    AusrichtenTabelleWKL82 sWelcheKab
    AktualisiereTerminTabelleWKL82
    
    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeTermineJeMitarbeiterWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check2_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    zeige_Freie_Termine ermSelBed, Label2(21).Caption, Check2(0), Check2(1), Check2(2), Check2(3), Check2(4), Check2(5), Check2(6)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo1_Change()
On Error GoTo LOKAL_ERROR

    Label2(0).Caption = Combo1.Text
    Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Change"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_Click()
On Error GoTo LOKAL_ERROR

    Label2(0).Caption = Combo1.Text
    Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo10_Change()
On Error GoTo LOKAL_ERROR

    Label2(11).Caption = Combo10.Text
    Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo10_Change"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo10_Click()
On Error GoTo LOKAL_ERROR

    Label2(11).Caption = Combo10.Text
    Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo10_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Faerbebed(sbed As String, lblx As Label)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lFarbcode As Long
    Dim lFarbe As Long
    
    If sbed = "" Then
        Exit Sub
    End If
    
    If IsNumeric(sbed) = False Then
        Exit Sub
    End If

    cSQL = "Select * from BEDTERM where bednu =  " & sbed
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!FARBCODE) Then
            lFarbcode = rsrs!FARBCODE
        Else
            lFarbcode = 0
        End If
        Select Case lFarbcode
            Case Is = 0
                lFarbe = &H404040    'vbBlack
                lblx.ForeColor = vbWhite
                lblx.BackColor = lFarbe
            Case Is = 1
                lFarbe = vbRed
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
            Case Is = 2
                lFarbe = vbGreen
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
            Case Is = 3
                lFarbe = vbYellow
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
            Case Is = 4
                lFarbe = vbBlue
                lblx.ForeColor = vbWhite
                lblx.BackColor = lFarbe
                
            Case Is = 5
                lFarbe = vbMagenta
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
            Case Is = 6
                lFarbe = vbCyan
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
            Case Is = 7
                lFarbe = vbWhite
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
                
            Case Is = 8
                lFarbe = &HC0C0FF
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
            Case Is = 9
                lFarbe = &H40C0&
                lblx.ForeColor = vbWhite
                lblx.BackColor = lFarbe
                
            Case Is = 10
                lFarbe = &H80C0FF 'Apricot
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
               
                
            Case Is = 11
                lFarbe = &HFF8080 'Hellblau
'                lFarbe = &H80000003 'Hellblau
                lblx.ForeColor = vbBlack
                lblx.BackColor = lFarbe
                
        End Select
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lblx.Refresh
 
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Faerbebed"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Function FarbeBackColor(lfarbeC As Long) As Long
On Error GoTo LOKAL_ERROR

Select Case lfarbeC
    Case Is = 0
        FarbeBackColor = &H404040    'vbBlack
    Case Is = 1
        FarbeBackColor = vbRed
    Case Is = 2
        FarbeBackColor = vbGreen
    Case Is = 3
        FarbeBackColor = vbYellow
    Case Is = 4
        FarbeBackColor = vbBlue
    Case Is = 5
        FarbeBackColor = vbMagenta
    Case Is = 6
        FarbeBackColor = vbCyan
    Case Is = 7
        FarbeBackColor = vbWhite
     Case Is = 8
        FarbeBackColor = &HC0C0FF 'rosa
    Case Is = 9
        FarbeBackColor = &H40C0&     'Braun
    Case Is = 10
        FarbeBackColor = &H80C0FF        'Aprikose
    Case Is = 11
        FarbeBackColor = &HFF8080 '&H80000003        'Hellblau
End Select
 
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FarbeBackColor"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function FarbeForeColor(lfarbeC As Long) As Long
On Error GoTo LOKAL_ERROR

Select Case lfarbeC
    Case Is = 0
        FarbeForeColor = vbWhite
    Case Is = 1
        FarbeForeColor = vbBlack
    Case Is = 2
        FarbeForeColor = vbBlack
    Case Is = 3
        FarbeForeColor = vbBlack
    Case Is = 4
        FarbeForeColor = vbWhite
    Case Is = 5
        FarbeForeColor = vbBlack
    Case Is = 6
        FarbeForeColor = vbBlack
    Case Is = 7
        FarbeForeColor = vbBlack
    Case Is = 8
        FarbeForeColor = vbBlack
    Case Is = 9
        FarbeForeColor = vbWhite
    Case Is = 10
        FarbeForeColor = vbBlack
    Case Is = 11
        FarbeForeColor = vbBlack
End Select
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FarbeForeColor"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function



Private Sub Combo13_Change()
On Error GoTo LOKAL_ERROR

LeseStandardTexte_inGrid Combo13.Text, MSFlexGrid2

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo13_Change"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo13_Click()
On Error GoTo LOKAL_ERROR

LeseStandardTexte_inGrid Combo13.Text, MSFlexGrid2

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo13_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo14_Click()
On Error GoTo LOKAL_ERROR

Combo14.Refresh
       

Select Case UCase(Combo14.Text)

    
    Case "NÄCHSTE WOCHE"
    
        'nachsten Montag ermitteln
        Dim iHeute As Integer
        iHeute = Weekday(DateValue(Now), vbMonday)
        
        Dim iDiff As Integer
        
        iDiff = 7 - iHeute
        
        
        DTPickerVon.value = DateValue(Now) + iDiff + 1
        DTPickerBis.value = DateValue(Now) + iDiff + 7
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
    
    Case "14 TAGE"
        DTPickerVon.value = DateValue(Now) + 1
        DTPickerBis.value = DateValue(Now) + 14
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value

    Case "HEUTE"
        DTPickerVon.value = DateValue(Now)
        DTPickerBis.value = DateValue(Now)
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
    Case "MORGEN"
        DTPickerVon.value = DateValue(Now) + 1
        DTPickerBis.value = DateValue(Now) + 1
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
    Case "ÜBERMORGEN"
        DTPickerVon.value = DateValue(Now) + 2
        DTPickerBis.value = DateValue(Now) + 2
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
        
    Case Else
    
        Dim sDatumtext As String
        sDatumtext = Combo14.Text
        
        Dim sMonatstext As String
        sMonatstext = Trim(Left(sDatumtext, InStr(1, sDatumtext, " ")))
        
        Dim sJahr As String
        sJahr = Right(sDatumtext, 4)
        
        Dim iMonat As Integer
        
        Select Case sMonatstext
            Case Is = "Januar"
                iMonat = 1
            Case Is = "Februar"
                iMonat = 2
            Case Is = "März"
                iMonat = 3
            Case Is = "April"
                iMonat = 4
            Case Is = "Mai"
                iMonat = 5
            Case Is = "Juni"
                iMonat = 6
            Case Is = "Juli"
                iMonat = 7
            Case Is = "August"
                iMonat = 8
            Case Is = "September"
                iMonat = 9
            Case Is = "Oktober"
                iMonat = 10
            Case Is = "November"
                iMonat = 11
            Case Is = "Dezember"
                iMonat = 12
        End Select
        
        If iMonat = Month(DateValue(Now)) Then
            DTPickerVon.value = DateValue(Now)
        Else
            DTPickerVon.value = "01." & iMonat & "." & sJahr
        End If
        
        
        
        Select Case iMonat
            Case 4, 6, 9, 11
                DTPickerBis.value = "30." & iMonat & "." & sJahr
            
            Case 2
                If sJahr = "2020" Or sJahr = "2024" Or sJahr = "2028" Then
                    DTPickerBis.value = "29." & iMonat & "." & sJahr
                Else
                    DTPickerBis.value = "28." & iMonat & "." & sJahr
                End If
            
            Case 1, 3, 5, 7, 8, 10, 12
                DTPickerBis.value = "31." & iMonat & "." & sJahr
        End Select
        
        DTPickerVon.Refresh
        DTPickerBis.Refresh
        
    
        zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
           
        
        
    
'        If Combo14.Text = MonthName(Month(DateValue(Now))) & " " & Year(DateValue(Now)) Then 'aktueller Monat
'
'            DTPickerVon.Value = "01." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
'            Select Case Month(DateValue(Now))
'                Case 4, 6, 9, 11
'                    DTPickerBis.Value = "30." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
'
'                Case 2
'                    If Year(DateValue(Now)) = 2020 Or Year(DateValue(Now)) = 2024 Or Year(DateValue(Now)) = 2028 Then
'                        DTPickerBis.Value = "29." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
'                    Else
'                        DTPickerBis.Value = "28." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
'                    End If
'
'                Case 1, 3, 5, 7, 8, 10, 12
'                    DTPickerBis.Value = "31." & Month(DateValue(Now)) & "." & Year(DateValue(Now))
'            End Select
'
'
'            zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.Value, DTPickerBis.Value
'
'
'
'
'
'
'
'        End If
        
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo14_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890." & Chr$(8) & gcUPPER & gcLower
    
    cZeichen = Chr$(KeyAscii)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

If KeyCode = 189 Then
    Command3_Click 8
    Combo2.SetFocus
ElseIf KeyCode = 187 Then
    Command3_Click 9
    Combo2.SetFocus
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Combo2_Click()
On Error GoTo LOKAL_ERROR

    GesternOderMorgen DateValue(Mid(Combo2.Text, 4, 8)), Label3(6)
    AktualisiereTerminTabelleWKL82
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo3_Click()
On Error GoTo LOKAL_ERROR

If Combo3.Text <> "" Then
    Label2(2).Caption = Label2(0).Caption
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo4_Click()
On Error GoTo LOKAL_ERROR

LeseStandardTexteWKL82 Combo4.Text

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo4_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo3_GotFocus()
On Error GoTo LOKAL_ERROR

    Combo3.BackColor = glSelBack1
    Label0.Caption = 800
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo3_LostFocus()
On Error GoTo LOKAL_ERROR

    Combo3.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If index <> 40 And index <> 42 Then
        If Label0.Caption >= 0 And Label0.Caption < 999 Then
            If Label0.Caption = 1 Then
                If index = 41 Then
                    Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & ":"
                Else
                    Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & Command0(index).Caption
                End If
                
                Text1(Label0.Caption).SetFocus
                Text1(Label0.Caption).SelStart = Len(Text1(Label0.Caption).Text)
            ElseIf Label0.Caption = "800" Then
                Combo3.Text = Combo3.Text & Command0(index).Caption
                Combo3.SetFocus
            Else
                Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & Command0(index).Caption
                Text1(Label0.Caption).SetFocus
                Text1(Label0.Caption).SelStart = Len(Text1(Label0.Caption).Text)
            End If
            
        End If
    ElseIf index = 40 Then          'Löschen
        If Label0.Caption >= 0 And Label0.Caption < 999 Then
        
            If Label0.Caption = "800" Then
            
                Combo3.Text = ""
                Combo3.SetFocus
                Combo3.SelStart = Len(Combo3.Text)
            
            Else
                Text1(Label0.Caption).Text = ""
                Text1(Label0.Caption).SetFocus
                Text1(Label0.Caption).SelStart = Len(Text1(Label0.Caption).Text)
            End If
        End If
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Command1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cDateum As String
    
    Select Case index
        Case 0
            gsDatum = Right(Combo2.Text, 8)
            cDateum = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDateum), vbMonday)), 2) & " " & cDateum
            
            GesternOderMorgen DateValue(cDateum), Label3(6)
            AktualisiereTerminTabelleWKL82
        Case 1
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
        Case 2
            Label3(12).Caption = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
        Case 3
            Label3(21).Caption = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            
            If Label3(22).Caption = "" Then
                Label3(22).Caption = Label3(21).Caption
            End If
            
            Label3(23).Visible = True
            Combo8.Visible = True
            Label3(18).Visible = True
            Combo11.Visible = True
            
            fuellecombo_Uhrzeiten_von
            fuellecombo_Uhrzeiten_bis
        Case 4
            Label3(22).Caption = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            
            If Label3(21).Caption = "" Then
                Label3(21).Caption = Label3(22).Caption
            End If
            
            Label3(18).Visible = True
            Combo11.Visible = True
            Label3(23).Visible = True
            Combo8.Visible = True
            
            fuellecombo_Uhrzeiten_von
            fuellecombo_Uhrzeiten_bis
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Frame0.Visible Then
        Frame0.Visible = False
    Else
        Frame0.Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cBediener       As String
    Dim cDatum          As String
    Dim i               As Integer
    Dim dateWochedat    As Date
    Dim cKW             As String
    Dim DateHeut        As Date
    
    Select Case index
        Case Is = 0     'Schließen
            Unload frmWKL82
        Case Is = 1     'runter
            For i = 0 To Combo1.ListCount - 1
                If Trim(Combo1.Text) = Trim(Combo1.list(i)) Then
                    If i + 1 <= Combo1.ListCount - 1 Then
                        Combo1.Text = Combo1.list(i + 1)
                        Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
                        Combo1.SetFocus
                    Else
                        Combo1.Text = Combo1.list(0)
                        Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
                    End If
                    Exit For
                End If
            Next i
        Case Is = 2     'rauf
            For i = 0 To Combo1.ListCount - 1
                If Trim(Combo1.Text) = Trim(Combo1.list(i)) Then
                    If i - 1 >= 0 Then
                        Combo1.Text = Combo1.list(i - 1)
                        Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
                        Combo1.SetFocus
                    Else
                        Combo1.Text = Combo1.list(Combo1.ListCount - 1)
                        Faerbebed Trim$(Left(Combo1.Text, 3)), Label2(0)
                    End If
                    Exit For
                End If
            Next i
        Case Is = 3     'Neuer Kunde
            gcKundenNr = ""
            frmWKL13.Show 1
        Case Is = 4     'Termin kopieren, Termin einfügen
            If Command3(4).Caption = "Termin einfügen" Then
                If globBuchNr > 0 Then
                    'dann speichere den Termin an angegebener Position
                    dupliziere_Termin globBuchNr, globRow, globCol
                End If
            ElseIf Command3(4).Caption = "Termin kopieren" Then
                globBuchNr = CLng(Label8.Caption)
                
                label10_Kundendaten (globBuchNr)
                
                
                
                
                
            End If
       
        Case Is = 5    'Drucken
        
            Drucke_Plan
            
        Case Is = 6    'Speicher speziinfo
            If IsNumeric(Label8.Caption = "") = False Then
                MsgBox "Nur bei bestehenden Terminen möglich.", vbInformation, "Winkiss Hinweis:"
            Else
                If Text3.Text <> "" Then
                    speicherSpeziInfo Text3.Text, Label8.Caption
                    AktualisiereTerminTabelleWKL82
                End If
            End If

        Case Is = 9     'runter
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            cDatum = Format(DateValue(cDatum) - 1, "DD.MM.YY")
            dateWochedat = Format(DateValue(cDatum) - 1, "DD.MM.YY")
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen dateWochedat, Label3(6)
            AktualisiereTerminTabelleWKL82
        Case Is = 8     'rauf
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            cDatum = Format(DateValue(cDatum) + 1, "DD.MM.YY")
            dateWochedat = Format(DateValue(cDatum) + 1, "DD.MM.YY")
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen dateWochedat, Label3(6)
            AktualisiereTerminTabelleWKL82
            
        Case Is = 10     'aktueller Tag
            dateWochedat = Format(DateValue(Now), "DD.MM.YY")
            cDatum = Format(DateValue(Now), "DD.MM.YY")
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen dateWochedat, Label3(6)
            AktualisiereTerminTabelleWKL82
            
        Case Is = 11    'Löschen speziinfo
            loeschenSpeziInfo Label8.Caption
            Text3.Text = ""
            AktualisiereTerminTabelleWKL82
        Case Is = 12
            Frame9.Visible = False
        Case Is = 13
            Frame3.Enabled = False
            Frame10.Visible = True
            
            List1.Clear
            List5.Clear
            
            If Label3(4).Caption <> "" Then
                Label2(7).Caption = Label3(4).Caption
                
                gckundnr = Left(Label2(7).Caption, InStr(1, Label2(7).Caption, " "))
                gckundnr = Trim$(gckundnr)
                
                SucheBediener gckundnr, List1
                SucheGelöschteTermine gckundnr
                DS_Unterschrieben gckundnr
                
                gckundnr = ""
            End If
            
        Case Is = 14     'runter
            For i = 0 To Combo10.ListCount - 1
                If Trim(Combo10.Text) = Trim(Combo10.list(i)) Then
                    If i + 1 <= Combo10.ListCount - 1 Then
                        Combo10.Text = Combo10.list(i + 1)
                        Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
                        Combo10.SetFocus
                    Else
                        Combo10.Text = Combo10.list(0)
                        Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
                    End If
                    Exit For
                End If
            Next i
            
        Case Is = 15     'rauf
            For i = 0 To Combo10.ListCount - 1
                If Trim(Combo10.Text) = Trim(Combo10.list(i)) Then
                    If i - 1 >= 0 Then
                        Combo10.Text = Combo10.list(i - 1)
                        Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
                        Combo10.SetFocus
                    Else
                        Combo10.Text = Combo10.list(Combo10.ListCount - 1)
                        Faerbebed Trim$(Left(Combo10.Text, 3)), Label2(11)
                    End If
                    Exit For
                End If
            Next i
        Case Is = 16
        
            Label3(21).Caption = ""
            Label3(22).Caption = ""
        
            Label3(23).Visible = False
            Combo8.Visible = False
            Label3(18).Visible = False
            Combo11.Visible = False
        
        
            Frame11.Visible = True
            fuellecboAbwesend_abwesend
            fuellecombo7
            fuellecboBediener_abwesend
            Text1(6).Text = "" 'Bediener, der den Eintrag vornimmt

    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub label10_Kundendaten(lBuchnr As Long)
On Error GoTo LOKAL_ERROR

    Dim lkunde As Long
    lkunde = ermKundeausTermin(lBuchnr)
                
    Dim sKundenname As String
    sKundenname = ermKundenName(CStr(lkunde))
    
    Dim cTerminDauer As String
    cTerminDauer = zeigeTerminDauer(CStr(lBuchnr))
    
    If cTerminDauer = "" Then
        cTerminDauer = zeigeRasterDauer(CStr(lBuchnr))
    End If
    
    
    Label10.Caption = sKundenname & "/" & cTerminDauer
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "label10_Kundendaten"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Drucke_Plan()
On Error GoTo LOKAL_ERROR

    Dim cDatum          As String
    Dim i               As Integer
    Dim DateHeut        As Date
    Dim cKW             As String
    Dim iRet            As Integer
    
    Select Case Combo12.Text
    
        Case Is = "Verfügbar"
            'pro Mitarbeiter die Verfügbarkeit der nächsten Tage
            
            DruckeMitarbeiterVerfügbarkeit
    
        Case Is = "Wochenansicht"
        
            cDatum = Combo2.Text
            DruckeTagesPlanNeuFarbe cDatum
            
            For i = 0 To 6
                Command3_Click 8
                cDatum = Combo2.Text
                DruckeTagesPlanNeuFarbe cDatum
            Next i
            
        Case Is = "Tagesansicht"
        
            cDatum = Combo2.Text
            DruckeTagesPlanNeuFarbe cDatum
            
        Case Is = "Detail Woche"
        
            DateHeut = DateValue(Right(Combo2.Text, 8))
            cKW = DatePart("ww", DateHeut)
            DruckeWochenPlanWKL82 cKW
            
        Case Is = "Einsatzplan"
            
            DruckeMitarbeiterEinsatzPlanWKL82
            
        Case Is = "Termine SMS"
            
            DruckeTermineSMS
            
            iRet = (MsgBox("Möchten Sie jetzt auch die Termine per SMS versenden?", vbQuestion + vbYesNo, "Winkiss Frage:"))
            If iRet = vbYes Then
                'Versende Termine per SMS
                
                VersendeTermineSMS DateValue(Right(Combo2.Text, 8))
                
            End If

    End Select
    
    


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_Plan"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Terminvergabe_erlaubt(sSpeicherdatum As String) As Boolean
    
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim sErgebnis As String
    Dim iTage As Integer
    
    Terminvergabe_erlaubt = False
    
    cSQL = "Select terminanlage from OPENINGS "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        
        If Not IsNull(rsrs!terminanlage) Then
            sErgebnis = rsrs!terminanlage
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Select Case sErgebnis
        Case "1 Monat"
            iTage = 31
        Case "2 Monate"
            iTage = 61
        Case "3 Monate"
            iTage = 92
        Case "4 Monate"
            iTage = 122
        Case "5 Monate"
            iTage = 153
        Case "ohne Begrenzung"
            iTage = 0
    End Select
    
    If iTage > 0 Then
        If DateValue(sSpeicherdatum) <= DateValue(Now) + iTage Then
            Terminvergabe_erlaubt = True
        End If
    Else
        Terminvergabe_erlaubt = True
    End If
    
    If Terminvergabe_erlaubt = False Then
        MsgBox "Termine ab dem " & DateValue(Now) + iTage & " können nicht vergeben werden.", vbInformation, "Winkiss Hinweis:"
    End If
    
    
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Terminvergabe_erlaubt"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub speichern_abwesend()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet    As Integer
    Dim cSQL    As String
    Dim rsrs    As DAO.Recordset

    If Label3(22).Caption = "" Then
        MsgBox "Zeitraum angeben! (von,bis)", vbInformation, "Winkiss Hinweis:"
        
        Exit Sub
    End If
    
    
    If Terminvergabe_erlaubt(DateValue(Label3(22).Caption)) = False Then
        Exit Sub
    End If

    If Text1(6).Text = "" Then
        MsgBox "Bedienernummer eingeben!", vbInformation, "Winkiss Hinweis:"
        Text1(6).SetFocus
        Exit Sub
    End If
    
    cSQL = "Select * from BEDNAME where BEDNU = " & Text1(6).Text & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If rsrs.EOF Then
        MsgBox "Die eingegebene Bediener-Nummer ist ungültig!", vbInformation, "Winkiss Hinweis:"
        
        Text1(6).Text = ""
        Text1(6).SetFocus
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'check

    iRet = fnPruefeEingabeMitarbeiter_abwesend()
    If iRet = 0 Then
speichern:
        iRet = fnPruefeVakanzMitarbeiter_abwesend()
        If iRet = 0 Then
            iRet = fnPruefeVakanzOrt_abwesend()
            If iRet = 0 Then
                If SchreibeTermin_abwesend Then
                    Frame11.Visible = False
                End If
            Else
                MsgBox "Gewählter Termin ist nicht verfügbar! (Doppelbelegung)", vbInformation, "Winkiss Hinweis:"
            End If
        Else
            MsgBox "Mitarbeiter/in ist nicht verfügbar!", vbInformation, "Winkiss Hinweis:"
            
        End If
    Else
        Select Case iRet
            Case Is = 1         'Datum
                MsgBox "Das Datum fehlt bzw. ist ungültig!" & vbCrLf & "(Format: TT.MM.JJJJ)", vbInformation, "Winkiss Hinweis:"
            Case Is = 2         'Uhrzeit
                MsgBox "Die Uhrzeit fehlt bzw. ist ungültig!" & vbCrLf & "(Format: HH:MM auf 15 Minuten-Basis)", vbInformation, "Winkiss Hinweis:"
            Case Is = 21         'Uhrzeit drunter
                MsgBox "Die Uhrzeit liegt vor Öffnungsbeginn!", vbInformation, "Winkiss Hinweis:"
            Case Is = 22         'Uhrzeit drüber
                MsgBox "Die Uhrzeit liegt nach Ladenschluss!", vbInformation, "Winkiss Hinweis:"
            Case Is = 3         'Dauer
                MsgBox "Die Behandlungsdauer fehlt bzw. ist ungültig!" & vbCrLf & "(Format: Zeitblöcke nach eigener Definition)", vbInformation, "Winkiss Hinweis:"
            Case Is = 99         'Rückwirkend
                iRet = MsgBox("Wollen Sie wirklich einen Termin für die Vergangenheit speichern?", vbQuestion + vbYesNo, "DATUM ABGELAUFEN")
                If iRet = vbYes Then
                    GoTo speichern
                Else

                End If
        End Select
    End If





Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichern_abwesend"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Function speicher_Termin() As Boolean
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As DAO.Recordset

    speicher_Termin = False

    If Terminvergabe_erlaubt(DateValue(Text1(0).Text)) = False Then
        Exit Function
    End If

    If Text1(4).Text = "" Then
        MsgBox "Bedienernummer eingeben!", vbInformation, "Winkiss Hinweis:"
        Text1(4).SetFocus
        Exit Function
    End If
    
    cSQL = "Select * from BEDNAME where BEDNU = " & Text1(4).Text & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If rsrs.EOF Then
        MsgBox "Die eingegebene Bediener-Nummer ist ungültig!", vbInformation, "Winkiss Hinweis:"
        
        Text1(4).Text = ""
        Text1(4).SetFocus
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing

    iRet = fnPruefeEingabeMitarbeiterWKL82()
    If iRet = 0 Then
speichern:
        iRet = fnPruefeVakanzMitarbeiterWKL82()
        If iRet = 0 Then
            iRet = fnPruefeVakanzKundeWKL82()
            If iRet = 0 Then
                iRet = fnPruefeVakanzOrtWKL82()
                If iRet = 0 Then
                
                    iRet = (MsgBox("Möchten Sie den Termin auch ausdrucken?", vbQuestion + vbYesNo, "Winkiss Frage:"))
                    If iRet = vbYes Then
                        'auch mit Bon drucken
                        DruckeTerminBonWKL82
                    End If
                    
                    If SchreibeTerminBuchungWKL82 Then
                        speicher_Termin = True
                        Label3(4).Caption = ""
                        Label2(7).Caption = ""
                        Label3(4).BackColor = glH1
                        Label2(7).BackColor = glH1
                        Command4_Click 2
                    End If
                Else
                    MsgBox "Gewählter Ort ist nicht verfügbar!", vbInformation, "Winkiss Hinweis:"
                    List4.SetFocus
                End If
            Else
                MsgBox "Kunde ist nicht verfügbar! (Termin mit Zeitüberschneidung) ", vbInformation, "Winkiss Hinweis:"
            End If
        Else
            MsgBox "Mitarbeiter/in ist nicht verfügbar!", vbInformation, "Winkiss Hinweis:"
            
        End If
    Else
        Select Case iRet
            Case Is = 1         'Datum
                MsgBox "Das Datum fehlt bzw. ist ungültig!" & vbCrLf & "(Format: TT.MM.JJJJ)", vbInformation, "Winkiss Hinweis:"
                Text1(0).SetFocus
            Case Is = 2         'Uhrzeit
                MsgBox "Die Uhrzeit fehlt bzw. ist ungültig!" & vbCrLf & "(Format: HH:MM auf 15 Minuten-Basis)", vbInformation, "Winkiss Hinweis:"
                Text1(1).SetFocus
            Case Is = 21         'Uhrzeit drunter
                MsgBox "Die Uhrzeit liegt vor Öffnungsbeginn!", vbInformation, "Winkiss Hinweis:"
                Text1(1).SetFocus
            Case Is = 22         'Uhrzeit drüber
                MsgBox "Die Uhrzeit liegt nach Ladenschluss!", vbInformation, "Winkiss Hinweis:"
                Text1(1).SetFocus
            Case Is = 44         'Kunde fehlt
                MsgBox "Kunde fehlt!", vbInformation, "Winkiss Hinweis:"
'                Text1(1).SetFocus
            Case Is = 3         'Dauer
                MsgBox "Die Behandlungsdauer fehlt bzw. ist ungültig!" & vbCrLf & "(Format: Zeitblöcke nach eigener Definition)", vbInformation, "Winkiss Hinweis:"
                Text1(2).SetFocus
            Case Is = 99         'Rückwirkend
                iRet = MsgBox("Wollen Sie wirklich einen Termin für die Vergangenheit speichern?", vbQuestion + vbYesNo, "DATUM ABGELAUFEN")
                If iRet = vbYes Then
                    GoTo speichern
                Else
                    Text1(2).SetFocus
                End If
        End Select
    End If



Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_Termin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function speicher_Dupli_Termin(lOLDBuchnr As Long, lrow As Long, lcol As Long, cDauer As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As DAO.Recordset

    speicher_Dupli_Termin = False

    If Terminvergabe_erlaubt(DateValue(Text1(0).Text)) = False Then
        Exit Function
    End If



    iRet = PruefeDatumZeit(lrow)
    If iRet = 0 Then
speichern:
        iRet = PruefeMitarbeiter(lOLDBuchnr, lrow, lcol, cDauer)
        If iRet = 0 Then
            iRet = PruefeOrt(lOLDBuchnr, lrow, lcol, cDauer)
            If iRet = 0 Then
            
                If SchreibeTerminBuchung_DUPLI(lOLDBuchnr, lrow, lcol, cDauer) Then
                
                    Command3(4).Caption = "Termin kopieren"
                    Command3(4).ForeColor = vbBlack
                    
                    'jetzt noch Tagesansicht aktualisieren
                    Command4_Click 2
                    
                End If
            Else
                MsgBox "Gewählter Ort ist nicht verfügbar!", vbInformation, "Winkiss Hinweis:"
            End If
        Else
            MsgBox "Mitarbeiter/in ist nicht verfügbar!", vbInformation, "Winkiss Hinweis:"
        End If
    Else
        Select Case iRet
            Case Is = 1         'Datum
                MsgBox "Das Datum fehlt bzw. ist ungültig!" & vbCrLf & "(Format: TT.MM.JJJJ)", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 2         'Uhrzeit
                MsgBox "Die Uhrzeit fehlt bzw. ist ungültig!" & vbCrLf & "(Format: HH:MM auf 15 Minuten-Basis)", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 21         'Uhrzeit drunter
                MsgBox "Die Uhrzeit liegt vor Öffnungsbeginn!", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 22         'Uhrzeit drüber
                MsgBox "Die Uhrzeit liegt nach Ladenschluss!", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 3         'Dauer
                MsgBox "Die Behandlungsdauer fehlt bzw. ist ungültig!" & vbCrLf & "(Format: Zeitblöcke nach eigener Definition)", vbInformation, "Winkiss Hinweis:"
                
            Case Is = 99         'Rückwirkend
                iRet = MsgBox("Wollen Sie wirklich einen Termin für die Vergangenheit speichern?", vbQuestion + vbYesNo, "DATUM ABGELAUFEN")
                If iRet = vbYes Then
                    GoTo speichern
                Else
                    
                End If
        End Select
    End If



Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_Dupli_Termin"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Sub ZeigeTermine(cKundnr As String, Listx)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sSatz As String
    Dim sFeld As String
    
    Listx.Clear
    sSQL = "select  "
    sSQL = sSQL & " DATUM "
    sSQL = sSQL & ", BEHANDLUNG"
    sSQL = sSQL & ", BEDNAME"
    
    sSQL = sSQL & " from Termine where KUNDNR = " & cKundnr
    sSQL = sSQL & " group by buchungsnr, Datum ,BEDNAME ,BEHANDLUNG "
    sSQL = sSQL & " order by Datum desc "
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        sFeld = ""
        
        If Not IsNull(rsrs!Datum) Then
            sFeld = rsrs!Datum
        End If
        
        sSatz = Format(sFeld, "DD.MM.YY")
        sSatz = sSatz & " "
        
        If Not IsNull(rsrs!Behandlung) Then
            sFeld = rsrs!Behandlung
        End If
        
        sFeld = SwapStr(sFeld, Chr(13), " ")
        sFeld = SwapStr(sFeld, Chr(10), " ")
        sFeld = SwapStr(sFeld, "  ", " ")
        
        If Len(sFeld) > 30 Then
            sSatz = sSatz & Left(sFeld, 27) & "... "
        Else
            sSatz = sSatz & sFeld & Space(31 - Len(sFeld))
        End If
        
        If Not IsNull(rsrs!bedname) Then
            sFeld = rsrs!bedname
        End If
        
        If Len(sFeld) > 20 Then
            sSatz = sSatz & Left(sFeld, 17) & "... "
        Else
            sSatz = sSatz & sFeld & Space(21 - Len(sFeld))
        End If
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    Else

'        Label1(4).Visible = False
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeTermine"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub ZeigeTermine_zumLoeschen(cKundnr As String, Listx, Optional lAbDatum As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sSatz As String
    Dim sFeld As String
    
    Listx.Clear
    sSQL = "select  "
    sSQL = sSQL & " DATUM "
    sSQL = sSQL & ", BEHANDLUNG"
    sSQL = sSQL & ", BEDNAME"
    sSQL = sSQL & ", min(Uhrzeit) As mini "
    sSQL = sSQL & ", buchungsnr "
    
    sSQL = sSQL & " from Termine where KUNDNR = " & cKundnr
    
    If lAbDatum > 0 Then
        sSQL = sSQL & " and datum = " & lAbDatum
    End If
    
    
    sSQL = sSQL & " group by buchungsnr, Datum ,BEDNAME ,BEHANDLUNG "
    sSQL = sSQL & " order by Datum asc, min(Uhrzeit) asc  "
    
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        sFeld = ""
        
        If Not IsNull(rsrs!Datum) Then
            sFeld = rsrs!Datum
        End If
        
        sSatz = Format(sFeld, "DD.MM.YY")
        sSatz = sSatz & " "
        
        If Not IsNull(rsrs!mini) Then
            sFeld = rsrs!mini
        End If
        
        sSatz = sSatz & Format(sFeld, "HH:MM:SS")
        sSatz = sSatz & " "
        
        If Not IsNull(rsrs!Behandlung) Then
            sFeld = rsrs!Behandlung
        End If
        
        sFeld = SwapStr(sFeld, Chr(13), " ")
        sFeld = SwapStr(sFeld, Chr(10), " ")
        sFeld = SwapStr(sFeld, "  ", " ")
        
        If Len(sFeld) > 30 Then
            sSatz = sSatz & Left(sFeld, 27) & "... "
        Else
            sSatz = sSatz & sFeld & Space(31 - Len(sFeld))
        End If
        
        If Not IsNull(rsrs!bedname) Then
            sFeld = rsrs!bedname
        End If
        
        If Len(sFeld) > 20 Then
            sSatz = sSatz & Left(sFeld, 17) & "... "
        Else
            sSatz = sSatz & sFeld & Space(21 - Len(sFeld))
        End If
        
        
        
        
        If Not IsNull(rsrs!BUCHUNGSNR) Then
            sFeld = rsrs!BUCHUNGSNR
        End If
        
        sSatz = sSatz & Space(100) & sFeld
        
        
        
        
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    Else

'        Label1(4).Visible = False
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeTermine_zumLoeschen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Behandlungen_anzeigen(bSichtbar As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Combo13.Visible = bSichtbar
    Label2(17).Visible = bSichtbar
    MSFlexGrid2.Visible = bSichtbar
    
    List10.Visible = bSichtbar
    Label2(19).Visible = bSichtbar
    Label2(18).Visible = bSichtbar
    Text1(5).Visible = bSichtbar
    
    
    If bSichtbar = True Then
        fuellecombo Combo13
    End If
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Behandlungen_anzeigen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim iRet            As Integer
    Dim ctmp            As String
    Dim lcount          As Long
    Dim cZeichen        As String
    Dim cKdnr           As String
    Dim lDatum          As Long
    Dim rsrs            As Recordset
    Dim cSQL            As String
    Dim dateWochedat    As Date
    Dim sDelBed         As String
    
    Select Case index
        Case Is = 0     'Speichern
            If speicher_Termin Then
            
            End If
        Case Is = 1     'Löschen
            iRet = MsgBox("Den bestehenden Termin komplett löschen?", vbYesNo + vbQuestion, "Winkiss Frage:")
            If iRet = vbYes Then
            
            
                Dim cTermDelgrund As String
                dlgTermDel.Show 1
                
                Select Case dlgTermDel.Back
                    Case 1
                        cTermDelgrund = "rechtzeitig abgesagt"
                    Case 2
                        cTermDelgrund = "rechtzeitig verlegt"
                    Case 3
                        cTermDelgrund = "kurzfristig verlegt"
                    Case 4
                        cTermDelgrund = "kurzfristig abgesagt"
                    Case 5
                        cTermDelgrund = "kurzfristig krank"
                    Case 6
                        cTermDelgrund = "nicht erschienen"
                    Case 7
                        cTermDelgrund = "falsch eingetragen"
                        
                        
                    Case 0 'Abbrechen
                        Exit Sub
                        
                End Select
            
                
                
                
                Do
                    sDelBed = InputBox("Geben Sie bitte Ihre Bedienernummer ein!", "Wer löscht den Termin?")
                    sDelBed = SwapStr(sDelBed, ",", "")
                Loop While IsNumeric(sDelBed) = False
                
                
'                sDelBed = InputBox("Geben Sie bitte Ihre Bedienernummer ein!", "Wer löscht den Termin?")
            
                LoescheTerminKomplettWKL82 cTermDelgrund, sDelBed
                
                Label4.Caption = -1
                MsgBox "Der Termin wurde gelöscht", vbOK + vbInformation, "Winkiss Hinweis:"
                
                'das gehört zusammen
                Dim cKundnr As String
                cKundnr = Trim(Mid(Label2(2).Caption, 1, InStr(1, Label2(2).Caption, " ") - 1))
                
                
                lblBed.Caption = sDelBed
                lblDelgrund.Caption = cTermDelgrund
                lblKunde.Caption = cKundnr
                lblDatum.Caption = CLng(DateValue(Text1(0).Text))
                
                If Gibt_es_Termine_in_Zunkunft(cKundnr, CLng(DateValue(Text1(0).Text))) > 0 Then
                
                    ZeigeTermine_zumLoeschen cKundnr, List9, CLng(DateValue(Text1(0).Text))
                    
                    Frame14.Visible = True
                    
                    'ENDE das gehört zusammen
                Else
                    Command4_Click 2
                End If
            End If
        Case Is = 2     'Beenden
        
            Screen.MousePointer = 11
            ZeigeTermineJeMitarbeiterWKL82 ""
            
            If Label2(10).Caption = "alle anzeigen" Then
                ZeigeTermineJeMitarbeiterWKL82 ""
            Else
                ZeigeTermineJeMitarbeiterWKL82 "alle"
            End If
            
            Frame4.Visible = False
            Frame3.Visible = True
            Command3(0).Visible = True
            Command3(3).Visible = True
            Command3(5).Visible = True
            Combo12.Visible = True
            Label2(2).Caption = ""
            lblUnter.Visible = False
            
            Dim dateVerf As Date
            dateVerf = DateValue(Mid(Combo2.Text, 4, 8))
            verfuegbar dateVerf 'dateDat
            Frame9.Visible = True
            
            Screen.MousePointer = 0
            
        Case 3         'Kundeninfo
            ZeigeKundenInfo Label2(2)
        Case 4          'Kundenhistorie
            ZeigeKundenHistorie Label2(2)
        Case 5          'den Termin auf Bon drucken
            DruckeTerminBonWKL82
            
        Case 6
            Text1(3).Text = ""
            Text1(2).Text = ""
        Case 7
            Screen.MousePointer = 11
            List2.Clear
            List6.Clear
            SucheKunde Label2(2), List6, List2
            
            If Label2(2).Caption <> "" Then
                Label3(4).Caption = Label2(2).Caption
                Label3(4).BackColor = Label2(2).BackColor
                
                If Gibt_es_Termine_in_Zunkunft(Left(Label3(4).Caption, InStr(1, Label3(4).Caption, " "))) > 0 Then
                    Label3(4).ToolTipText = "weitere Termine (Doppelklick)"
                    Label3(4).FontUnderline = True
                    Label3(4).ForeColor = glLink
                Else
                    Label3(4).ToolTipText = ""
                    Label3(4).FontUnderline = False
                    Label3(4).ForeColor = glS1
                End If
            End If
        Case 8
            Frame10.Visible = False
            Frame3.Enabled = True
            Label3(4).Caption = Label2(7).Caption
            Label3(4).BackColor = Label2(7).BackColor
            
            
            If Gibt_es_Termine_in_Zunkunft(Left(Label3(4).Caption, InStr(1, Label3(4).Caption, " "))) > 0 Then
                Label3(4).ToolTipText = "weitere Termine (Doppelklick)"
                Label3(4).FontUnderline = True
                Label3(4).ForeColor = glLink
                
              
            Else
                Label3(4).ToolTipText = ""
                Label3(4).FontUnderline = False
                Label3(4).ForeColor = glS1
            End If
        Case 9              'Kundendatenblatt
            ZeigeDatenblatt Label2(2)
        Case 10             'Kundendatenblatt
            ZeigeDatenblatt Label2(7)
        Case 11        'Kundeninfo
            ZeigeKundenInfo Label2(7)
        Case 12         'Kundenhistorie
            ZeigeKundenHistorie Label2(7)
        Case 13
            Screen.MousePointer = 11
            List1.Clear
            List5.Clear
            SucheKunde Label2(7), List1, List5
            If Label2(7).Caption <> "" Then
                Command4(8).Caption = "wählen"
                Behandlungen_anzeigen True
            Else
                Command4(8).Caption = "schließen"
                
                Behandlungen_anzeigen False
            End If
        Case 14
        
            speichern_abwesend
            
            'aktualisiere die nachfolgende Anzeige
            If Label2(10).Caption = "alle anzeigen" Then
                ZeigeTermineJeMitarbeiterWKL82 ""
            Else
                ZeigeTermineJeMitarbeiterWKL82 "alle"
            End If
            
            dateVerf = DateValue(Mid(Combo2.Text, 4, 8))
            verfuegbar dateVerf 'dateDat
            
        Case 15
            If Command4(15).Caption = "Achtung" Then
                Command4(15).Caption = "schließen"
                List2.Visible = True
            Else
                Command4(15).Caption = "Achtung"
                List2.Visible = False
            End If
        Case 16
            If Command4(16).Caption = "Achtung" Then
                Command4(16).Caption = "schließen"
                List5.Visible = True
            Else
                Command4(16).Caption = "Achtung"
                List5.Visible = False
            End If
        Case 17 'verschieben = löschen und speichern
        
        
            Do
                sDelBed = InputBox("Geben Sie bitte Ihre Bedienernummer ein!", "Wer verändert den Termin?")
                sDelBed = SwapStr(sDelBed, ",", "")
            Loop While IsNumeric(sDelBed) = False

            
            
            
            
''            do while IsNumeric(sDelBed) = False
'
'            loop while
'            If IsNumeric(sDelBed) = False Then
'                MsgBox "", vbInformation, "Winkiss Hinweis:"
'                Exit Sub
'            End If
            
            Dim sKundeDel As String
            sKundeDel = "0"
        
            sKundeDel = Val(Left(Label2(2).Caption, InStr(1, Label2(2).Caption, " ")))
            sKundeDel = Trim$(sKundeDel)
            Dim cDauer As String
            cDauer = "0"
            cDauer = Trim(Text1(2).Text)
        
            LoescheTermin_vorVerschieben
            
            Text1(4).Text = sDelBed
            
            If speicher_Termin = False Then
                'rückgängig
                Termin_Zurückholen
            Else
                löschInfo_schreiben "Termin verschoben", sDelBed, sKundeDel, cDauer
            End If
        Case 18
            Frame11.Visible = False
        Case 19
            Frame12.Visible = False
        Case 21 'Notizen
        
            ZeigeKundenNotizen Label2(2)
            
        Case 22              'Datenschutzblatt
            DatenSchutzblatt_Drucken Label2(2)
            
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKundenNotizen(labelx As Label)
On Error GoTo LOKAL_ERROR

    Text2.Text = ""
    Label3(26).Caption = ""

    If labelx.Caption = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    Dim rsrs As DAO.Recordset
    Dim sSQL As String
    Dim sKund As String
    
    sKund = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    
    If IsNumeric(sKund) Then
        Label3(26).Caption = sKund
        sSQL = "Select Notizen from Kunden where kundnr = " & sKund
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!NOTIZEN) Then
                Text2.Text = rsrs!NOTIZEN
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        Frame13.Visible = True
            
        Text2.SetFocus
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseKundenNotizen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub SpeicherKundenNotizen(labelx As Label)
On Error GoTo LOKAL_ERROR

    If labelx.Caption = "" Then
        Exit Sub
    End If
    
    Dim sSQL As String
    Dim sKund As String
    
    sKund = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    
    If IsNumeric(sKund) Then
    
        sSQL = "Update kunden set notizen = '" & Text2.Text & "'"
        sSQL = sSQL & " , Status = 'E' "
        sSQL = sSQL & " , SynStatus = 'E' "
        sSQL = sSQL & " where kundnr = " & sKund
        gdBase.Execute sSQL, dbFailOnError
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherKundenNotizen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ZeigeKundenInfo(labelx As Label)
On Error GoTo LOKAL_ERROR

    If labelx.Caption = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    

    gckundnr = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    gsARTNR = ""
    
    If IsNumeric(gckundnr) Then
        frmWKL94.Show 1
        gckundnr = ""
    Else
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKundenInfo"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ZeigeKundenHistorie(labelx As Label)
On Error GoTo LOKAL_ERROR

    If labelx.Caption = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    

    gckundnr = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    gsARTNR = ""
    
    If IsNumeric(gckundnr) Then
        frmWKL74.Show 1
        gckundnr = ""
    Else
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKundenHistorie"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub ZeigeDatenblatt(labelx As Label)
    On Error GoTo LOKAL_ERROR

    If labelx.Caption = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    gckundnr = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    gsARTNR = ""
    
    If IsNumeric(gckundnr) Then
        gcKundenNr = gckundnr
        iKasse = 3
    
        frmWKL13.Show 1
        gckundnr = ""
    Else
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeDatenblatt"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DatenSchutzblatt_Drucken(labelx As Label)
    On Error GoTo LOKAL_ERROR

    If labelx.Caption = "" Then
        MsgBox "Kein Kunde gewählt!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    Dim cKundnr As String
    cKundnr = Trim(Mid(labelx.Caption, 1, InStr(1, labelx.Caption, " ") - 1))
    
    DatenschutzblattKundeDrucken cKundnr
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DatenSchutzblatt_Drucken"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheKunde(labelx As Label, Listx As ListBox, Listy As ListBox)
On Error GoTo LOKAL_ERROR

    frmWKL134.Show 1

    labelx.Caption = ""
    labelx.BackColor = glH1
    labelx.Refresh

    Dim cKürzel As String
    Dim cName As String
    Dim cVorname As String
    Dim SFarbe As String

    If gckundnr <> "" Then
        If IsNumeric(gckundnr) Then
        
            SFarbe = ermFarbe(Trim(gckundnr))
            If Trim(SFarbe) = "0" Then
                labelx.BackColor = glH1
            Else
                labelx.BackColor = glfarbe(SFarbe)
            End If
        
        
            
        
            cKürzel = lookingForKundendaten(gckundnr).Kuerzel
            cName = lookingForKundendaten(gckundnr).nachname
            cVorname = lookingForKundendaten(gckundnr).vorname
        
            labelx.Caption = gckundnr & "  "
            labelx.Caption = labelx.Caption & Space$(5 - Len(cKürzel)) & cKürzel & "  "
            labelx.Caption = labelx.Caption & cName & ", "
            labelx.Caption = labelx.Caption & cVorname
            
            lblUnter.Visible = False
                
            If SucheUnter(gckundnr) Then
                lblUnter.ForeColor = glWarn
                lblUnter.Visible = True
            End If
            
            SucheBediener gckundnr, Listx
            SucheGelöschteTermine gckundnr
            DS_Unterschrieben gckundnr
            
            

        End If
    End If
    gckundnr = ""
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheKunde"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheTerminKomplettWKL82(sDel As String, sbed As String, Optional sBuchnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sKundeDel As String
    Dim sErsteller As String
    Dim cBuchungsNr As String
    Dim cDatum As String
    Dim czeit As String
    Dim cDauer As String
    Dim cErstDatum As String
    Dim cErstZeit As String
    
    sErsteller = "0"
    
    
    
    If sBuchnr <> "" Then
        cBuchungsNr = sBuchnr
    Else
        cBuchungsNr = Label4.Caption
    End If
    If cBuchungsNr <> "-1" Then
    
        cSQL = "Select min(uhrzeit) as mini, datum, bednu from Termine where BUCHUNGSNR = " & cBuchungsNr & " "
        cSQL = cSQL & " group by datum, bednu "
    
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!BEDNU) Then
                sErsteller = Val(rsrs!BEDNU)
            End If
            
            If Not IsNull(rsrs!Datum) Then
                cErstDatum = rsrs!Datum
            End If
            
            If Not IsNull(rsrs!mini) Then
                cErstZeit = rsrs!mini
            End If
        
            
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    If cBuchungsNr <> "-1" Then
        loeschNEW "Terme_DEl_" & srechnertab, gdBase
        cSQL = "Select * into Terme_DEl_" & srechnertab & " from TERMINE where BUCHUNGSNR = " & cBuchungsNr & " "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Delete from TERMINE where BUCHUNGSNR = " & cBuchungsNr & " "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from TERMINE_ANL where BUCHUNGSNR = " & cBuchungsNr & " "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If cBuchungsNr <> "-1" Then
    
        cDatum = Fix(Now)
        czeit = Format$(Now, "HH:MM:SS")
        
        sKundeDel = "0"
        
        sKundeDel = Val(Left(Label2(2).Caption, InStr(1, Label2(2).Caption, " ")))
        sKundeDel = Trim$(sKundeDel)
        
        cDauer = "0"
        cDauer = Trim(Text1(2).Text)
        If cDauer = "" Then cDauer = "0"
        
        cSQL = "Insert into TERMDEL (KUNDNR,GRUND,BED,ADATE,AZEIT,ERSTBED,DAUER,BEGINDAT,BEGINZEIT)"
        cSQL = cSQL & " values ("
        cSQL = cSQL & " " & sKundeDel & " "
        cSQL = cSQL & ", '" & sDel & "' "
        cSQL = cSQL & ", " & sbed & " "
        cSQL = cSQL & ", '" & cDatum & "'"
        cSQL = cSQL & ", '" & czeit & "'"
        cSQL = cSQL & ", " & sErsteller & " "
        cSQL = cSQL & ", " & cDauer & " "
        cSQL = cSQL & ", " & cErstDatum & " "
        cSQL = cSQL & ", '" & cErstZeit & "'"
        cSQL = cSQL & " ) "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheTerminKomplettWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten. " & cSQL
    
    Fehlermeldung1
End Sub
Private Sub LoescheTermin_vorVerschieben()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim cBuchungsNr     As String
    
    cBuchungsNr = Label4.Caption
    
    If cBuchungsNr <> "-1" Then
        loeschNEW "Terme_DEl_" & srechnertab, gdBase
        cSQL = "Select * into Terme_DEl_" & srechnertab & " from TERMINE where BUCHUNGSNR = " & cBuchungsNr & " "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "Delete from TERMINE where BUCHUNGSNR = " & cBuchungsNr & " "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheTermin_vorVerschieben"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub löschInfo_schreiben(sDel As String, sbed As String, sKundeDel As String, cDauer As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sErsteller As String
    Dim cDatum As String
    Dim czeit As String
    Dim cErstDatum As String
    Dim cErstZeit As String
    
    sErsteller = "0"
    
    If NewTableSuchenDBKombi("Terme_DEl_" & srechnertab, gdBase) = True Then
    
        cSQL = "Select min(uhrzeit) as mini, datum, bednu from Terme_DEl_" & srechnertab
        cSQL = cSQL & " group by datum, bednu "
    
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!BEDNU) Then
                sErsteller = Val(rsrs!BEDNU)
            End If
            
            If Not IsNull(rsrs!Datum) Then
                cErstDatum = rsrs!Datum
            End If
            
            If Not IsNull(rsrs!mini) Then
                cErstZeit = rsrs!mini
            End If
        
            
        End If
        rsrs.Close: Set rsrs = Nothing
    
    
        cDatum = Fix(Now)
        czeit = Format$(Now, "HH:MM:SS")
        
        
        
        cSQL = "Insert into TERMDEL (KUNDNR,GRUND,BED,ADATE,AZEIT,ERSTBED,DAUER,BEGINDAT,BEGINZEIT)"
        cSQL = cSQL & " values ("
        cSQL = cSQL & " " & sKundeDel & " "
        cSQL = cSQL & ", '" & sDel & "' "
        cSQL = cSQL & ", " & sbed & " "
        cSQL = cSQL & ", '" & cDatum & "'"
        cSQL = cSQL & ", '" & czeit & "'"
        cSQL = cSQL & ", " & sErsteller & " "
        cSQL = cSQL & ", " & cDauer & " "
        cSQL = cSQL & ", " & cErstDatum & " "
        cSQL = cSQL & ", '" & cErstZeit & "'"
        cSQL = cSQL & " ) "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "löschInfo_schreiben"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Termin_Zurückholen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    If NewTableSuchenDBKombi("Terme_DEl_" & srechnertab, gdBase) = True Then
        
        cSQL = "Insert into TERMINE Select * from Terme_DEl_" & srechnertab & " "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "Terme_DEl_" & srechnertab, gdBase
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Termin_Zurückholen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cBezeich        As String
    Dim cDauer          As String
    Dim dateWochedat    As Date
    Dim cDatum          As String
    Dim cMonTeil        As String
    Dim cTagTeil        As String
    Dim cJahrTeil       As String
    
    Select Case index
        Case Is = 0
            If List7.ListIndex < 0 Then
                MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbInformation, "Winkiss Hinweis:"
                List7.SetFocus
            Else
                cLBSatz = List7.list(List7.ListIndex)
                cBezeich = Mid(cLBSatz, 8, 30)
                cDauer = Right(cLBSatz, 10)
                cBezeich = Trim$(cBezeich)
                cDauer = Trim$(cDauer)
                If Trim$(Text1(3).Text) = "" Then
                    Text1(3).Text = cBezeich
                    Text1(2).Text = cDauer
                    Text1(3).Text = Text1(3).Text & Chr$(13) & Chr$(10)
                    
                Else
                    Text1(3).Text = Text1(3).Text & cBezeich
                    Text1(2).Text = Trim$(Str$(Val(Text1(2).Text) + Val(cDauer)))
                    Text1(3).Text = Text1(3).Text & Chr$(13) & Chr$(10)
                End If
                Text1(2).SetFocus
            End If
        Case 1 'Zurück aus Notizen
            SpeicherKundenNotizen Label2(2)
            Frame13.Visible = False
            
        Case 2
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            
            dateWochedat = Format(DateValue(cDatum) + 7, "DD.MM.YY")
            cDatum = Format(DateValue(cDatum) + 7, "DD.MM.YY")
            
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen dateWochedat, Label3(6)
            AktualisiereTerminTabelleWKL82
            
        Case 3
            If Command6(3).Caption = "+" Then
                Command6(3).Left = dWidth - Command6(3).Width - 300
                Command6(3).Caption = "-"
                Frame3.Width = dWidth - 300
                MSFlexGrid1.Width = dWidth - 300
                Frame9.Visible = False
            ElseIf Command6(3).Caption = "-" Then
                Command6(3).Left = (dWidth / 5 * 4) - Command6(3).Width
                Command6(3).Caption = "+"
                Frame3.Width = dWidth / 5 * 4
                MSFlexGrid1.Width = dWidth / 5 * 4
                Frame9.Visible = True
            End If
        Case 4
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            dateWochedat = Format(DateValue(cDatum) - 7, "DD.MM.YY")

            cDatum = Format(DateValue(cDatum) - 7, "DD.MM.YY")
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen dateWochedat, Label3(6)
            AktualisiereTerminTabelleWKL82
            
        Case 5
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            
            cMonTeil = Mid(cDatum, 4, 2)
            cTagTeil = Left(cDatum, 2)
            cJahrTeil = Right(cDatum, 2)
            
            If Val(cMonTeil) = 1 Then
                cMonTeil = "12"
                cJahrTeil = Val(cJahrTeil) - 1
            Else
                cMonTeil = Val(cMonTeil) - 1
            End If
            
            If Val(cTagTeil) = "31" Then
                Select Case Val(cMonTeil)
                    Case 4, 6, 9, 11
                        cTagTeil = "30"
                    Case 2
                        cTagTeil = "28"
                End Select
            End If
            
            If Val(cTagTeil) = "31" And Val(cMonTeil) = "2" Then
                cTagTeil = "28"
            End If
            
            If Len(cTagTeil) = 1 Then
                cTagTeil = "0" & cTagTeil
            End If
            
            If Len(cJahrTeil) = 1 Then
                cJahrTeil = "0" & cJahrTeil
            End If
            
            If Len(cMonTeil) = 1 Then
                cMonTeil = "0" & cMonTeil
            End If
            
            cDatum = cTagTeil & "." & cMonTeil & "." & cJahrTeil
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen DateValue(cDatum), Label3(6)
            AktualisiereTerminTabelleWKL82
            
        Case 6
            cDatum = Combo2.Text
            cDatum = Right(cDatum, 8)
            
            cMonTeil = Mid(cDatum, 4, 2)
            cTagTeil = Left(cDatum, 2)
            cJahrTeil = Right(cDatum, 2)
            
            If Val(cMonTeil) = 12 Then
                cMonTeil = "1"
                cJahrTeil = Val(cJahrTeil) + 1
            Else
                cMonTeil = Val(cMonTeil) + 1
            End If
            
            If Val(cTagTeil) = "31" Then
                Select Case Val(cMonTeil)
                    Case 4, 6, 9, 11
                        cTagTeil = "30"
                    Case 2
                        cTagTeil = "28"
                End Select
            End If
            
            If Val(cTagTeil) = "31" And Val(cMonTeil) = "2" Then
                
                cTagTeil = "28"
                    
            End If
            
            If Len(cTagTeil) = 1 Then
                cTagTeil = "0" & cTagTeil
            End If
            
            If Len(cJahrTeil) = 1 Then
                cJahrTeil = "0" & cJahrTeil
            End If
            
            If Len(cMonTeil) = 1 Then
                cMonTeil = "0" & cMonTeil
            End If
            
            cDatum = cTagTeil & "." & cMonTeil & "." & cJahrTeil
            Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDatum), vbMonday)), 2) & " " & cDatum
            
            GesternOderMorgen DateValue(cDatum), Label3(6)
            AktualisiereTerminTabelleWKL82
        Case 7
            Frame14.Visible = False
            Command4_Click 2
        Case 8
            If List9.ListIndex < 0 Then
                MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbInformation, "Winkiss Hinweis:"
                List9.SetFocus
            Else
                Dim sBuchnr As String
                
                cLBSatz = List9.list(List9.ListIndex)
                sBuchnr = Trim(Right(cLBSatz, 10))
                
                LoescheTerminKomplettWKL82 lblDelgrund.Caption, lblBed.Caption, sBuchnr
                
                ZeigeTermine_zumLoeschen lblKunde.Caption, List9, CLng(lblDatum.Caption)
                
            End If
        Case 10 'Drucken Kundenstammdatenblatt
            StammdatenblattKundeDrucken Label3(26).Caption, True, ""
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GesternOderMorgen(dateDat As Date, lblx As Label)
On Error GoTo LOKAL_ERROR

    If dateDat = DateValue(Now) Then
        lblx.Caption = "Heute"
    ElseIf dateDat = DateValue(Now) + 1 Then
        lblx.Caption = "Morgen"
    ElseIf dateDat = DateValue(Now) + 2 Then
        lblx.Caption = "Übermorgen"
    ElseIf dateDat = DateValue(Now) - 1 Then
        lblx.Caption = "Gestern"
    ElseIf dateDat = DateValue(Now) - 2 Then
        lblx.Caption = "Vorgestern"
    Else
        lblx.Caption = ""
    End If
    
    lblx.Refresh
    
    Dim dateVerf As Date
    dateVerf = DateValue(Mid(Combo2.Text, 4, 8))
    verfuegbar dateVerf 'dateDat
    Label2(8).Caption = Combo2.Text
    Label2(8).Refresh

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GesternOderMorgen"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim i       As Integer
    Dim ctmp    As String
    
    gsLastKunde = ""
    
    gsfrmComeFrom = "Terminkalender"
    
    PositionierenWKL82
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Command3(2).BackColorTo = vbWhite
    Command3(2).BackColorFrom = vbWhite
    
    Command3(1).BackColorTo = vbWhite
    Command3(1).BackColorFrom = vbWhite
    
    Command3(14).BackColorTo = vbWhite
    Command3(14).BackColorFrom = vbWhite
    
    Command3(15).BackColorTo = vbWhite
    Command3(15).BackColorFrom = vbWhite
    
    Command1(0).BackColorTo = vbWhite
    Command1(0).BackColorFrom = vbWhite
    
    lese_Termin_Optionen
    
    If FileExists(App.Path & "\NoSkalieren.cfg") Then
        dWidth = 12000
    Else
        dWidth = Screen.Width
    End If
    
    Command6(3).Left = (dWidth / 5 * 4) - Command6(3).Width
    Command6(3).Caption = "+"
    
    Frame3.Width = dWidth / 5 * 4
    MSFlexGrid1.Width = dWidth / 5 * 4
    
    Frame3.ForeColor = vbBlue
    
    If NewTableSuchenDBKombi("TERMINDAUER", gdBase) = False Then
        CreateTableT2 "TERMINDAUER", gdBase
    End If
    
    If NewTableSuchenDBKombi("TERMDEL", gdBase) = False Then
        CreateTableT2 "TERMDEL", gdBase
    End If
    
    If Datendrin("Openings", gdBase) = False Then
        loeschNEW "OPENINGS", gdBase
        CreateTable "OPENINGS", gdBase
        For i = 1 To 7
            cSQL = "Insert into Openings (WOTAG,LFDNR,VON,BIS,ZEITBLOCK) values (" & i & ",1,'09:00','18:00',15) "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into Openings (WOTAG,LFDNR,VON,BIS,ZEITBLOCK) values (" & i & ",2,'','',15) "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into Openings (WOTAG,LFDNR,VON,BIS,ZEITBLOCK) values (" & i & ",3,'','',15) "
            gdBase.Execute cSQL, dbFailOnError
        Next i
    End If
    
    If Datendrin("PFLEGORT", gdBase) = False Then
        loeschNEW "PFLEGORT", gdBase
        CreateTable "PFLEGORT", gdBase
        
        For i = 1 To 8
            ctmp = "Kabine " & i
            cSQL = "Insert into PFLEGORT (Bezeich,anzeigen) values ('" & ctmp & "',1) "
            gdBase.Execute cSQL, dbFailOnError
        Next i
    End If
    
    CheckIndex "Termine", "datum", "", gdBase
    CheckIndex "Termine", "bednu", "", gdBase
    CheckIndex "Termine", "uhrzeit", "", gdBase
    CheckIndex "Termine", "Buchungsnr", "", gdBase
    
    CheckIndex "OPENINGS", "WOTAG", "", gdBase
    CheckIndex "OPENINGS", "LFDNR", "", gdBase
    
    cSQL = "Delete * from TERMINE where BEDNU = 0 "
    gdBase.Execute cSQL, dbFailOnError
     
    
    LeseStandardTexteWKL82 ""
    LeseOpeningsWKL82
    
    Screen.MousePointer = 11
'    entferneoverandunder
    Screen.MousePointer = 0
    
    fuellecboBedienerWKL82
    fuellecboDruckansicht
    
    
    
    fuellecboDatum
    
    fuellecombo Combo4
    
    Label3(6).Caption = "Heute"
    Label3(6).Refresh

    If gcTerm_Datum <> "" Then
        Combo2.Text = Left(WeekdayName(Weekday(DateValue(gcTerm_Datum), vbMonday)), 2) & " " & Format(DateValue(gcTerm_Datum), "DD.MM.YY")
    End If
    ZeigeTermineJeMitarbeiterWKL82 ""

    GesternOderMorgen DateValue(Mid(Combo2.Text, 4, 8)), Label3(6)


    If gbBILDTAST = False Then
        Frame0.Visible = False
    Else
        Frame0.Visible = True
    End If
    
    Label0.Caption = "1000"
    
    Label2(4).ForeColor = vbRed
    Label2(5).ForeColor = vbRed
    Label2(3).ForeColor = vbRed
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    cbox.Clear
    
    sSQL = "select distinct(gliederung) from TERM_STD  order by gliederung "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Gliederung) Then
                cbox.AddItem rsrs!Gliederung
                If cbox.Text = "" Then
                    cbox.Text = rsrs!Gliederung
                End If
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
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo_Zeitraeume(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR

    cbox.Clear
    
    '1.Eintrag
    cbox.AddItem "nächste Woche"
    cbox.AddItem "14 Tage"
    cbox.AddItem MonthName(Month(DateValue(Now))) & " " & Year(DateValue(Now)) 'aktueller Monat
    
    '2.Eintrag
    If Month(DateValue(Now)) = 12 Then
        cbox.AddItem MonthName(1) & " " & Year(DateValue(Now)) + 1
    Else
        cbox.AddItem MonthName(Month(DateValue(Now)) + 1) & " " & Year(DateValue(Now))
    End If
    
    '3.Eintrag

    If Month(DateValue(Now)) = 11 Then
        cbox.AddItem MonthName(1) & " " & Year(DateValue(Now)) + 1
    ElseIf Month(DateValue(Now)) = 12 Then
        cbox.AddItem MonthName(2) & " " & Year(DateValue(Now)) + 1
    Else
        cbox.AddItem MonthName(Month(DateValue(Now)) + 2) & " " & Year(DateValue(Now))
    End If
    
    
    
    cbox.AddItem "Heute"
    cbox.AddItem "Morgen"
    cbox.AddItem "Übermorgen"
'    cbox.AddItem "dieser Monat"
'    cbox.AddItem "eigener Zeitraum"
    
    'erweiterbar um gängige
    'November 2018
    'Dezember 2018
    
    
    cbox.Text = "Heute"
               
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo_Zeitraeume"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL82()
On Error GoTo LOKAL_ERROR

'    Frame10.Top = 1320
'    Frame10.Left = 1560
'    Frame10.Height = 2415
'    Frame10.Width = 4695

    Frame10.Top = 960
    Frame10.Left = 0
    Frame10.Height = 7455
    Frame10.Width = 9615
    Frame10.BorderStyle = 1
    
    Frame4.Top = 0
    Frame4.Left = 0
    Frame4.Height = 9000
    Frame4.Width = 12000
    
    Frame12.Top = 1440
    Frame12.Left = 1560
    Frame12.Height = 3495
    Frame12.Width = 5655
    
    
    Frame11.Top = 0
    Frame11.Left = 0
    Frame11.Height = 9000
    Frame11.Width = 12000
    
    Frame3.Top = 0
    Frame3.Left = 0
    Frame3.Height = 8655
    
    Frame0.Top = 6480
    Frame0.Left = 120
    Frame0.Height = 2055
    Frame0.Width = 9135
    
    Frame9.Top = 120
    Frame9.Left = 9660
    Frame9.Height = 7695
    Frame9.Width = 2175
    
    
    
    Label6.Top = 240
    Label6.Left = 9840
    Label6.Height = 2655 '2775 '3735
    Label6.Width = 2175
    
    Label11.Top = Label6.Top + Label6.Height
    Label11.Left = 9840
    Label11.Height = 375 '240
    Label11.Width = 2175
    
    Frame13.Top = 0
    Frame13.Left = 120
    Frame13.Height = 3495
    Frame13.Width = 11535
    
    Frame14.Top = 0
    Frame14.Left = 120
    Frame14.Height = 8415
    Frame14.Width = 11655
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL82"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "TPLAN", gdBase
    loeschNEW "TERMPRINT", gdBase
    loeschNEW "TERMPRINT_EP", gdBase
    loeschNEW "TERMPRINT_MEP", gdBase
    loeschNEW "Terme_DEl_" & srechnertab, gdBase
    LogtoEnd Me
    gsfrmComeFrom = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1(0).ForeColor = glS1
End Sub
Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1(19).ForeColor = glS1
End Sub
Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    If index = 19 Then
        frmWKL193.Show 1
        
    ElseIf index = 0 Then
        gckundnr = Val(Left(Label2(2).Caption, InStr(1, Label2(2).Caption, " ")))
        
        
        If SucheUnter(gckundnr) Then
            MsgBox "Achtung: Noch offene Kassiervorgänge für diesen Kunden!", vbInformation, "Winkiss Hinweis:"
        End If
        
'        lese_Termin_Optionen
        If gbTerm_BedKass = True Then
            gcTerm_Bed = Trim$(Left(Label2(0).Caption, 3)) ' Text1(4).Text
        Else
            gcTerm_Bed = ""
        End If
        
        'auch Artikelarray füllen
        
        fuelle_Artikel_Array
        
        Unload frmWKL82
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If index = 19 Then
        Label1(19).ForeColor = glLink
    End If
    
    If index = 0 Then
        Label1(0).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermSelBed() As Integer
On Error GoTo LOKAL_ERROR
    
    ermSelBed = 0
    
    Dim i As Integer
    
    If Tree11.SelectedItem Is Nothing Then
        Exit Function
    Else
        For i = 1 To Tree11.Nodes.Count
            If Tree11.Nodes(i).Selected = True Then
        
                ermSelBed = Tree11.Nodes(i).Tag
                
                Exit For
            End If
        Next i
    End If
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
    

Private Sub Label2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cKürzel As String
    Dim cName As String
    Dim cVorname As String
    Dim SFarbe As String
    
    If index = 21 Then
    
    
        
    
    
    
    
        If Label2(21).Caption = "ganzer Tag" Then
            Label2(21).Caption = "vormittags"
            zeige_Freie_Termine ermSelBed, Label2(21).Caption, Check2(0), Check2(1), Check2(2), Check2(3), Check2(4), Check2(5), Check2(6)
        ElseIf Label2(21).Caption = "vormittags" Then
            Label2(21).Caption = "nachmittags"
            zeige_Freie_Termine ermSelBed, Label2(21).Caption, Check2(0), Check2(1), Check2(2), Check2(3), Check2(4), Check2(5), Check2(6)
        ElseIf Label2(21).Caption = "nachmittags" Then
            Label2(21).Caption = "ganzer Tag"
            zeige_Freie_Termine ermSelBed, Label2(21).Caption, Check2(0), Check2(1), Check2(2), Check2(3), Check2(4), Check2(5), Check2(6)
        End If
        
    End If
    
    If index = 6 Then
        Frame9.Visible = True
    End If
    
    If index = 19 Then
        Text1(5).Text = ""
        List10.Clear
        
        Label2(20).Visible = False
        Label2(21).Visible = False
        
        Check2(0).Visible = False
        Check2(1).Visible = False
        Check2(2).Visible = False
        Check2(3).Visible = False
        Check2(4).Visible = False
        Check2(5).Visible = False
        Check2(6).Visible = False
    
    
        Tree11.Visible = False
        
        DTPickerVon.Visible = False
        DTPickerBis.Visible = False
        
        Combo14.Visible = False
        List11.Visible = False
    End If
    
    If index = 10 Then
    
        If Label2(10).Caption = "alle anzeigen" Then
            Label2(10).Caption = "einschränken"
            ZeigeTermineJeMitarbeiterWKL82 "alle"
        Else
            Label2(10).Caption = "alle anzeigen"
            ZeigeTermineJeMitarbeiterWKL82 ""
        End If
    End If
    
    
    If index = 14 Then
        Screen.MousePointer = 11
        
        'gsLastKunde
        
        
        Label2(2).Caption = ""
        Label2(2).BackColor = glH1
        Label2(2).Refresh
        
        List6.Clear
        List2.Clear
    
        
    
        If gsLastKunde <> "" Then
            If IsNumeric(gsLastKunde) Then
            
                SFarbe = ermFarbe(Trim(gsLastKunde))
                If Trim(SFarbe) = "0" Then
                    Label2(2).BackColor = glH1
                Else
                    Label2(2).BackColor = glfarbe(SFarbe)
                End If
            
                cKürzel = lookingForKundendaten(gsLastKunde).Kuerzel
                cName = lookingForKundendaten(gsLastKunde).nachname
                cVorname = lookingForKundendaten(gsLastKunde).vorname
            
                Label2(2).Caption = gsLastKunde & "  "
                Label2(2).Caption = Label2(2).Caption & Space$(5 - Len(cKürzel)) & cKürzel & "  "
                Label2(2).Caption = Label2(2).Caption & cName & ", "
                Label2(2).Caption = Label2(2).Caption & cVorname
                
                lblUnter.Visible = False
                
                If SucheUnter(gsLastKunde) Then
                    lblUnter.ForeColor = glWarn
                    lblUnter.Visible = True
                End If
                
                SucheBediener gsLastKunde, List6
                SucheGelöschteTermine gsLastKunde
                DS_Unterschrieben gsLastKunde
                
                
                
    
            End If
        End If
        gckundnr = ""
        
        If Label2(2).Caption <> "" Then
            Label3(4).Caption = Label2(2).Caption
            Label3(4).BackColor = Label2(2).BackColor
            
            If Gibt_es_Termine_in_Zunkunft(Left(Label3(4).Caption, InStr(1, Label3(4).Caption, " "))) > 0 Then
                Label3(4).ToolTipText = "weitere Termine (Doppelklick)"
                Label3(4).FontUnderline = True
                Label3(4).ForeColor = glLink
            Else
                Label3(4).ToolTipText = ""
                Label3(4).FontUnderline = False
                Label3(4).ForeColor = glS1
            End If
        End If
        
    End If
        
    If index = 16 Then
        Screen.MousePointer = 11
    
        'gsLastKunde
        
        Label2(7).Caption = ""
        Label2(7).BackColor = glH1
        Label2(7).Refresh
        
        List5.Clear
        List1.Clear
    
        
    
        If gsLastKunde <> "" Then
            If IsNumeric(gsLastKunde) Then
            
                SFarbe = ermFarbe(Trim(gsLastKunde))
                If Trim(SFarbe) = "0" Then
                    Label2(7).BackColor = glH1
                Else
                    Label2(7).BackColor = glfarbe(SFarbe)
                End If
            
                cKürzel = lookingForKundendaten(gsLastKunde).Kuerzel
                cName = lookingForKundendaten(gsLastKunde).nachname
                cVorname = lookingForKundendaten(gsLastKunde).vorname
            
                Label2(7).Caption = gsLastKunde & "  "
                Label2(7).Caption = Label2(7).Caption & Space$(5 - Len(cKürzel)) & cKürzel & "  "
                Label2(7).Caption = Label2(7).Caption & cName & ", "
                Label2(7).Caption = Label2(7).Caption & cVorname
                
                SucheBediener gsLastKunde, List1
                SucheGelöschteTermine gsLastKunde
                DS_Unterschrieben gsLastKunde
    
            End If
        End If
        gckundnr = ""
        
        If Label2(7).Caption <> "" Then
            Command4(8).Caption = "wählen"
        Else
            Command4(8).Caption = "schließen"
        End If
        
    End If
        
    Screen.MousePointer = 0
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Gibt_es_Termine_in_Zunkunft(Left(Label3(4).Caption, InStr(1, Label3(4).Caption, " "))) > 0 Then
        Frame12.Visible = True
        ZeigeTermine Trim(Left(Label3(4).Caption, InStr(1, Label3(4).Caption, " "))), List8
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List4_Click()
    On Error GoTo LOKAL_ERROR
    
    Label2(1).Caption = Trim$(UCase$(List4.list(List4.ListIndex)))
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List7_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command6_Click 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List7_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List8_dblClick()
On Error GoTo LOKAL_ERROR

    Dim cDateum As String
    Dim cLBSatz As String
    
    If List8.ListIndex < 0 Then

    Else
        cLBSatz = List8.list(List8.ListIndex)
        cDateum = Mid(cLBSatz, 1, 8)
        Combo2.Text = Left(WeekdayName(Weekday(DateValue(cDateum), vbMonday)), 2) & " " & cDateum
            
        GesternOderMorgen DateValue(cDateum), Label3(6)
        AktualisiereTerminTabelleWKL82
        
        Frame12.Visible = False
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List8_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lcol As Long
    Dim lrow As Long
    Dim cbednu As String
    Dim cDatum As String
    Dim czeit As String
    Dim cOrt As String
    Dim cAnzeige As String
    Dim cBuchnr As String
    
    Label6.Caption = ""
    Label6.Refresh
    
    Label11.Caption = ""
    Label11.Refresh
    Text1(4).Text = ""
    
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    'brauch ich um den Termin einzufügen
    globRow = lrow
    globCol = lcol
    
    If lcol = 0 Then
        Exit Sub
    End If
    
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = lcol
    
    Label3(7).Caption = MSFlexGrid1.TextMatrix(lrow, 0)
    Label3(7).Refresh
    
    cAnzeige = MSFlexGrid1.Text
    cAnzeige = Trim$(cAnzeige)
    
    'Mitarbeiter
    If MSFlexGrid1.Text <> "" Then
        cbednu = Mid(MSFlexGrid1.Text, 51, Len(MSFlexGrid1.Text) - 50)
    End If
    
    'Datum
    Text1(0).Text = Right(Combo2.Text, 8)
    cDatum = Right(Combo2.Text, 8)
    
    'abwesend Datum
    Label3(5) = Right(Combo2.Text, 8)
    Label3(12) = Right(Combo2.Text, 8)
    
    'Uhrzeit
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = 0
    czeit = MSFlexGrid1.Text
    
    fuellecombo5 czeit
    
    fuellecombo6
    
    'Ort
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = lcol
    cOrt = MSFlexGrid1.Text
    
    Text3.Text = ""
    If cAnzeige <> "" Then
        fnHoleKundendaten cDatum, czeit, cbednu, cOrt
        
        Label10.Caption = ""
        Frame9.Visible = False
    Else
    
    
        Label8.Caption = "0"
    
    
        If globBuchNr > 0 Then
            Command3(4).Caption = "Termin einfügen"
            Command3(4).ForeColor = vbRed
            label10_Kundendaten (globBuchNr)
            
            Frame9.Visible = False
        Else
            Frame9.Visible = True
            
        End If
        
        
        
        
        
    End If
    
    MSFlexGrid1.Col = lcol
    MSFlexGrid1.Row = lrow
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo5(czeit As String)
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    
    Combo5.Clear
    
    MSFlexGrid1.Redraw = False
    
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lcount
        MSFlexGrid1.Col = 0
        If TimeValue(czeit) < TimeValue(MSFlexGrid1.Text) Then
            Combo5.AddItem MSFlexGrid1.Text
        End If
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
    Combo5.AddItem gcEndeZeit
    Combo5.Text = gcEndeZeit
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo5"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo_Uhrzeiten_von()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim dUhrzeit As Double
    Dim iAnzahl As Integer

    Combo11.Clear

    iAnzahl = (TimeValue(gcEndeZeit) - TimeValue(gcStartZeit)) / TimeValue(gcZeitBlock)
    dUhrzeit = TimeValue(gcStartZeit)
    
    For lcount = 1 To iAnzahl
        Combo11.AddItem Format$(dUhrzeit + (TimeValue(gcZeitBlock) * lcount), "HH:MM")
    Next lcount
    
    Combo11.Text = Format$(dUhrzeit + (TimeValue(gcZeitBlock)), "HH:MM")
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo_Uhrzeiten_von"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo_Uhrzeiten_bis()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim dUhrzeit As Double
    Dim iAnzahl As Integer

    Combo8.Clear

    iAnzahl = (TimeValue(gcEndeZeit) - TimeValue(gcStartZeit)) / TimeValue(gcZeitBlock)
    dUhrzeit = TimeValue(gcStartZeit)
    
    For lcount = 1 To iAnzahl
        Combo8.AddItem Format$(dUhrzeit + (TimeValue(gcZeitBlock) * lcount), "HH:MM")
    Next lcount
    
    Combo8.Text = gcEndeZeit
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo_Uhrzeiten_bis"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo7()
On Error GoTo LOKAL_ERROR

    Combo7.Clear
    
    Combo7.AddItem "alle Tage"
    Combo7.AddItem "nur montags"
    Combo7.AddItem "nur dienstags"
    Combo7.AddItem "nur mittwochs"
    Combo7.AddItem "nur donnerstags"
    Combo7.AddItem "nur freitags"
    Combo7.AddItem "nur samstags"
    Combo7.AddItem "nur sonntags"
    
    Combo7.Text = "alle Tage"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo7"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecombo6()
On Error GoTo LOKAL_ERROR

    Combo6.Clear
    
    Combo6.AddItem "alle Tage"
    Combo6.AddItem "nur montags"
    Combo6.AddItem "nur dienstags"
    Combo6.AddItem "nur mittwochs"
    Combo6.AddItem "nur donnerstags"
    Combo6.AddItem "nur freitags"
    Combo6.AddItem "nur samstags"
    Combo6.AddItem "nur sonntags"
    
    Combo6.Text = "alle Tage"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo6"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lcol    As Long
    Dim sTemp   As String
    Dim iRet    As Integer
    
    'Geklickte Zelle merken
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    
    If MSFlexGrid1.Text = "ORT" Then
        MsgBox "Terminvergabe bitte über TERMINE/MITARBEITER", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = lcol
    If MSFlexGrid1.Text = "" Then
    
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.Col = 1
        If MSFlexGrid1.Text <> "" Then
            sTemp = "Terminvergabe hier nicht möglich!" & vbCrLf & vbCrLf
            sTemp = sTemp & "Möchten Sie trotzdem einen Termin vereinbaren?"
            iRet = MsgBox(sTemp, vbYesNo + vbQuestion, "Winkiss Frage:")
            
            MSFlexGrid1.Col = lcol
            
            If iRet = vbNo Then
                Exit Sub
            End If
        End If
        
    End If
    MSFlexGrid1.Col = lcol
    
    EinzelDatenMitarbeiterWKL82
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_SelChange()
On Error GoTo LOKAL_ERROR

Dim cBezeich As String
Dim lDauer As Long


cBezeich = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)
lDauer = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2)


If cBezeich <> "" Then
    List10.AddItem cBezeich
    Text1(5).Text = CStr(Val(Text1(5).Text) + lDauer)
    
    Label2(20).Visible = True
    Label2(21).Visible = True
    Label2(21).Caption = "ganzer Tag"
    
    Check2(0).Visible = True
    Check2(1).Visible = True
    Check2(2).Visible = True
    Check2(3).Visible = True
    Check2(4).Visible = True
    Check2(5).Visible = True
    Check2(6).Visible = True
    
    Check2(0).value = vbChecked
    Check2(1).value = vbChecked
    Check2(2).value = vbChecked
    Check2(3).value = vbChecked
    Check2(4).value = vbChecked
    Check2(5).value = vbChecked
    Check2(6).value = vbChecked
    
    
    
    
    Tree11.Visible = True
    
    Combo14.Visible = True
    
    DTPickerVon.Visible = True
    DTPickerBis.Visible = True
    
    DTPickerVon.value = DateValue(Now)
    DTPickerBis.value = DateValue(Now)
    
    fuellecombo_Zeitraeume Combo14
    
    zeigefreie_Terminbloecke_ProBediener Text1(5).Text, DTPickerVon.value, DTPickerBis.value
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigefreie_Terminbloecke_ProBediener(sDauer As String, dateVon As Date, dateBis As Date)
On Error GoTo LOKAL_ERROR

    List11.Clear
    Tree11.Nodes.Clear
    
    Screen.MousePointer = 11
    
    loeschNEW "FREIEZEIT", gdBase
    CreateTableT3 "FREIEZEIT", gdBase

    'bedienerschleife
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cbednu As String
    Dim cbedname As String
    Dim iDate As Date
    Dim bFeiertag As Boolean
    
    cSQL = "Select * from BEDTERM order by bednu desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                cbednu = rsrs!BEDNU
            Else
                cbednu = ""
            End If
            
            If Not IsNull(rsrs!bedname) Then
                cbedname = rsrs!bedname
            Else
                cbedname = ""
            End If
            
            'freie Zeitblöcke pro Bediener
            
            For iDate = dateVon To dateBis
            
                bFeiertag = False
            
                If IsThis_EinFeiertag(Format(iDate, "DD.MM.YYYY")) Then
                    bFeiertag = True
                End If
                
                If bFeiertag = False Then
                    ermittle_freiZeitBloecke_ProBed iDate, cbednu
                End If
            Next iDate
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Delete * from FREIEZEIT where DauerinMinuten <  " & CInt(sDauer)
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "FREIE_TERMINE", gdBase
    CreateTableT3 "FREIE_TERMINE", gdBase
    
    Dim sStartZeit As String
    Dim sEndeZeit As String
    Dim dateAktuell As Date
    


    cSQL = "Select * from FREIEZEIT  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDIENER) Then
                cbednu = rsrs!BEDIENER
            Else
                cbednu = ""
            End If
            
            If Not IsNull(rsrs!startzeit) Then
                sStartZeit = rsrs!startzeit
            Else
                sStartZeit = ""
            End If
            
            If Not IsNull(rsrs!endezeit) Then
                sEndeZeit = rsrs!endezeit
            Else
                sEndeZeit = ""
            End If
            
            If Not IsNull(rsrs!Datum) Then
                dateAktuell = rsrs!Datum
            Else
                dateAktuell = ""
            End If
            
            WasGehtZwischenStartundEnde sStartZeit, sEndeZeit, sDauer, cbednu, dateAktuell
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ZeigSumme_FreieTerminproBediener
    
    Screen.MousePointer = 0


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigefreie_Terminbloecke_ProBediener"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigSumme_FreieTerminproBediener()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsBed As Recordset
    Dim ibednu As Integer
    Dim lFarbe As Long
    Dim lAnzahl As Long
    Dim cLBSatz As String
    
    loeschNEW "SUM_FREIE_TERMINE", gdBase
    CreateTableT3 "SUM_FREIE_TERMINE", gdBase
    
    cSQL = "Insert into SUM_FREIE_TERMINE  "
    cSQL = cSQL & " Select count(*) as Anzahl "
    cSQL = cSQL & ", Bediener from Freie_termine group by Bediener "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update SUM_FREIE_TERMINE inner join BEDTERM on SUM_FREIE_TERMINE.Bediener = BEDTERM.bednu "
    cSQL = cSQL & " set SUM_FREIE_TERMINE.FARBCODE = BEDTERM.FARBCODE "
    gdBase.Execute cSQL, dbFailOnError
    
    Tree11.Nodes.Clear
    
    


    cSQL = "Select * from SUM_FREIE_TERMINE order by Bediener asc "
    Set rsBed = gdBase.OpenRecordset(cSQL)
    If Not rsBed.EOF Then
        rsBed.MoveFirst
        Do While Not rsBed.EOF
        
            If Not IsNull(rsBed!BEDIENER) Then
                ibednu = rsBed!BEDIENER
            Else
                ibednu = 0
            End If
            
            If Not IsNull(rsBed!ANZAHL) Then
                lAnzahl = rsBed!ANZAHL
            Else
                lAnzahl = 0
            End If
            
            If Not IsNull(rsBed!FARBCODE) Then
                lFarbe = rsBed!FARBCODE
            Else
                lFarbe = 0
            End If
            

            cLBSatz = lAnzahl & " freie Termine bei " & ermBEDbez(CLng(ibednu))
                            
            Tree11.Nodes.Add Text:=cLBSatz
            Tree11.Nodes(Tree11.Nodes.Count).Tag = ibednu
            Tree11.Nodes(Tree11.Nodes.Count).BackColor = FarbeBackColor(lFarbe)
            Tree11.Nodes(Tree11.Nodes.Count).ForeColor = FarbeForeColor(lFarbe)
                            
               
            
            rsBed.MoveNext
        Loop
    End If
    rsBed.Close: Set rsBed = Nothing
    
    Tree11.Refresh
    
    
    
    
    
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigSumme_FreieTerminproBediener"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub ermittle_freiZeitBloecke_ProBed(dateVon As Date, sBednu As String)
On Error GoTo LOKAL_ERROR

    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim rsBuch As Recordset
    Dim lBuchungsnr As Long
    Dim lZeit1 As Long
    Dim cFeld As String
    Dim bFirst As Boolean
    Dim lFarbe As Long
    
    Dim ldynStartzeit As Long
    
    Dim lStartzeit As Long
    Dim lEndzeit As Long
    Dim lzeitblock As Long
    Dim cLBSatz As String
    
    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim cDauer As String
    
    cFeld = gcZeitBlock
    cFeld = SwapStr(cFeld, ":", "")
    lzeitblock = CLng(cFeld)
    
    Dim iWeekday As Integer
    iWeekday = Weekday(dateVon, vbMonday)
    
    Dim lMinIndex As Long
    Dim lMaxIndex As Long
    
    lMinIndex = ((iWeekday - 1) * 3) + 1
    lMaxIndex = ((iWeekday - 1) * 3) + 3
    
    Dim dUhrzeit As Double
    Dim sZeitblock As String
    sZeitblock = SwapStr(gcZeitBlock, ":", "")
    dUhrzeit = Val(sZeitblock) / 1440
    Dim dZeit As Double
    Dim dStartzeit As Double
    
    Dim k As Integer
    
    For k = lMinIndex To lMaxIndex

        If gZeiten(k).Von <> "" And gZeiten(k).Bis <> "" Then
        
            dZeit = TimeValue(gZeiten(k).Von)
            dStartzeit = dZeit - dUhrzeit
            
            cFeld = Format$(dStartzeit, "HH:MM") 'gcStartZeit
            cFeld = SwapStr(cFeld, ":", "")
            lStartzeit = CLng(cFeld)
            
            lStartzeit = lStartzeit + lzeitblock
                            
            If Right(CStr(lStartzeit), 2) = "60" Then
                lStartzeit = lStartzeit + 40
            End If
            
            cFeld = gZeiten(k).Bis 'gcEndeZeit
            cFeld = SwapStr(cFeld, ":", "")
            lEndzeit = CLng(cFeld)
            
            ldynStartzeit = lStartzeit
            
            bFirst = True
            
            cSQL = "Select max(Buchungsnr) as maxBuch,min(uhrzeit) as mini from Termine where "
            cSQL = cSQL & " datum = " & CLng(dateVon) & " "
            cSQL = cSQL & " and bednu = " & sBednu
            cSQL = cSQL & " group by Buchungsnr order by min(uhrzeit) asc"
            Set rsBuch = gdBase.OpenRecordset(cSQL)
            If Not rsBuch.EOF Then
                rsBuch.MoveFirst
                Do While Not rsBuch.EOF
                    If Not IsNull(rsBuch!maxBuch) Then
                        lBuchungsnr = rsBuch!maxBuch
                    Else
                        lBuchungsnr = 0
                    End If
        
                    lZeit1 = ermMinZeitperBuchung(lBuchungsnr)
                    
                    If lZeit1 > ldynStartzeit Then
                    
                        If bFirst = True Then
                            cLBSatz = ermBEDbez(CLng(sBednu))
                            
                            bFirst = False
                        End If
                        
                        dStart = TimeValue(zeitanz(ldynStartzeit))
                        dEnde = TimeValue(zeitanz(lZeit1))
                
                        dDauer = dEnde - dStart
                        cDauer = Format$(dDauer, "HH:MM")
                    
                        cLBSatz = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lZeit1)
                        
                        insert_FreieZeit zeitanz(ldynStartzeit), zeitanz(lZeit1), cDauer, sBednu, dateVon
        
                    End If
                    
                    ldynStartzeit = ermMaxZeitperBuchung(lBuchungsnr)
                    ldynStartzeit = ldynStartzeit + lzeitblock

                    If Right(CStr(ldynStartzeit), 2) = "60" Then
                        ldynStartzeit = ldynStartzeit + 40
                    End If
                    
        
                rsBuch.MoveNext
                Loop
            Else
            
                cLBSatz = ermBEDbez(CLng(sBednu))
        
                dStart = TimeValue(zeitanz(lStartzeit))
                dEnde = TimeValue(zeitanz(lEndzeit))
        
                dDauer = dEnde - dStart
                cDauer = Format$(dDauer, "HH:MM")
                
                cLBSatz = "von " & zeitanz(lStartzeit) & " bis " & gZeiten(k).Bis
                insert_FreieZeit zeitanz(lStartzeit), gZeiten(k).Bis, cDauer, sBednu, dateVon
            End If
            
            If ldynStartzeit <> lStartzeit Then
                If ldynStartzeit < lEndzeit Then
                    If bFirst = True Then
                        cLBSatz = ermBEDbez(CLng(sBednu))
                        
        
                        bFirst = False
                    End If
                    
                    dStart = TimeValue(zeitanz(ldynStartzeit))
                    dEnde = TimeValue(zeitanz(lEndzeit))
        
                
                    dDauer = dEnde - dStart
                    cDauer = Format$(dDauer, "HH:MM")
                
                
                    cLBSatz = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lEndzeit)
                    insert_FreieZeit zeitanz(ldynStartzeit), zeitanz(lEndzeit), cDauer, sBednu, dateVon
                    
                End If
            End If
            
            rsBuch.Close: Set rsBuch = Nothing
        End If
    Next k
        
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittle_freiZeitBloecke_ProBed"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub insert_FreieZeit(sStartZeit As String, sEndeZeit As String, sDauer As String, sbed As String, dateTag As Date)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    Dim iDauerinMinuten As Integer
    
    If Len(sDauer) = 5 Then
        iDauerinMinuten = (Left(sDauer, 2) * 60) + Right(sDauer, 2)
    ElseIf Len(sDauer) = 4 Then
        iDauerinMinuten = (Left(sDauer, 1) * 60) + Right(sDauer, 2)
    End If
    
    sSQL = "Insert into FreieZeit (Startzeit,Endezeit, Dauer, DauerInMinuten, Bediener, Datum) values  "
    sSQL = sSQL & " ( '" & sStartZeit & "','" & sEndeZeit & "','" & sDauer & "'," & iDauerinMinuten & "," & sbed & ",'" & dateTag & "') "
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insert_FreieZeit"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WasGehtZwischenStartundEnde(sStartZeit As String, sEndeZeit As String, sDauer As String, sbed As String, dateTag As Date)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sDauerTermin As String
    
    
    Dim dStartZeitraum As Double
    Dim dEndeZeitraum As Double
    
    Dim dStartTermin As Double
    Dim dEndeTermin As Double
    Dim dDauerTermin As Double
    
    Dim dZeitblock As Double
    Dim sZeitblock As String
    sZeitblock = SwapStr(gcZeitBlock, ":", "")
    
    dZeitblock = Val(sZeitblock) / 1440
    
    Dim dUhrzeit As Double
    dUhrzeit = Val(sDauer) / 1440
    sDauerTermin = Format$(dUhrzeit, "HH:MM")
    
    
    dDauerTermin = TimeValue(sDauerTermin)
    
    dStartZeitraum = TimeValue(sStartZeit)
    dEndeZeitraum = TimeValue(sEndeZeit)
    
    dStartTermin = dStartZeitraum
    


    
    Do While Round(dStartTermin + dDauerTermin, 10) <= Round(dEndeZeitraum, 10)
    
'        MsgBox dStartTermin + dDauerTermin & " " & dEndeZeitraum
    
        dEndeTermin = dStartTermin + dDauerTermin
        insert_Freie_TERMINE Format$(dStartTermin, "HH:MM"), Format$(dEndeTermin, "HH:MM"), sDauer, sbed, dateTag
        
        dStartTermin = dStartTermin + dZeitblock 'TimeValue(gcZeitBlock)
        
    Loop
    
'    MsgBox "Danach " & dStartTermin + dDauerTermin & " " & dEndeZeitraum
    


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WasGehtZwischenStartundEnde"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub insert_Freie_TERMINE(sStartZeit As String, sEndeZeit As String, sDauer As String, sbed As String, dateTag As Date)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    Dim iDauerinMinuten As Integer
    
    If Len(sDauer) = 5 Then
        iDauerinMinuten = (Left(sDauer, 2) * 60) + Right(sDauer, 2)
    ElseIf Len(sDauer) = 4 Then
        iDauerinMinuten = (Left(sDauer, 1) * 60) + Right(sDauer, 2)
    End If
    
    Dim sWeekdayname As String
    sWeekdayname = WeekdayName(Weekday(dateTag, vbMonday))
    
    Dim sTageszeit As String
    
    If Val(Left(sStartZeit, 2)) <= 13 Then
        sTageszeit = "vormittag"
    Else
        sTageszeit = "nachmittag"
    End If
   
    
    sSQL = "Insert into Freie_TERMINE (Startzeit,Endezeit, Dauer, DauerInMinuten, Bediener, Datum,WOCHENTAG,TAGESZEIT) values  "
    sSQL = sSQL & " ( '" & sStartZeit & "','" & sEndeZeit & "','" & sDauer & "'," & iDauerinMinuten & " "
    sSQL = sSQL & " ," & sbed & ",'" & dateTag & "','" & sWeekdayname & "','" & sTageszeit & "') "
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insert_Freie_TERMINE"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(index).BackColor = glSelBack1
    Label0.Caption = index
    If index = 1 Or index = 4 Then
        Command0(41).Caption = ":"
    Else
        Command0(41).Caption = "."
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case index
    
        Case Is = 5
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%;,:.-_"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
        
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command4_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheBediener(cKundnr As String, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    Dim cLBSatz As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    
    Listx.Clear
   
    If cKundnr = "" Then
        Exit Sub
    End If
    
    If Not IsNumeric(cKundnr) Then
        Exit Sub
    End If
    
    loeschNEW "BELBED", gdBase
    
    cSQL = "Select distinct(Buchungsnr),bednu into BELBED from Termine where Kundnr = " & cKundnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select count(bednu)as count,bednu from BELBED group by bednu" ' order by count "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Count) Then
                cFeld = rsrs!Count
            Else
                cFeld = "0"
            End If
           
            cFeld = Space$(4 - Len(cFeld)) & cFeld
            cSatz = cFeld & "x von "
            
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = "0"
            End If
            
            cFeld = ermBEDbez(CLng(cFeld))
            cSatz = cSatz & cFeld
            
            Listx.AddItem cSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheBediener"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheGelöschteTermine(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    Dim cLBSatz As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    
    List2.Clear
    List5.Clear
   
    If cKundnr = "" Then
        Exit Sub
    End If
    
    If Not IsNumeric(cKundnr) Then
        Exit Sub
    End If
    
    cSQL = "Select * from TERMDEL where Kundnr = " & cKundnr & "  order by adate desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
                cFeld = rsrs!ADATE
            End If
            List2.AddItem cFeld
            List5.AddItem cFeld
            
            If Not IsNull(rsrs!GRUND) Then
                cFeld = rsrs!GRUND
            End If
            List2.AddItem cFeld
            List5.AddItem cFeld
           
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If List2.ListCount > 0 Then
        Command4(15).BackColor = vbRed
        Command4(15).Caption = "Achtung"
        
        Command4(16).BackColor = vbRed
        Command4(16).Caption = "Achtung"
        
        gckundnr = cKundnr
        If Nicht_erschienen_vorhanden(cKundnr) Then
            frmWKL212.Show 1
            
        End If
        gckundnr = ""
    Else
        Command4(15).BackColor = Command4(7).BackColor
        Command4(16).BackColor = Command4(7).BackColor
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheGelöschteTermine"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DS_Unterschrieben(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    
    Command4(22).ForeColor = Command4(7).ForeColor
   
    If cKundnr = "" Then
        Exit Sub
    End If
    
    If Not IsNumeric(cKundnr) Then
        Exit Sub
    End If
    
    cSQL = "Select * from Kunden where Kundnr = " & cKundnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!DS) Then
            If rsrs!DS = True Then
                
            Else
                Command4(22).ForeColor = vbRed
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DS_Unterschrieben"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SucheUnter(cKundnr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    
    SucheUnter = False
    
    If cKundnr = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(cKundnr) Then
        Exit Function
    End If
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    sSQL = "Select * from BONPAUSE where KdNr = '" & cKundnr & "'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        SucheUnter = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheUnter"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Nicht_erschienen_vorhanden(cKundnr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Nicht_erschienen_vorhanden = False
   
    If cKundnr = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(cKundnr) Then
        Exit Function
    End If
    
    cSQL = "Select * from TERMDEL where Kundnr = " & cKundnr & " "
    cSQL = cSQL & " and (Grund = 'nicht erschienen'"
    cSQL = cSQL & " or Grund = 'kurzfristig abgesagt')"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Nicht_erschienen_vorhanden = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Nicht_erschienen_vorhanden"
    Fehler.gsFehlertext = "Im Programmteil gelöschte Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub verfuegbar(dateDat As Date)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsBed As Recordset
    Dim rsBuch As Recordset
    Dim lBuchungsnr As Long
    Dim ibednu As Integer
    Dim lZeit1 As Long
    Dim cFeld As String
    Dim bFirst As Boolean
    Dim lFarbe As Long
    
    Dim ldynStartzeit As Long

    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim cDauer As String
    
    Dim lStartzeit As Long
    Dim lEndzeit As Long
    Dim lzeitblock As Long
    Dim cLBSatz As String
    
    cFeld = gcZeitBlock
    cFeld = SwapStr(cFeld, ":", "")
    lzeitblock = CLng(cFeld)
    
    
    List3.Visible = False
    List3.Nodes.Clear
    
    'hier den Feiertag regeln
    
    Dim bFeiertag As Boolean
    bFeiertag = False
    
    If IsThis_EinFeiertag(Format(dateDat, "DD.MM.YYYY")) Then
        bFeiertag = True
        Exit Sub
    End If
    
    
    
    Dim iWeekday As Integer
    iWeekday = Weekday(dateDat, vbMonday)
    
    Dim lMinIndex As Long
    Dim lMaxIndex As Long
    
    lMinIndex = ((iWeekday - 1) * 3) + 1
    lMaxIndex = ((iWeekday - 1) * 3) + 3
    
    Dim dUhrzeit As Double
    Dim sZeitblock As String
    sZeitblock = SwapStr(gcZeitBlock, ":", "")
    dUhrzeit = Val(sZeitblock) / 1440
    Dim dZeit As Double
    Dim dStartzeit As Double
    
    
    Dim k As Integer
    
    For k = lMinIndex To lMaxIndex
    
        If gZeiten(k).Von <> "" And gZeiten(k).Bis <> "" Then
    
            dZeit = TimeValue(gZeiten(k).Von)
            dStartzeit = dZeit - dUhrzeit
            
            cFeld = Format$(dStartzeit, "HH:MM") 'gcStartZeit
            cFeld = SwapStr(cFeld, ":", "")
            lStartzeit = CLng(cFeld)
            
            lStartzeit = lStartzeit + lzeitblock
                            
            If Right(CStr(lStartzeit), 2) = "60" Then
                lStartzeit = lStartzeit + 40
            End If
            
            
            
            cFeld = gZeiten(k).Bis 'gcEndeZeit
            cFeld = SwapStr(cFeld, ":", "")
            lEndzeit = CLng(cFeld)
            
            
            
            
            
            
            
            
            
            Screen.MousePointer = 11
        
            cSQL = "Select * from BEDTERM order by bednu asc "
            Set rsBed = gdBase.OpenRecordset(cSQL)
            If Not rsBed.EOF Then
                rsBed.MoveFirst
                Do While Not rsBed.EOF
                    If Not IsNull(rsBed!BEDNU) Then
                        ibednu = rsBed!BEDNU
                    Else
                        ibednu = 0
                    End If
                    
                    If Not IsNull(rsBed!FARBCODE) Then
                        lFarbe = rsBed!FARBCODE
                    Else
                        lFarbe = 0
                    End If
                    
                    ldynStartzeit = lStartzeit
                    
                    bFirst = True
                    
                    cSQL = "Select max(Buchungsnr) as maxBuch,min(uhrzeit) as mini from Termine where datum = " & CLng(dateDat)
                    cSQL = cSQL & " and bednu = " & ibednu
                    cSQL = cSQL & " group by Buchungsnr order by min(uhrzeit) asc"
                    Set rsBuch = gdBase.OpenRecordset(cSQL)
                    If Not rsBuch.EOF Then
                        rsBuch.MoveFirst
                        Do While Not rsBuch.EOF
                            If Not IsNull(rsBuch!maxBuch) Then
                                lBuchungsnr = rsBuch!maxBuch
                            Else
                                lBuchungsnr = 0
                            End If
        
                            lZeit1 = ermMinZeitperBuchung(lBuchungsnr)
                            
                            If lZeit1 > ldynStartzeit Then
                            
                                If bFirst = True Then
                                    cLBSatz = ermBEDbez(CLng(ibednu))
                                    
                                    List3.Nodes.Add Text:=cLBSatz
                                    List3.Nodes(List3.Nodes.Count).BackColor = FarbeBackColor(lFarbe)
                                    List3.Nodes(List3.Nodes.Count).ForeColor = FarbeForeColor(lFarbe)
                                    
                                    bFirst = False
                                End If
                                
                                
                                dStart = TimeValue(zeitanz(ldynStartzeit))
                                dEnde = TimeValue(zeitanz(lZeit1))
                        
                                dDauer = dEnde - dStart
                                cDauer = Format$(dDauer, "HH:MM")
                            
                                cLBSatz = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lZeit1) & " (" & cDauer & ")"
                                List3.Nodes.Add Text:=cLBSatz
                            End If
                            
                            ldynStartzeit = ermMaxZeitperBuchung(lBuchungsnr)
                            ldynStartzeit = ldynStartzeit + lzeitblock
                            
                            If Right(CStr(ldynStartzeit), 2) = "60" Then
                                ldynStartzeit = ldynStartzeit + 40
                            End If
                            
        
                        rsBuch.MoveNext
                        Loop
                    Else
                    
                        cLBSatz = ermBEDbez(CLng(ibednu))
                        List3.Nodes.Add Text:=cLBSatz
                        
                        
                        List3.Nodes(List3.Nodes.Count).BackColor = FarbeBackColor(lFarbe)
                        List3.Nodes(List3.Nodes.Count).ForeColor = FarbeForeColor(lFarbe)
                        
                        dStart = TimeValue(zeitanz(lStartzeit))
                        dEnde = TimeValue(zeitanz(lEndzeit))
        
                        dDauer = dEnde - dStart
                        cDauer = Format$(dDauer, "HH:MM")
                        
                        cLBSatz = "von " & zeitanz(lStartzeit) & " bis " & gZeiten(k).Bis & " (" & cDauer & ")"
                        List3.Nodes.Add Text:=cLBSatz
                        
        
                    End If
                    
                    If ldynStartzeit <> lStartzeit Then
                        If ldynStartzeit < lEndzeit Then
                            If bFirst = True Then
                                cLBSatz = ermBEDbez(CLng(ibednu))
                                
                                List3.Nodes.Add Text:=cLBSatz
                                List3.Nodes(List3.Nodes.Count).BackColor = FarbeBackColor(lFarbe)
                                List3.Nodes(List3.Nodes.Count).ForeColor = FarbeForeColor(lFarbe)
                                
                                bFirst = False
                            End If
                        
        
                            dStart = TimeValue(zeitanz(ldynStartzeit))
                            dEnde = TimeValue(zeitanz(lEndzeit))
        
                
                            dDauer = dEnde - dStart
                            cDauer = Format$(dDauer, "HH:MM")
                            
                            cLBSatz = "von " & zeitanz(ldynStartzeit) & " bis " & zeitanz(lEndzeit) & " (" & cDauer & ")"
                            List3.Nodes.Add Text:=cLBSatz
                            
                        End If
                    End If
                    
                    rsBuch.Close: Set rsBuch = Nothing
                    rsBed.MoveNext
                Loop
            End If
            rsBed.Close: Set rsBed = Nothing
        
        End If
    Next k
    
    List3.Refresh
    List3.Visible = True
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "verfuegbar"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next

End Sub
Private Function zeitanz(lZeit As Long) As String
    On Error GoTo LOKAL_ERROR
    
    
    Dim cFeld As String
    Dim cTeil1 As String
    Dim cTeil2 As String
    
    
    cFeld = CStr(lZeit)
    cTeil2 = Right(cFeld, 2)
    
    If Len(cFeld) = 3 Then
        cTeil1 = Left(cFeld, 1)
        zeitanz = "0" & cTeil1 & ":" & cTeil2
    ElseIf Len(cFeld) = 4 Then
        cTeil1 = Left(cFeld, 2)
        zeitanz = cTeil1 & ":" & cTeil2
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeitanz"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermMinZeitperBuchung(lBuch As Long) As Long
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lZeit1 As Long
    Dim cFeld As String
    
    ermMinZeitperBuchung = 0

    cSQL = "Select min(Uhrzeit) as MINI from Termine where Buchungsnr = " & lBuch
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!mini) Then
            cFeld = rsrs!mini
            cFeld = SwapStr(cFeld, ":", "")
            ermMinZeitperBuchung = CLng(cFeld)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermMinZeitperBuchung"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermMaxZeitperBuchung(lBuch As Long) As Long
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lZeit1 As Long
    Dim cFeld As String
    
    ermMaxZeitperBuchung = 0

    cSQL = "Select max(Uhrzeit) as MAX from Termine where Buchungsnr = " & lBuch
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Max) Then
            cFeld = rsrs!Max
            cFeld = SwapStr(cFeld, ":", "")
            ermMaxZeitperBuchung = CLng(cFeld)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermMaxZeitperBuchung"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function WirdDieKabineAngezeigt(cKab As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    WirdDieKabineAngezeigt = True
    
    cSQL = "Select * from PFLEGORT where bezeich = '" & cKab & "'"
    cSQL = cSQL & " and anzeigen = false "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        WirdDieKabineAngezeigt = False
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WirdDieKabineAngezeigt"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
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
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
End Sub



Private Sub Tree11_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo LOKAL_ERROR
    
    List11.Visible = True

    Dim i As Integer
    Dim ibednu  As Integer
    
    If Tree11.SelectedItem Is Nothing Then
        Exit Sub
    Else
        For i = 1 To Tree11.Nodes.Count
            If Tree11.Nodes(i).Selected = True Then
        
                ibednu = Tree11.Nodes(i).Tag
                zeige_Freie_Termine ibednu, Label2(21).Caption, Check2(0), Check2(1), Check2(2), Check2(3), Check2(4), Check2(5), Check2(6)
                
                Label2(22).Caption = ermfromBed("BEDNAME", CStr(ibednu))
                
                Exit For
            End If
        Next i
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tree11_NodeClick"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."


End Sub
Private Sub zeige_Freie_Termine(iBed As Integer, sTageszeit As String, chkMo As CheckBox, chkDi As CheckBox, _
chkMi As CheckBox, chkDo As CheckBox, chkFr As CheckBox, chkSa As CheckBox, chkSo As CheckBox)

    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    
    If iBed = 0 Then
        Exit Sub
    End If
    
'    WOCHENTAG,TAGESZEIT
    
    List11.Clear
   
    cSQL = "Select * from Freie_termine where bediener = " & iBed & " "
    
    Select Case sTageszeit
        Case "vormittags"
            cSQL = cSQL & " and Tageszeit = 'vormittag' "
        Case "nachmittags"
            cSQL = cSQL & " and Tageszeit = 'nachmittag' "
        Case "ganzer Tag"
        
    End Select
    
    If chkMo.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Montag' "
    End If
    
    If chkDi.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Dienstag' "
    End If
    
    If chkMi.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Mittwoch' "
    End If
    
    If chkDo.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Donnerstag' "
    End If
    
    If chkFr.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Freitag' "
    End If
    
    If chkSa.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Samstag' "
    End If
    
    If chkSo.value = vbUnchecked Then
        cSQL = cSQL & " and WOCHENTAG <> 'Sonntag' "
    End If
    
    
    cSQL = cSQL & " order by datum asc, startzeit asc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cSatz = ""
            If Not IsNull(rsrs!WOCHENTAG) Then
                cFeld = rsrs!WOCHENTAG
            End If
            cSatz = Left(cFeld, 2) & " "
'            cSatz = cFeld & Space(10 - Len(cFeld)) & " "
            
            If Not IsNull(rsrs!Datum) Then
                cFeld = rsrs!Datum
            End If
            
            Dim iTageZukunft As Integer
            iTageZukunft = DateDiff("d", DateValue(Now), CDate(cFeld))
            
            
            
            cSatz = cSatz & cFeld & Space(12 - Len(cFeld)) & " "
            
            If Not IsNull(rsrs!startzeit) Then
                cFeld = rsrs!startzeit
            End If
            
            cSatz = cSatz & cFeld & " - "
            
            If Not IsNull(rsrs!endezeit) Then
                cFeld = rsrs!endezeit
            End If
            cSatz = cSatz & cFeld & "  "
            
            If iTageZukunft = 1 Then
                cSatz = cSatz & " in " & iTageZukunft & " Tag"
            ElseIf iTageZukunft > 1 Then
                cSatz = cSatz & " in " & iTageZukunft & " Tagen"
            End If
        
            List11.AddItem cSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Label2(23).Caption = List11.ListCount
    
   
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Freie_Termine"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
