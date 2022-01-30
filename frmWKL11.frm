VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmWKL11 
   Appearance      =   0  '2D
   BackColor       =   &H00C0C000&
   Caption         =   "Stammdaten einlesen: KISS Format"
   ClientHeight    =   8595
   ClientLeft      =   1185
   ClientTop       =   1815
   ClientWidth     =   11880
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL11.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   975
      Left            =   9840
      TabIndex        =   57
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C000&
         Caption         =   "ab 20€ kalkulieren"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   4080
         TabIndex        =   143
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   4
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   141
         ToolTipText     =   "Aufschlag in Prozent auf den Listenverkaufspreis"
         Top             =   5160
         Width           =   735
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   10
         Left            =   1920
         TabIndex        =   63
         Top             =   5280
         Width           =   1815
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
         Caption         =   "Hinzufügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   62
         Top             =   4800
         Width           =   1335
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   240
         TabIndex        =   61
         Top             =   480
         Width           =   5415
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   60
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   59
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   58
         ToolTipText     =   "Kalender"
         Top             =   4800
         Width           =   360
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
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "automatisch anwenden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   4080
         TabIndex        =   144
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "LVK Aufschlag"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   142
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Kalkulation ausgeschlossen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   65
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Anzahl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   4440
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame4"
      Height          =   8355
      Left            =   120
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   13455
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "alte Artikel räumen"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   11
         Left            =   2880
         TabIndex        =   145
         Top             =   7400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   9
         Left            =   10560
         TabIndex        =   76
         Top             =   7560
         Width           =   1095
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
         Caption         =   "Kalk + Rund"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0C000&
         Caption         =   "ab 20€ kalkulieren"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   10440
         TabIndex        =   140
         Top             =   7840
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Listen-EK"
         Height          =   855
         Left            =   5400
         TabIndex        =   135
         Top             =   7440
         Width           =   1455
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   10
            TabIndex        =   136
            ToolTipText     =   "Aufschlagsfaktor auf den Listeneinkaufswert"
            Top             =   480
            Width           =   495
         End
         Begin sevCommand3.Command Command3 
            Height          =   300
            Index           =   13
            Left            =   960
            TabIndex        =   139
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
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
            Caption         =   "OK"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label22 
            Caption         =   "%"
            Height          =   255
            Index           =   4
            Left            =   640
            TabIndex        =   138
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "LEK Abschlag"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   137
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   6015
         Left            =   120
         TabIndex        =   113
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10610
         _Version        =   393216
         ForeColorSel    =   8454143
         Enabled         =   -1  'True
         FocusRect       =   0
         AllowUserResizing=   1
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   10
         Left            =   5760
         TabIndex        =   133
         Top             =   7080
         Value           =   1  'Aktiviert
         Width           =   1695
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   12
         Left            =   8400
         TabIndex        =   116
         Top             =   7560
         Width           =   855
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
         Caption         =   "Kalk"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "neuer KVK"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   6600
         Value           =   1  'Aktiviert
         Width           =   1575
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   94
         Top             =   8040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Beenden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   93
         Top             =   8040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Notizen"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   92
         Top             =   7080
         Value           =   1  'Aktiviert
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Mindestmenge(VPE)"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   91
         Top             =   6600
         Value           =   1  'Aktiviert
         Width           =   2415
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Top             =   8040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Mindestbestand"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   89
         Top             =   6840
         Value           =   1  'Aktiviert
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Neuheiten auf  ""GEFÜHRT"" "
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   88
         Top             =   6600
         Value           =   1  'Aktiviert
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "LVK"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   87
         Top             =   6840
         Value           =   1  'Aktiviert
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "EKPR"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   86
         Top             =   7080
         Value           =   1  'Aktiviert
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Linien"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   85
         Top             =   6840
         Value           =   1  'Aktiviert
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "AGN"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   84
         Top             =   7080
         Value           =   1  'Aktiviert
         Width           =   975
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   5
         Left            =   6960
         TabIndex        =   83
         Top             =   6720
         Width           =   735
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
         Caption         =   "l. Ziffern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox TxtRunden 
         Height          =   285
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   82
         Text            =   "9"
         Top             =   6720
         Width           =   375
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   6
         Left            =   8280
         TabIndex        =   81
         Top             =   6720
         Width           =   975
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
         Caption         =   "zurücksetzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   7
         Left            =   7680
         TabIndex        =   80
         Top             =   7080
         Width           =   495
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
         Caption         =   "Farbe?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   8
         Left            =   8280
         TabIndex        =   79
         Top             =   7080
         Width           =   975
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
         Caption         =   " entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "PGN"
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   78
         Top             =   6840
         Value           =   1  'Aktiviert
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   77
         ToolTipText     =   "Aufschlagsfaktor auf den Listeneinkaufswert"
         Top             =   7560
         Width           =   735
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   10
         Left            =   10800
         TabIndex        =   75
         Top             =   6720
         Width           =   495
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
         Caption         =   "Ausn."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   300
         Index           =   11
         Left            =   11400
         TabIndex        =   74
         Top             =   6720
         Width           =   255
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
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   3
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   73
         ToolTipText     =   "Aufschlag in Prozent auf den Listenverkaufspreis"
         Top             =   6960
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar pbrUebernahme 
         Height          =   225
         Left            =   120
         TabIndex        =   114
         Top             =   7755
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin sevCommand3.Command BTNSteuersenkung 
         Height          =   300
         Left            =   6960
         TabIndex        =   148
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
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
         Caption         =   "Steuersenkung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label4 
         BackColor       =   &H00004000&
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
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   2
         Left            =   10920
         TabIndex        =   112
         Top             =   8040
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00004000&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   10200
         TabIndex        =   111
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00004000&
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
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   9120
         TabIndex        =   110
         Top             =   8040
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Schritt 4: Übernahme der Artikel"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   120
         TabIndex        =   109
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   108
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   107
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   106
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   105
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   104
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   103
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   102
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label15 
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
         Left            =   6120
         TabIndex        =   101
         Top             =   0
         Width           =   5175
      End
      Begin VB.Label Label22 
         Caption         =   "LVK Aufschlag"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   100
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "LEK Faktor"
         Height          =   255
         Index           =   1
         Left            =   9360
         TabIndex        =   99
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   10560
         TabIndex        =   98
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  '2D
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   97
         Tag             =   "Shape"
         Top             =   7440
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Preisänderungen"
         Height          =   255
         Index           =   35
         Left            =   360
         TabIndex        =   96
         Top             =   7440
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   8280
      TabIndex        =   66
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   5415
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   68
         Top             =   4800
         Width           =   1815
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   67
         Top             =   5280
         Width           =   1815
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label17 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sperrliste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Anzahl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   4440
         Width           =   1695
      End
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H0000FFFF&
      Height          =   3375
      Left            =   600
      TabIndex        =   15
      Top             =   720
      Width           =   3015
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   118
         Top             =   2150
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "neue/alte Linien"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   46
         Top             =   1290
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "auto Kalkulation"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   1720
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Wochenänderung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   2580
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   0
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "alle Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   430
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "neue Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   860
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "nur Preisänderungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C0FF&
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
      Height          =   1275
      Left            =   240
      TabIndex        =   27
      Top             =   8160
      Visible         =   0   'False
      Width           =   2235
      Begin MSComctlLib.ProgressBar pbrlieferanten 
         Height          =   345
         Left            =   6120
         TabIndex        =   115
         Top             =   6840
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   7320
         TabIndex        =   52
         Top             =   7440
         Width           =   1215
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   1
         Left            =   8640
         TabIndex        =   51
         Top             =   7440
         Width           =   975
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
         Caption         =   "Sperren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   4
         Left            =   9720
         TabIndex        =   50
         Top             =   7440
         Width           =   1335
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
         Caption         =   "Sperrliste"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   7440
         Width           =   1215
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
         Caption         =   "Anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   6840
         TabIndex        =   48
         Top             =   8040
         Width           =   1935
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   3
         Left            =   9000
         TabIndex        =   39
         Top             =   7920
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6015
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10610
         _Version        =   393216
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   7920
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   2280
         TabIndex        =   29
         Top             =   7920
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "alle zurücksetzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   4440
         TabIndex        =   28
         Top             =   7920
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "alle auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label5 
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   6840
         Width           =   5895
      End
      Begin VB.Label Label16 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant für Stammdaten sperren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   55
         Top             =   7200
         Width           =   3495
      End
      Begin VB.Label Label19 
         BackColor       =   &H00008080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1440
         TabIndex        =   54
         Top             =   7440
         Width           =   5295
      End
      Begin VB.Label Label16 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sperrinfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   53
         Top             =   7800
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Schritt 3: Auswahl der Lieferanten"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   11775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   12135
      Begin VB.PictureBox picprogress 
         Height          =   255
         Left            =   6600
         ScaleHeight     =   195
         ScaleWidth      =   1635
         TabIndex        =   131
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtstatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   130
         Top             =   6120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   125
         Top             =   3720
         Width           =   375
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   124
         ToolTipText     =   "Klicken Sie hier, dann werden alle Artikelstammdaten in kürzester Zeit eingelesen"
         Top             =   6240
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
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
         Caption         =   "Lieferantendaten holen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   315
         Index           =   6
         Left            =   1920
         TabIndex        =   120
         Top             =   4680
         Width           =   315
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
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   6
         TabIndex        =   119
         Top             =   4680
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   5
         Left            =   9120
         TabIndex        =   117
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "V Proto"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox CG 
         Caption         =   "nur geführte Artikel anzeigen"
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
         Left            =   8280
         TabIndex        =   47
         Top             =   3120
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   2535
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   4
         Left            =   360
         TabIndex        =   45
         ToolTipText     =   "Klicken Sie hier, dann werden alle Artikelstammdaten in kürzester Zeit eingelesen"
         Top             =   3480
         Width           =   1920
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Wochendaten holen"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   "Klicken Sie hier, so gelangen Sie zur Bearbeitung einzelner Artikel"
         Top             =   3480
         Width           =   1680
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C000&
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
         ForeColor       =   &H00800000&
         Height          =   2295
         Left            =   5520
         TabIndex        =   13
         Top             =   720
         Width           =   5295
         Begin sevCommand3.Command Command8 
            Height          =   360
            Left            =   2160
            TabIndex        =   129
            ToolTipText     =   "alte Masterdatei kopieren"
            Top             =   1440
            Width           =   975
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
            Caption         =   "löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command6 
            Height          =   360
            Left            =   2160
            TabIndex        =   36
            ToolTipText     =   "alte Masterdatei kopieren"
            Top             =   360
            Width           =   975
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   ">>"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.FileListBox File2 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            Pattern         =   "M*!.*"
            TabIndex        =   35
            Top             =   360
            Width           =   1935
         End
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   3240
            Pattern         =   "m*!.*"
            TabIndex        =   14
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label24 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   123
            Top             =   1840
            Width           =   5055
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "neue Dateien"
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
            Left            =   3240
            TabIndex        =   44
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "alte Dateien"
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
            TabIndex        =   43
            Top             =   120
            Width           =   1935
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   2
         Left            =   4680
         TabIndex        =   5
         ToolTipText     =   "Klicken Sie hier, dann werden alle Artikelstammdaten in kürzester Zeit eingelesen"
         Top             =   3480
         Width           =   1920
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Schnellupdate"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   1
         Left            =   6600
         TabIndex        =   4
         Top             =   3480
         Width           =   1680
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Height          =   615
         Index           =   3
         Left            =   8280
         TabIndex        =   12
         Top             =   3480
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Protokoll anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   360
         Left            =   10320
         TabIndex        =   134
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
         Picture         =   "frmWKL11.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWKL11.frx":0AD4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   147
         Top             =   5310
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "neue Lieferanten (hier klicken)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   360
         MouseIcon       =   "frmWKL11.frx":0B95
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   146
         Top             =   5055
         Width           =   2415
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferantennummer"
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
         Left            =   360
         TabIndex        =   132
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   1
         Left            =   3000
         TabIndex        =   128
         Top             =   4440
         Width           =   7815
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Lizenz liegt vor"
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
         Left            =   3000
         TabIndex        =   127
         Top             =   4200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "KW"
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
         Left            =   2400
         TabIndex        =   126
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferanten Stammdaten"
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
         Left            =   360
         TabIndex        =   121
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2295
         Left            =   360
         TabIndex        =   41
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00008080&
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   3120
         Width           =   7695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "Schritt 1: Kopieren/Entpacken der Wochendatei"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   8655
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   735
      Left            =   10680
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
      Begin VB.ListBox List1 
         Enabled         =   0   'False
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
         TabIndex        =   23
         Top             =   240
         Width           =   11415
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
         Height          =   6984
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   11415
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   7680
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   1
         Left            =   2880
         TabIndex        =   20
         Top             =   7680
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00004040&
         Height          =   252
         Index           =   0
         Left            =   7920
         TabIndex        =   26
         Top             =   7800
         Width           =   972
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00004040&
         Height          =   252
         Index           =   1
         Left            =   8880
         TabIndex        =   25
         Top             =   7800
         Width           =   612
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00004040&
         Height          =   252
         Index           =   2
         Left            =   9480
         TabIndex        =   24
         Top             =   7800
         Width           =   972
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   3840
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   122
         Text            =   "frmWKL11.frx":0E9F
         Top             =   600
         Width           =   11535
      End
      Begin sevCommand3.Command Command3 
         Height          =   612
         Index           =   4
         Left            =   5640
         TabIndex        =   40
         Top             =   7800
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Beenden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComctlLib.ProgressBar pbrLinbez 
         Height          =   375
         Left            =   6120
         TabIndex        =   34
         Top             =   7200
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin sevCommand3.Command Command7 
         Height          =   612
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   7800
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   7800
         Width           =   2655
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lbl9 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Zu Ihrer Information drucken Sie sich diesen Text aus."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   7200
         Width           =   5895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Schritt 2: Einsehen/Ausdrucken der Infodatei"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   11775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Einlesen der Stammdaten-Änderungen"
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
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "frmWKL11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gbEinstell As Boolean
Dim sEinlesedat As String
Dim cWelche As String

Dim bduplisEAN As Boolean

Dim SpaltennummerArtnr  As Byte
Dim SpaltennummerBEZEICH  As Byte
Dim SpaltennummerLINR  As Byte
Dim SpaltennummerLPZ  As Byte
Dim SpaltennummerLIBESNR  As Byte
Dim SpaltennummerLEKPR  As Byte
Dim SpaltennummerVKPR  As Byte
Dim SpaltennummerKVKPR1  As Byte
Dim SpaltennummerMINBEST  As Byte
Dim SpaltennummerGEFUEHRT  As Byte
Dim SpaltennummerRABATT_OK  As Byte
Dim SpaltennummerPREISSCHU  As Byte
Dim SpaltennummerNOTIZEN  As Byte
Dim SpaltennummerAGN  As Byte
Dim SpaltennummerRKZ  As Byte
Dim SpaltennummerEAN  As Byte
Dim SpaltennummerMINMEN  As Byte
Dim SpaltennummerMWST  As Byte
Dim SpaltennummerMNOTIZEN  As Byte
Dim SpaltennummerKVKNEU As Byte
Dim SpaltennummerPGN  As Byte
Dim SpaltennummerKVKALT As Byte
Dim SpaltennummerAWM As Byte
Private Sub WKL11Positionieren()
    On Error GoTo LOKAL_ERROR

    With Frame0
        .Top = 100
        .Left = 100
        .Height = 8600
        .Width = 11775
    End With
    
    With Frame1
        .Top = 1020
        .Left = 240
        .Height = 6855 '6375
        .Width = 11775
    End With
    
    With Frame2
        .Top = 840
        .Left = 4800
        .Height = 5775
        .Width = 5895
        .BorderStyle = 0
    End With
    
    With Frame3
        .Top = 840
        .Left = 5800
        .Height = 5775
        .Width = 5895
        .BorderStyle = 0
    End With
   
    With Frame4
        .Top = 100
        .Left = 100
        .Height = 8600
        .Width = 11775
        .BorderStyle = 0
    End With

   With Frame6
        .Top = 960
        .Left = 5520
        .Height = 2412
        .Width = 5295
    End With
    
    With Frame7
        .Height = 3375
        .Left = 8280
        .Top = 5250
        .Width = 3015
    End With
    
    With Frame8
        .Top = 100
        .Left = 100
        .Height = 8775
        .Width = 11775
    End With
    
    With Frame9
        .Top = 100
        .Left = 100
        .Height = 8600
        .Width = 11775
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL11Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub BTNSteuersenkung_Click()
On Error GoTo LOKAL_ERROR
    
    
    Steuersenkung
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BTNSteuersenkung_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Steuersenkung()
On Error GoTo LOKAL_ERROR
    
    'Preisschutz, steht nur für kvkpr
'        Dim sSpalte As String = "VKPR"

        Screen.MousePointer = 11

        Dim dAltersteuer As Double
        Dim dNeuersteuer As Double

        Dim dAltersteuerErm As Double
        Dim dNeuersteuerErm As Double
        
        dAltersteuer = 119
        dNeuersteuer = 116

        dAltersteuerErm = 107
        dNeuersteuerErm = 105
        
        Dim sArtnr As String
        Dim sMWST As String
        Dim i As Integer
        Dim dVkPr As Double
        Dim lpreisszahler As Long
    
        MSFlexGrid2.Redraw = False
    
        lpreisszahler = 0
        
        MSFlexGrid2.Row = 0
        For i = 1 To MSFlexGrid2.Rows - 1
        
            MSFlexGrid2.Row = i
            MSFlexGrid2.Col = SpaltennummerArtnr
            sArtnr = MSFlexGrid2.Text
            
            If IsNumeric(sArtnr) Then
                        
                If BISTDUPreisschutz(sArtnr) = True Then
                
                    
                    MSFlexGrid2.Col = SpaltennummerMWST
                    sMWST = MSFlexGrid2.Text
                
                    MSFlexGrid2.Col = SpaltennummerVKPR
                    dVkPr = CDbl(MSFlexGrid2.Text)
                    
                    MSFlexGrid2.Col = SpaltennummerKVKNEU
                    If sMWST = "V" Then
                        MSFlexGrid2.Text = Format$(dVkPr * dNeuersteuer / dAltersteuer, "####.00")
                    ElseIf sMWST = "E" Then
                        MSFlexGrid2.Text = Format$(dVkPr * dNeuersteuerErm / dAltersteuerErm, "####.00")
                    End If
                Else
                    lpreisszahler = lpreisszahler + 1
                    MSFlexGrid2.Col = SpaltennummerKVKNEU
                    MSFlexGrid2.CellFontItalic = True
                    MSFlexGrid2.CellForeColor = vbRed
                    
                End If
            End If
            
        Next i
        
        MSFlexGrid2.Refresh
        
        MSFlexGrid2.Redraw = True
    
        If lpreisszahler > 0 Then
            anzeige "rot", lpreisszahler & " x Preisschutz ", Label22(2)
        Else
            anzeige "normal", "", Label22(2)
        End If
        
        
        

        Screen.MousePointer = 0
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Steuersenkung"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command0_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
    
        Case Is = 2
            Text2_KeyUp 1, vbKeyF2, 0
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "AWM"
    gstab = "MASTEMP"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad   As String
    Dim iRet    As Integer
    
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Screen.MousePointer = 11
    Label9.Caption = ""
    Label9.Refresh
    
    Select Case index
        Case Is = 0     '** weiter **
                If File1.ListCount = 0 Then
                    Screen.MousePointer = 0
                    Label6.Caption = "Keine Datei vorhanden. Bitte wählen Sie eine aus!"
                    Label6.Refresh
                    
                    If File2.ListCount > 0 Then
                        File2.Selected(0) = True
                        Label6.Caption = "Bitte wählen Sie eine Datei aus!"
                        Label6.Refresh
                    End If
                Else
                    gbEinstell = False
                    
                    sEinlesedat = File1.list(File1.ListIndex)
                    
                    schreibeProtokollStamda "______________________________________________________________"
                    schreibeProtokollStamda "Die Datei " & File1.list(File1.ListIndex) & " wird entpackt..."
                    
                    Label6.Caption = "Die Datei " & File1.list(File1.ListIndex) & " wird entpackt..."
                    Label6.Refresh
                    
                    
                    If DeKompStada = False Then
                        DeKompMasterDateiWKL11
                        
                        If Not NewTableSuchenDBKombi("Master", gdBase) Then
                            Label6.Caption = "Das Entpacken der Datei ist gescheitert. Versuchen Sie es nochmal!"
                            Label6.Refresh
                            Screen.MousePointer = 0
                            File1.Pattern = "M*!.*"
                            File1.Refresh
                            Exit Sub
                        End If
                    Else
                        Check1(11).Visible = True
                    End If
                    
                    Label6.Caption = " "
                    Label6.Refresh
                    
                    ZeigeInfoDateiWKL11
                End If
        Case Is = 1
            Unload frmWKL11
        Case Is = 2     '**Schnellupdate **
            iRet = MsgBox("Möchten Sie wirklich das Schnellupdate durchführen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbNo Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If File1.ListCount = 0 Then
                Screen.MousePointer = 0
                Label6.Caption = "Keine Datei vorhanden. Bitte wählen Sie eine aus!"
                Label6.Refresh
                
                If File2.ListCount > 0 Then
                    File2.Selected(0) = True
                    Label6.Caption = "Bitte wählen Sie eine Datei aus!"
                    Label6.Refresh
                End If
            Else
                gbEinstell = True
                sEinlesedat = File1.list(File1.ListIndex)
                
                Label6.Caption = "Die Datei " & File1.list(File1.ListIndex) & " wird entpackt..."
                Label6.Refresh
                
                If DeKompStada = False Then
                    DeKompMasterDateiWKL11
                    
                    If Not NewTableSuchenDBKombi("Master", gdBase) Then
                        Label6.Caption = "Das Entpacken der Datei ist gescheitert. Versuchen Sie es nochmal!"
                        Label6.Refresh
                        Screen.MousePointer = 0
                        File1.Pattern = "M*!.*"
                        File1.Refresh
                        Exit Sub
                    End If
                End If
                    
                   
                schreibeSUProtokoll "Datei: " & sEinlesedat & " über Schnellupdate eingelesen"

                Label6.Caption = " "
                Label6.Refresh
        
                ZeigeInfoDateiWKL11
                
                Command7_Click 0
                Command2_Click 0
                Command3_Click 0
    
            End If
            
        Case Is = 3    '** vergangenes Protokol zeigen **
        
            If Not NewTableSuchenDBKombi("Stadapro", gdBase) Then
                anzeige "rot", "Keine Protokolle vorhanden", Label9
            Else
                If Datendrin("Stadapro", gdBase) = True Then
                    Frame7.Visible = True: CG.Visible = True: Label26(1).Visible = False
                Else
                    anzeige "rot", "Keine Protokolle vorhanden", Label9
                End If
            End If
        Case 4    'Datenholen per FTP
        
            giWochendat = 0
            If IsNumeric(Text3(1).Text) Then
                giWochendat = Text3(1).Text
            End If
            
            Dim bmerke As Boolean
            bmerke = gbFTPautomatic
            gbFTPautomatic = True
            
            giKissFtpMode = 3 ' FTPMODE= 3 , Programmupdates/ Stammdaten holen aus WKL11 Stammdaten einlesen
            frmWKL38.Show 1
            
            gbFTPautomatic = bmerke
            
            File1.Pattern = "M*!.*"
            File1.Refresh
        Case 5
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "STAMDA.txt"
            Screen.MousePointer = 0
        Case 6
            Screen.MousePointer = 0
            Text3_KeyUp vbKeyF2, 0, 0
'        Case 7
'            Screen.MousePointer = 0
'            Text3_KeyUp vbKeyF2, 0, 0
            
        Case 8   'Datenholen per FTP Wochendatei
            If Text3(0).Text <> "" Then
                If IsNumeric(Text3(0).Text) Then
                    If Val(Text3(0).Text) > 0 Then
                        
                        glLiNr = Val(Text3(0).Text)
                        
                        If glLiNr > 199999 Then
                        
                            'alte Lieferantendaten löschen!
                            Kill cPfad & "M" & glLiNr & "!.111"
                            
                            giKissFtpMode = 17
                            frmWKL38.Show 1
                            
                        End If
                    End If
                End If
            End If
            
            File1.Pattern = "M*!.*"
            File1.Refresh
            
            File2.Refresh
            
            Text3(0).Text = ""
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ZeigeInfoDateiWKL11()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr       As Integer
    Dim cdatei        As String
    Dim cPfad         As String
    Dim lStart        As Long
    Dim cZeile        As String
    Dim cZeichen      As String
    Dim bKeinLinie    As Boolean
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "IN\"
    iFileNr = FreeFile
    
    Open cPfad & "MASTER.TXT" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cdatei = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cdatei
        Close iFileNr
    Else
        Close iFileNr
        Kill cPfad & "MASTER.TXT"
        bKeinLinie = True
        GoTo weiter
    End If
    
    Text4.Font = "Courier New"
    Text4.Text = cdatei
    
weiter:
If bKeinLinie = False Then
    Frame1.Visible = False
    Frame7.Visible = False
    Frame9.Visible = True
Else
    If NewTableSuchenDBKombi("MLINBEZ", gdBase) Then
        If Datendrin("MLINBEZ", gdBase) Then
            AlleLinBezCSV
        Else
            AlleLinBezNehmenWKL11
        End If
    Else
        AlleLinBezNehmenWKL11
    End If
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeInfoDateiWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub FormatiereGrid1(lAnzSatz As Long)
On Error GoTo LOKAL_ERROR

        With MSFlexGrid1
            .Cols = 9
            .Rows = lAnzSatz + 1
            .Row = 0
            .Col = 0
            .ColWidth(0) = 320
            .Text = "OK"
            .Col = 1
            .ColWidth(1) = 800
            .Text = "Lieferant"
            .Col = 2
            .ColWidth(2) = 3300
            .Text = "Lieferantenbezeichnung"
            .Col = 3
            .ColWidth(3) = 3300
            .Text = "Strasse"
            .Col = 4
            .ColWidth(4) = 600
            .Text = "PLZ"
            .Col = 5
            .ColWidth(5) = 2500
            .Text = "Stadt"
            .Col = 6
            .ColWidth(6) = 1500
            .Text = "Telefon"
            .Col = 7
            .ColWidth(7) = 1500
            .Text = "Fax"
            .Col = 8
            .ColWidth(8) = 3000
            .Text = "Lieferantenname"
        End With
        
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereGrid1"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EinzelnLisrtNehmenWKL11()
    On Error GoTo LOKAL_ERROR
    
    Dim lcountg     As Long
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    With Frame0

        .Visible = True
    End With
    
    cSQL = "Select * from MLISRT where linr not in (Select linr from DELSTADAL) order by liefBEZ"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzSatz = rsrs.RecordCount
        FormatiereGrid1 lAnzSatz
        MSFlexGrid1.Visible = False
        rsrs.MoveFirst
        lAktSatz = 0
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            With MSFlexGrid1
                .Row = lAktSatz
                .RowHeight(lAktSatz) = 270
                .Col = 0
                .Text = "X"
            End With
            
            If Not IsNull(rsrs!linr) Then
                ctmp = rsrs!linr
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = ctmp
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                ctmp = rsrs!LIEFBEZ
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = ctmp
            
            If Not IsNull(rsrs!strasse) Then
                ctmp = rsrs!strasse
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = ctmp
            
            If Not IsNull(rsrs!Plz) Then
                ctmp = rsrs!Plz
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = ctmp
        
            If Not IsNull(rsrs!STADT) Then
                ctmp = rsrs!STADT
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = ctmp
        
            If Not IsNull(rsrs!Tel) Then
                ctmp = rsrs!Tel
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = ctmp
        
            If Not IsNull(rsrs!Fax) Then
                ctmp = rsrs!Fax
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 7
            MSFlexGrid1.Text = ctmp
        
            If Not IsNull(rsrs!LINAME) Then
                ctmp = rsrs!LINAME
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 8
            MSFlexGrid1.Text = ctmp
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame9.Visible = False
    Label5.Caption = "Es sind alle Lieferanten ausgewählt."
    Label5.Refresh
    
    cSQL = "Select * from MLISRT where linr  in (Select linr from DELSTADAL) "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcountg = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lcountg = 1 Then
        anzeige "normal", lcountg & " gesperrter Lieferant wird nicht angezeigt", Label19
    Else
        anzeige "normal", lcountg & " gesperrte Lieferanten werden nicht angezeigt", Label19
    End If
    MSFlexGrid1.Redraw = True
    
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    MSFlexGrid1.Visible = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EinzelnLisrtNehmenWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
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
Private Sub KopiereDateiWKL11()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet    As Long
    Dim lfail   As Long
    Dim cdatei  As String
    Dim cPfad   As String
    Dim cQuelle As String
    Dim cZiel   As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cdatei = File2.list(File2.ListIndex)
    
    cQuelle = cPfad & cdatei
    cZiel = cPfad & "IN\" & cdatei
    
    lRet = CopyFile(cQuelle, cZiel, lfail)
    If lRet = 0 Then
        MsgBox "Datei " & cdatei & " konnte nicht kopiert werden!"
        
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KopiereDateiWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheUrAltMasterWKL11()
    On Error GoTo LOKAL_ERROR
    
    
    Dim lAnz        As Long
    Dim lcount      As Long
    Dim lHeute      As Long
    Dim lDateiDatum As Long
    Dim cdatei      As String
    Dim cPfad       As String
    
    lHeute = Fix(Now)
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    lAnz = File2.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = File2.list(lcount)
        lDateiDatum = FileDateTime(cPfad & cdatei)
        If lHeute - lDateiDatum > 49 Then
            schreibeProtokollStamda "Datei: " & cdatei & " wurde gelöscht(zu alt)"
            Kill cPfad & cdatei
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheUrAltMasterWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function MerkeMarkierteLiNrWKL11() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim lcount   As Long
    Dim ctmp     As String
    
    MerkeMarkierteLiNrWKL11 = True
    lAnzSatz = MSFlexGrid1.Rows
    
    lcount = 0
    glLiNr = 0
    
    For lAktSatz = 1 To lAnzSatz - 1
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Row = lAktSatz
        ctmp = MSFlexGrid1.Text
        If ctmp = "X" Then
            lcount = lcount + 1
            MSFlexGrid1.Col = 1
            ctmp = MSFlexGrid1.Text
            glLiNr = glLiNr + 1
            gclinr11(glLiNr) = ctmp
            schreibeProtokollStamda "Lieferant: " & ctmp & " wurde gewählt"
        End If
    Next lAktSatz
    
    glAnzLiNr = lcount
                                                                      
    If MoveMarkierteLiNrWKL11 = False Then
        MerkeMarkierteLiNrWKL11 = False
    End If

Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MerkeMarkierteLiNrWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function MoveMarkierteLiNrWKL11() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    Dim cSQL    As String
    Dim rsrs1   As Recordset
    Dim rsRs2   As Recordset
    
    MoveMarkierteLiNrWKL11 = True
    
    If glLiNr = 0 Then
        Exit Function
    End If
    
    Label5.Caption = "Artikeldaten werden erstellt..."
    Label5.Refresh
    
    For lcount = 1 To glLiNr
        cSQL = "Select * from MLISRT where LINR = " & gclinr11(lcount)
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        
        cSQL = "Select * from LISRT where LINR = " & gclinr11(lcount)
        Set rsRs2 = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not rsRs2.EOF Then
                rsRs2.Edit
                rsRs2!SYNStatus = "E"
            Else
                rsRs2.AddNew
                rsRs2!SYNStatus = "A"
            End If
            
            rsRs2!linr = rsrs1!linr
            rsRs2!LIEFBEZ = rsrs1!LIEFBEZ
            rsRs2!strasse = rsrs1!strasse
            rsRs2!Plz = rsrs1!Plz
            rsRs2!STADT = rsrs1!STADT
            rsRs2!LASTDATE = DateValue(Now)
            rsRs2!LASTTIME = TimeValue(Now)
            rsRs2.Update
        End If
    Next lcount
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
    rsrs1.Close: Set rsrs1 = Nothing
    If LeseMasterWKL11 = False Then
        MoveMarkierteLiNrWKL11 = False
    End If
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveMarkierteLiNrWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function LeseMasterWKL11() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    Dim k           As Integer
    Dim sSQL        As String
    Dim ctmp        As String

    LeseMasterWKL11 = True
    'Grid formatieren
    Tabcheck "MASTEMP"
    FormatGridOverTablay "MASTEMP"

    With MSFlexGrid2
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
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With
    'Daten ermitteln
   
    SucheBildschirmdaten
    
    Label5.Caption = "Tabelle wird gefüllt..."
    Label5.Refresh
    
    'Grid fuellen
    GridFuellen
    
    Me.Refresh
    
    Label5.Caption = "Tabellenbreite anpassen..."
    Label5.Refresh
    

    Tabellenbreiteanpassen MSFlexGrid2, 1.25 * gdTabfak
    
    
    
    'Etiketten löschen
    
    Label5.Caption = "Etiketten aktualisieren..."
    Label5.Refresh
    
    etidruleeren
    ermittlespalten
    
    Label5.Caption = "Artikelfarben hinzufügen..."
    Label5.Refresh
    
    FaerbenGrid MSFlexGrid2, CInt(SpaltennummerAWM), CInt(SpaltennummerAWM)
    
    Label5.Caption = "Prüfe auf EAN - Duplikate..."
    Label5.Refresh
    
    'wenn esüdro dann keine eanduplis
    'dann ean aus vorhandenen artikel löschen
    Dim lcount As Long
    
    Me.Refresh
    
    For lcount = 1 To glLiNr
         If gclinr11(lcount) = 100000 Then
            checkthislinrOfDuplis MSFlexGrid2, SpaltennummerEAN, SpaltennummerArtnr

            Exit For
         End If
    Next lcount
    
    Me.Refresh
    
    
    
    'hier beginnt die neue EAN - Duplikatsbehandlung
    'wenn Duplikate vorhanden sind, dann folgendes:
    'Bestand und KVK auffangen
    
    Behandle_EAN_Duplis MSFlexGrid2, SpaltennummerEAN, SpaltennummerArtnr
    
    

    Label5.Caption = "unterschiedliche Preise darstellen..."
    Label5.Refresh
    
    Faerbewegenpreisunter MSFlexGrid2, SpaltennummerKVKALT, SpaltennummerKVKNEU
    
    
    Label5.Caption = "Voreinstellungen laden..."
    Label5.Refresh
    
    If NewTableSuchenDBKombi("STAMDAE", gdBase) Then
        lastvoreinstellungzeigen "STAMDAE", frmWKL11, 11
    End If
    
    If gbKVKSicher = True Then
        Check1(0).value = vbUnchecked
    End If

Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseMasterWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub etidruleeren()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from ETIDRU where artnr in (Select artnr from MASTEMP) "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "etidruleeren"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheBildschirmdaten()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim counter         As Long
    Dim cPfad           As String
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim dNettospanne    As Double
    Dim dEK             As Double
    Dim cMWST           As String
    Dim cNewKassenPr    As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If Not SpalteInTabellegefundenNEW("MASTER", "PGN", gdBase) Then
        cSQL = " Alter table MASTER add PGN BYTE "
        gdBase.Execute cSQL, dbFailOnError

        cSQL = "Update MASTER Set PGN = 0"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Create Index LINR on MASTER (LINR)"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create Index ARTNR on MASTER (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "MASTEMP", gdBase
    CreateTable "MASTEMP", gdBase
    
    pbrlieferanten.Max = 20
    pbrlieferanten.Visible = True
    
    counter = 0
    For lcount = 1 To glLiNr
    
        If counter = 20 Then
            counter = 0
        End If
        counter = counter + 1
        pbrlieferanten.value = counter

        cSQL = "Insert into MASTEMP SELECT "
        cSQL = cSQL & "  MASTER.ARTNR "
        cSQL = cSQL & ", MASTER.BEZEICH "
        cSQL = cSQL & ", MASTER.LINR "
        cSQL = cSQL & ", MASTER.LPZ "
        cSQL = cSQL & ", MASTER.LIBESNR "
        cSQL = cSQL & ", MASTER.LEKPR "
        cSQL = cSQL & ", MASTER.VKPR "
        cSQL = cSQL & ", MASTER.VKPR as KVKPR1"
        cSQL = cSQL & ", MASTER.AGN "
        cSQL = cSQL & ", MASTER.PGN "
        cSQL = cSQL & ", MASTER.RKZ "
        cSQL = cSQL & ", MASTER.EAN "
        cSQL = cSQL & ", MASTER.MINMEN "
        cSQL = cSQL & ", MASTER.MWST "
        cSQL = cSQL & ", MASTER.NOTIZEN as MNOTIZEN "
        cSQL = cSQL & ", '98' as AWM "
        cSQL = cSQL & ", 'N' as ETIMERK "
        cSQL = cSQL & ", 0 as SPANNE "
        cSQL = cSQL & ", 'N' as SHOP "
        cSQL = cSQL & " from MASTER "
        cSQL = cSQL & " Where MASTER.LINR = " & gclinr11(lcount)
        gdBase.Execute cSQL, dbFailOnError
    Next lcount
    
    cSQL = "Update MASTEMP inner join INTERART on "
    cSQL = cSQL & " MASTEMP.artnr = INTERART.artnr "
    cSQL = cSQL & " set MASTEMP.SHOP = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Update MASTEMP left join Artikel on "
    cSQL = cSQL & " ARTIKEL.ARTNR = MASTEMP.ARTNR "
    cSQL = cSQL & "Set "
    cSQL = cSQL & " MASTEMP.KVKALT = ARTIKEL.KVKPR1 "
    cSQL = cSQL & ", MASTEMP.LVKALT = ARTIKEL.VKPR "
    cSQL = cSQL & ", MASTEMP.MINBEST = ARTIKEL.MINBEST "
    cSQL = cSQL & ", MASTEMP.GEFUEHRT = ARTIKEL.GEFUEHRT "
    cSQL = cSQL & ", MASTEMP.RABATT_OK = ARTIKEL.RABATT_OK "
    cSQL = cSQL & ", MASTEMP.PREISSCHU = ARTIKEL.PREISSCHU "
    cSQL = cSQL & ", MASTEMP.NOTIZEN = ARTIKEL.NOTIZEN "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = " Update MASTEMP inner join Artikel on "
    cSQL = cSQL & " ARTIKEL.ARTNR = MASTEMP.ARTNR "
    cSQL = cSQL & "Set "
    cSQL = cSQL & " MASTEMP.AWM = ARTIKEL.AWM "
    gdBase.Execute cSQL, dbFailOnError
    
    'ab Hier check der Autokalkulierung
    '1. Unter Voreinstellung basierend auf LEK
    If gsSpanne = "LEK" Then
        '2. Etimerk = J
        '3. Not Preisschu
        '4. Feld Nettospanne gefüllt
        '5. LEK gefüllt
        
        cSQL = " Update MASTEMP inner join Artikel on "
        cSQL = cSQL & " ARTIKEL.ARTNR = MASTEMP.ARTNR "
        cSQL = cSQL & "Set "
        cSQL = cSQL & " MASTEMP.PREISSCHU = ARTIKEL.PREISSCHU "
        cSQL = cSQL & " where ARTIKEL.PREISSCHU = 'N'"
        cSQL = cSQL & " and MASTEMP.LEKPR > 0 "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = " Update MASTEMP inner join ARTLIEF on "
        cSQL = cSQL & " ARTLIEF.ARTNR = MASTEMP.ARTNR and ARTLIEF.LINR = MASTEMP.LINR"
        cSQL = cSQL & " Set "
        cSQL = cSQL & " MASTEMP.SPANNE = ARTLIEF.SPANNE "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update MASTEMP "
        cSQL = cSQL & " Set ETIMERK = 'J' "
        cSQL = cSQL & " where MASTEMP.SPANNE > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update MASTEMP set DIFFVKWERT =  lVKALT- kVKALT  "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update MASTEMP set DIFFVK = 100 * DIFFVKWERT/lVKALT"
        cSQL = cSQL & " where lVKALT <> 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update MASTEMP set DIFFKVKWERT = KVKPR1 - kvkalt"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = " Update MASTEMP set DIFFKVK = 100 * DIFFKVKWERT/kvkalt"
        cSQL = cSQL & " where kvkalt <> 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select * from Mastemp where ETIMERK = 'J' and PREISSCHU = 'N' and (not MASTEMP.Spanne is null and MASTEMP.Spanne <> 0 ) "
        Set rsrs = gdBase.OpenRecordset(cSQL)

        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                'Wir brauchen den LEK aus Mastemp
                'MWST
                'Nettospanne
                If Not IsNull(rsrs!lekpr) Then
                    dEK = rsrs!lekpr
                Else
                    dEK = 0
                End If

                If Not IsNull(rsrs!MWST) Then
                    cMWST = rsrs!MWST
                Else
                    cMWST = "V"
                End If


                If Not IsNull(rsrs!SPANNE) Then
                    dNettospanne = rsrs!SPANNE
                Else
                    dNettospanne = 0
                End If
                cNewKassenPr = Runden(CDbl(fnVKneuNS(dEK, cMWST, dNettospanne)))

                rsrs.Edit
                rsrs!KVKPR1 = cNewKassenPr
                rsrs!AWM = "97"
                rsrs.Update

                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 53 Or err.Number = 3376 Or err.Number = 3375 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SucheBildschirmdaten"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub GridFuellen()
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
    Dim sSQL        As String
    Dim lMax        As Long
    
    sSQL = "Select * from Mastemp order by Linr,lpz"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    With MSFlexGrid2
    .Redraw = False
    
    pbrlieferanten.Visible = True
'    pbrlieferanten.Max = 1000
    counter = 0
    
    lrow = 1
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        lMax = rsrs.RecordCount
        pbrlieferanten.Max = lMax
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
'            If counter = lMax / 2 Then
'                counter = 0
'            End If
            counter = counter + 1
            pbrlieferanten.value = counter
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "Listen - VK", "KVK alt", "KVK neu", "LVK alt"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            
                        Case Is = "Diff KVK", "Diff VK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00") & " %"
                            
                        Case Is = "Preisschutz", "Geführt"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "N"
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                        Case Is = "Artikelbezeichnung"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                If gbTagAkt = True Then
                                    sWert = UCase(rsrs(sSpaltenbez(i)))
                                Else
                                    sWert = rsrs(sSpaltenbez(i))
                                End If
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                        Case Is = "Rabatt"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "J"
                            End If
                            .Row = lrow
                            .Text = sWert
                         
                        Case Is = "MinBest"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert
                        
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
                                
            If Not IsNull(rsrs!AWM) Then
                sWert = rsrs!AWM
                If Trim(sWert) = "98" Then
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbWhite
                        .CellForeColor = &HFF&
                    Next j
                ElseIf Trim(sWert) = "97" Then
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbYellow
                        .CellForeColor = vbBlue
                    Next j
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.5
    Next i
        
    rsrs.Close: Set rsrs = Nothing
    pbrlieferanten.Visible = False
    If byAnzahlSpalten < 2 Then
    Else
        .FixedCols = 1
    End If
    .RowHeight(1) = 0
    lrow = lrow - 1
    Label4(2).Caption = lrow

    .Redraw = True
    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cLin As String
    
    Dim cLinr As String
    

    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0     '** Übernehmen **
            
            Label5.Caption = ""
            Label5.Refresh
            Me.Refresh
            If MerkeMarkierteLiNrWKL11 = False Then
                Frame4.Visible = False
                
                Frame0.Visible = False
                Frame1.Visible = True
                
                File1.Refresh
                
                anzeige "rot", "Abbruch - keine Daten eingelesen!", Label9
                schreibeProtokollStamda "Abbruch - keine Daten eingelesen!"
                
            Else
            
                If glAnzLiNr = 0 Then
                    Frame0.Visible = True
                    Label5.Caption = "Sie haben keinen Lieferanten ausgewählt."
                    Label5.Refresh
                Else
                    
                    Frame0.Visible = False
                    Frame4.Visible = False
                    Frame9.Visible = False
                        
                    Label4(0).Caption = "1"
                    Command3(0).Enabled = True
                    
                    If Text2(4).Text <> "" Then
                        Text2(3).Text = Text2(4).Text
                        Check6.value = Check2.value
                        Command3_Click 9
                    End If
                    
                    If glAnzLiNr = 1 Then
                    
                        cLinr = gclinr11(1)
                        If Val(cLinr) > 0 Then
                            Text2(2).Text = ermLEK_ABSCHLAG_Lief(CLng(cLinr))
                            
                        End If
                    
                    End If
                    
                    
                    Frame4.Visible = True
                    Me.Refresh
                    
                End If
            End If
            
        Case Is = 1    '** sperren **
            
            Frame2.Visible = False
            If Text1(0).Text <> "" Then
                If IsNumeric(Text1(0).Text) Then
                    Sperrlinr Text1(0).Text, Text1(1).Text
                    
                End If
            End If
            Text1(0).Text = ""
            Text1(1).Text = ""
            
            EinzelnLisrtNehmenWKL11
            
        Case 2      '** Alle zurücksetzen **
            Label5.Caption = "Sie haben keinen Lieferanten ausgewählt."
            Label5.Refresh
            SchalteLieferantenWKL11 index
            
        Case 3      '** Alle auswählen **
            Label5.Caption = "Sie haben alle Lieferanten ausgewählt."
            Label5.Refresh
            SchalteLieferantenWKL11 index
            
        Case 4
            Screen.MousePointer = 0
            cWelche = "alle"
            ZeigSperrListe cWelche
        Case 5
            Screen.MousePointer = 0

            cLin = Trim(Left(List4.list(List4.ListIndex), 6))
            If IsNumeric(cLin) Then
                DELinSperr cLin
                ZeigSperrListe cWelche
            End If
        Case 6
            Screen.MousePointer = 0
            Frame2.Visible = False
            EinzelnLisrtNehmenWKL11
        Case 7
            Screen.MousePointer = 0
            cWelche = "nur"
            ZeigSperrListe cWelche
        Case 8
            Screen.MousePointer = 0
            voreinstellungspeichernE11C
            Frame3.Visible = False
        Case 9
            Screen.MousePointer = 0

            cLin = Trim(Left(List5.list(List5.ListIndex), 6))
            If IsNumeric(cLin) Then
                KalkinSperr cLin
                ZeigNOKALKL
            End If
        Case 10
            If Text2(1).Text <> "" Then
                If IsNumeric(Text2(1).Text) Then
                    INNOKALKL Text2(1).Text
                    
                End If
            End If
            Text2(1).Text = ""
            ZeigNOKALKL
    End Select
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Sperrlinr(cLinr As String, cBeschr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not SpalteInTabellegefundenNEW("DELSTADAL", "BESCHR", gdBase) Then
        SpalteAnfuegenNEW "DELSTADAL", "BESCHR", "Text(30)", gdBase
    End If
    
    sSQL = "Delete from DELSTADAL where LINR = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DELSTADAL (LINR,BESCHR) values ( " & cLinr & ",'" & cBeschr & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Sperrlinr"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub INNOKALKL(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from NOKALKL where LINR = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into NOKALKL select LINR,LIEFBEZ from lisrt where LINR = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "INNOKALKL"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub DELinSperr(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from DELSTADAL where LINR = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DELinSperr"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub KalkinSperr(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from NOKALKL where LINR = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KalkinSperr"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ZeigSperrListe(sWelche As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lcount As Long
    
    If Not SpalteInTabellegefundenNEW("DELSTADAL", "BESCHR", gdBase) Then
        SpalteAnfuegenNEW "DELSTADAL", "BESCHR", "Text(30)", gdBase
    End If
    
    List4.Clear
    lcount = 0
    
    If sWelche = "alle" Then
    
        sSQL = "Select * from DELSTADAL order by linr  "
    
    Else
        sSQL = "Select * from DELSTADAL where linr in (Select linr from MLISRT) order by linr  "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                If Not IsNull(rsrs!BESCHR) Then
                
                    List4.AddItem rsrs!linr & Space(8 - Len(rsrs!linr)) & rsrs!BESCHR
                Else
                    List4.AddItem rsrs!linr
                    
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lcount = 1 Then
        anzeige "normal", lcount & " Lieferant", Label18
    Else
        anzeige "normal", lcount & " Lieferanten", Label18
    End If
    Frame2.BackColor = glH2
    Frame2.Visible = True
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigSperrListe"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ZeigNOKALKL()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lcount As Long
    
    
    
    List5.Clear
    lcount = 0
    
    sSQL = "Select * from NOKALKL order by linr  "
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                If Not IsNull(rsrs!LIEFBEZ) Then
                
                    List5.AddItem rsrs!linr & Space(8 - Len(rsrs!linr)) & rsrs!LIEFBEZ
                Else
                    List5.AddItem rsrs!linr
                    
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    If lcount = 1 Then
        anzeige "normal", lcount & " Lieferant", Label20
    Else
        anzeige "normal", lcount & " Lieferanten", Label20
    End If
    
    Check2.BackColor = glH2
    Frame3.BackColor = glH2
    Frame3.Visible = True
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigNOKALKL"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub SchalteLieferantenWKL11(iSchaltung As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    
    lRows = MSFlexGrid1.Rows
    lRows = lRows - 1
    lcol = 0
    
    For lrow = 1 To lRows
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.Col = lcol
        If iSchaltung = 2 Then
            MSFlexGrid1.Text = ""
        End If
        If iSchaltung = 3 Then
            MSFlexGrid1.Text = "X"
        End If
    Next lrow
    
    With MSFlexGrid1
        .Row = 1
        .Col = 0
        .SetFocus
    End With
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchalteLieferantenWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "ARTNR"
            SpaltennummerArtnr = i
            Case Is = "BEZEICH"
            SpaltennummerBEZEICH = i
            Case Is = "LINR"
            SpaltennummerLINR = i
            Case Is = "LPZ"
            SpaltennummerLPZ = i
            Case Is = "LIBESNR"
            SpaltennummerLIBESNR = i
            Case Is = "LEKPR"
            SpaltennummerLEKPR = i
            Case Is = "VKPR"
            SpaltennummerVKPR = i
            Case Is = "KVKPR1"
            SpaltennummerKVKPR1 = i
            Case Is = "MINBEST"
            SpaltennummerMINBEST = i
            Case Is = "GEFUEHRT"
            SpaltennummerGEFUEHRT = i
            Case Is = "RABATT_OK"
            SpaltennummerRABATT_OK = i
            Case Is = "PREISSCHU"
            SpaltennummerPREISSCHU = i
            Case Is = "NOTIZEN"
            SpaltennummerNOTIZEN = i
            Case Is = "PGN"
            SpaltennummerPGN = i
            Case Is = "AGN"
            SpaltennummerAGN = i
            Case Is = "RKZ"
            SpaltennummerRKZ = i
            Case Is = "EAN"
            SpaltennummerEAN = i
            Case Is = "MINMEN"
            SpaltennummerMINMEN = i
            Case Is = "MWST"
            SpaltennummerMWST = i
            Case Is = "MNOTIZEN"
            SpaltennummerMNOTIZEN = i
            Case Is = "AWM"
            SpaltennummerAWM = i
        End Select
        Select Case UCase$(sSpaltenname(i))
            Case Is = "KVK NEU"
            SpaltennummerKVKNEU = i
            Case Is = "KVK ALT"
            SpaltennummerKVKALT = i
        End Select
    Next i
    
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbenGrid(grid As MSFlexGrid, iawmSpalte As Integer, Izufarbspalte As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    
    Dim cAWM                As String
    
    With grid
        .Redraw = False
    
        For i = 0 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
            .Col = j
                If .Col = iawmSpalte Then
                    cAWM = .TextMatrix(i, j)
                    If cAWM = "" Then cAWM = "0"
                    FaerbenFlex cAWM, grid, Izufarbspalte, i
                End If
                
            Next j
        Next i
        .Redraw = True
    
        
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FaerbenGrid"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbeBestvor(gridx As MSFlexGrid, spalteEAN As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim sArtnr      As String
    Dim sEAN        As String
    Dim counter     As Long
    Dim lMax        As Long
    
    
    pbrlieferanten.Visible = True
'    pbrlieferanten.Max = 1000
    counter = 0
    
    With gridx
        .Redraw = False
        lMax = .Rows
        pbrlieferanten.Max = lMax
        For j = 1 To .Rows - 1
        
'            If counter = 1000 Then
'                counter = 0
'            End If
            counter = counter + 1
            pbrlieferanten.value = counter
        
        
            .Row = j
            .Col = spalteartnr
            sArtnr = .Text
            
            .Col = spalteEAN
            sEAN = .Text
            
            If checkthisean(sEAN, sArtnr) = False Then
                If checkthiseankoml(sArtnr) = True Then
                    bduplisEAN = True
                    dupliEANSstada sEAN, sArtnr
                    .Col = spalteartnr
                    .CellBackColor = &H80000012
                    .CellForeColor = vbRed
                Else
                    .Col = spalteartnr
                    .CellBackColor = glfarbe(0)
                    .CellForeColor = vbBlack
                    .Text = ""
                End If
            Else
                .Col = spalteartnr
                .CellBackColor = glfarbe(0)
                .CellForeColor = vbBlack
            End If
            
        Next j
        .Redraw = True
    
    End With
    
    pbrlieferanten.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeBestvor"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Behandle_EAN_Duplis(gridx As MSFlexGrid, spalteEAN As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim sArtnr      As String
    Dim sEAN        As String
    Dim counter     As Long
    Dim lMax        As Long
    Dim cSQL        As String
    
    loeschNEW "ALTINFO" & srechnertab, gdBase
    
    cSQL = "Create Table ALTINFO" & srechnertab
    cSQL = cSQL & " ( "
    cSQL = cSQL & " altartnr int"
    cSQL = cSQL & ", EAN varchar(13)"
    cSQL = cSQL & ", neuArtnr int"
    cSQL = cSQL & ", Bestand int"
    cSQL = cSQL & ", KVKPR1 real "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    pbrlieferanten.Visible = True
    counter = 0
    
    With gridx
        .Redraw = False
        lMax = .Rows
        pbrlieferanten.Max = lMax
        For j = 1 To .Rows - 1
        

            counter = counter + 1
            pbrlieferanten.value = counter
        
        
            .Row = j
            .Col = spalteartnr
            sArtnr = .Text
            
            .Col = spalteEAN
            sEAN = .Text
            
            If checkthisean(sEAN, sArtnr) = False Then
                'ja, ein Duplikat
                
                'Altinformation speichern
                Altinformation_speichern sEAN, sArtnr
                
            Else

            End If
            
        Next j
        .Redraw = True
    
    End With
    pbrlieferanten.Visible = False
    
    
'    If Datendrin("ALTINFO" & srechnertab, gdBase) Then
'
'        'Bestand updaten
'
'        'KVKPR1 updaten
'
'        Dim lAltArtnr As Long
'        Dim rsRS As dao.Recordset
'        Dim sSQL As String
'
'        sSQL = "Select altartnr from Altinfo" & srechnertab
'        Set rsRS = gdBase.OpenRecordset(sSQL)
'        If Not rsRS.EOF Then
'            rsRS.MoveFirst
'            Do While Not rsRS.EOF
'
'                lAltArtnr = 0
'                If Not IsNull(rsRS!altartnr) Then
'                    lAltArtnr = rsRS!altartnr
'                End If
'
'                If lAltArtnr > 0 Then
'                    'EAN bei altartnr entfernen
'                    EAN_Updaten lAltArtnr
'
'                End If
'
'                rsRS.MoveNext
'            Loop
'
'        End If
'        rsRS.Close: Set rsRS = Nothing
'
'    End If
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Behandle_EAN_Duplis"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkthislinrOfDuplis(gridx As MSFlexGrid, spalteEAN As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim sArtnr      As String
    Dim sEAN        As String
    Dim counter     As Long
    Dim lMax        As Long
    
    pbrlieferanten.Visible = True
'    pbrlieferanten.Max = 1000
    counter = 0
    
    With gridx
        .Redraw = False
        lMax = .Rows
        pbrlieferanten.Max = lMax
        
        For j = 1 To .Rows - 1
        
'            If counter = 1000 Then
'                counter = 0
'            End If
            counter = counter + 1
            pbrlieferanten.value = counter
        
        
            .Row = j
            .Col = spalteartnr
            sArtnr = .Text
            
            .Col = spalteEAN
            sEAN = .Text
            
            If checkthisean(sEAN, sArtnr) = False Then
                dupliEANloesch sEAN, sArtnr
            End If
            
        Next j
        .Redraw = True
    
    End With
    
    pbrlieferanten.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "checkthislinrOfDuplis"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lesenEinstellungenKeinDel()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    giKeinDel = -1
    
    checkFFE
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!KeinDel) Then
                If rsrs!KeinDel = True Then
                    giKeinDel = 0
                Else
                    giKeinDel = -1
                End If
            Else
                giKeinDel = -1
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "lesenEinstellungenKeinDel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Function ArtikeluebernahmeMaster1() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lcount       As Long
    Dim lAnzRows     As Long
    Dim lAktFeld     As Long
    Dim lUebernommen As Long
    Dim counter      As Long
    
    Dim dLEKPR      As Double
    Dim dKVkPr1     As Double
    Dim dKVkPr1Neu  As Double
    Dim dVkPrAlt    As Double
    Dim dVkPrNeu    As Double
    Dim dKVkPrAlt   As Double
    Dim dKVkPrNeu   As Double
    Dim dKarVKPr    As Double

    Dim cPfad       As String
    Dim cSQL        As String

    Dim cKVkPr1     As String
    Dim cKVkPr1NEU  As String
    Dim ctmp        As String
    Dim cFeld       As String
    Dim cWert       As String
    Dim cArtNr      As String
    Dim cLinr       As String
    Dim cLEKPR      As String
    Dim cLiBesNr    As String
    Dim cMinMen     As String
    Dim cMinBest    As String
    Dim cGefuehrt   As String
    Dim cRabatt_OK  As String
    Dim cPreiSchutz As String
    Dim iRet         As Integer
    Dim bPreisschutz As Boolean
    Dim rsrs1        As Recordset
    Dim rsRs2        As Recordset
    Dim rsRs3        As Recordset
    Dim rsStadapro   As Recordset
    
    Dim iStufe          As Integer
    iStufe = 0
    
    ArtikeluebernahmeMaster1 = False
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If Not gbEinstell Then 'Bei Schnellupdate keine Nachfrage
        iRet = MsgBox("Sind alle Übernahme-Schalter korrekt gesetzt?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet <> vbYes Then
            Screen.MousePointer = 0
            Exit Function
        End If
        schreibeProtokollStamda "Artikelübernahme im Standardeinlesemodus"
    Else
        schreibeProtokollStamda "Artikelübernahme im Schnelleinlesemodus"
    End If
    
    iStufe = 1
     
    'welche Haken sind gesetzt
    For lcount = 0 To 10
        If Check1(lcount).value = vbChecked Then
            gbTransfer(lcount + 1) = True
        Else
            gbTransfer(lcount + 1) = False
            schreibeProtokollStamda "Dieser Übernahmeschalter wurde nicht gesetzt: " & Check1(lcount).Caption
        End If
    Next
    
    iStufe = 2
    
    Dim gbPause As Boolean
    
    If BistDualleineinderDatenbank Then
        gbPause = False
    Else
        gbPause = True
    End If
    
    iStufe = 3
    
    schreibeProtokollStamda "Artikelübernahme beginnt..."
    
    loeschNEW "STADAPRO", gdBase 'stadapro ist neu und das ganze gprotom-geschreibe nur für dbase - Kunden
    CreateTable "STADAPRO", gdBase
    
    iStufe = 4
    Set rsStadapro = gdBase.OpenRecordset("STADAPRO", dbOpenTable)
    
    pbrUebernahme.Max = 50
    pbrUebernahme.Visible = True
    
    MSFlexGrid2.Redraw = False

    
    lAnzRows = MSFlexGrid2.Rows
    lAnzRows = lAnzRows - 1
    
    
    
    Dim cLinrRaeum As String
    
    If Check1(11).value = vbChecked Then
    
        cLinrRaeum = "0"
    
        cSQL = "Select distinct(linr) from Master "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
            
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            
            If Not IsNull(rsrs1!linr) Then
                cLinrRaeum = rsrs1!linr
            End If
        End If
        
        rsrs1.Close: Set rsrs1 = Nothing
        
        If Val(cLinrRaeum) > 0 Then
            cSQL = "Update Artlief set RKZ = 'J'"
            cSQL = cSQL & ", EXDAT = '" & DateValue(Now) & "' "
            cSQL = cSQL & " where LINR = " & cLinrRaeum
            gdBase.Execute cSQL, dbFailOnError
        End If
        
    
    End If
    
    
    lesenEinstellungenKeinDel
    
    
    
    
    schreibeProtokollStamda lAnzRows - 1 & " Artikel werden jetzt eingelesen"
    
    iStufe = 5
    
    For lcount = 2 To lAnzRows  'Ab hier wird das ganze Grid abgeklappert
nochmal:
        rsStadapro.AddNew 'Protokoll schreiben ein neuer Datensatz
        rsStadapro!Quelldat = sEinlesedat
        rsStadapro!Datum = DateValue(Now)
        
        MSFlexGrid2.Row = lcount
            
        If counter = 50 Then
            counter = 0
        Else
            counter = counter + 1
            pbrUebernahme.value = counter
        End If
        
        iStufe = 100
        
        Label4(0).Caption = lcount
        Label4(0).Refresh
        
        MSFlexGrid2.Col = SpaltennummerArtnr 'Artikelnummer holen
        cArtNr = MSFlexGrid2.Text
        
        'Artikelnummer Check
        
        If Trim$(Left(UCase$(cArtNr), 1)) = "X" Or cArtNr = "" Or cArtNr = "entfernt" Then
            If lcount = lAnzRows Then
                Exit For
            ElseIf lcount < lAnzRows Then
                lcount = lcount + 1
                GoTo nochmal
            End If
        End If
        
        
        iStufe = 101
        
        MSFlexGrid2.Col = SpaltennummerLINR
        cLinr = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerLIBESNR
        cLiBesNr = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerLEKPR
        cLEKPR = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerKVKPR1
        cKVkPr1 = MSFlexGrid2.Text
        cKVkPr1 = fnMoveComma2Point$(cKVkPr1)
        dKVkPr1 = Val(cKVkPr1)
        
        MSFlexGrid2.Col = SpaltennummerKVKNEU
        cKVkPr1NEU = MSFlexGrid2.Text
        cKVkPr1NEU = fnMoveComma2Point$(cKVkPr1NEU)
        dKVkPr1Neu = Val(cKVkPr1NEU)
        
        MSFlexGrid2.Col = SpaltennummerMINMEN
        cMinMen = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerMINBEST
        cMinBest = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerGEFUEHRT
        cGefuehrt = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerRABATT_OK
        cRabatt_OK = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerPREISSCHU
        cPreiSchutz = MSFlexGrid2.Text
        
        MSFlexGrid2.Col = SpaltennummerVKPR
        If MSFlexGrid2.Text <> "" Then
            dKarVKPr = MSFlexGrid2.Text
        Else
            dKarVKPr = 0
        End If
        
        iStufe = 102
        
        cSQL = "Select * from MASTER where ARTNR = " & cArtNr
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
            
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            
            If gdStadaPause > 0 Then
                If gbPause Then
                    PauseSi (CSng(gdStadaPause))
                    Me.Refresh
                End If
            End If
            
            'ist der Artikel als Neu gekennzeichnet
            
            If Not IsNull(rsrs1!FLAG) Then
                If rsrs1!FLAG = "N" Then
                    If giKeinDel = -1 Then
                        LoescheArtikelSofort cArtNr, "bei Stadaübernahme gelöscht."
                    End If
                End If
            End If
            
            
            
            
            iStufe = 103
    
            cSQL = "Select * from ARTIKEL where ARTNR = " & cArtNr
            Set rsRs2 = gdBase.OpenRecordset(cSQL)
            iStufe = 1031
            If Not rsRs2.EOF Then
            
                If Not IsNull(rsRs2!PREISSCHU) Then
                    If rsRs2!PREISSCHU = "J" Then
                        bPreisschutz = True
                    Else
                        bPreisschutz = False
                    End If
                Else
                    bPreisschutz = False
                End If
                
                iStufe = 1032
                
                rsRs2.Edit 'ab hier Editmodus in der Artikel
                rsRs2!SYNStatus = "E"
                rsStadapro!Akt = "Änderung"
                
                rsRs2!RABATT_OK = cRabatt_OK
                rsRs2!PREISSCHU = cPreiSchutz
                rsRs2!GEFUEHRT = cGefuehrt
                
                iStufe = 1033
                If Not IsNull(rsRs2!AWM) Then
                    ctmp = rsRs2!AWM
                Else
                    ctmp = "0"
                End If
                rsRs2!AWM = ctmp
                rsStadapro!FARBNR = Val(ctmp)
                rsStadapro!GEFUEHRT = cGefuehrt
                
                iStufe = 1034
                If gbTransfer(10) = True Then
                    If gbTagAkt = True Then
                        rsRs2!BEZEICH = UCase(rsrs1!BEZEICH)
                    Else
                        rsRs2!BEZEICH = rsrs1!BEZEICH
                    End If
                End If
                
                If gbTransfer(3) = True Then
                    rsRs2!MINBEST = cMinBest
                Else
                
                    'man lässt es einfach
'                    rsRs2!MINBEST = 0
                End If
                If Not IsNull(rsRs2!GEFUEHRT) Then
                    ctmp = rsRs2!GEFUEHRT
                Else
                    ctmp = ""
                End If
                
                iStufe = 1035
                
                If ctmp = "J" Then
                    rsRs2!LASTDATE = DateValue(Now)
                    rsRs2!LASTTIME = TimeValue(Now)
                End If
            Else
                iStufe = 1036
                
                Sicherheitslöschen cArtNr 'artlief
                iStufe = 1037
                rsRs2.AddNew    'ab AddnewModus in der Artikel
                rsRs2!AUFDAT = DateValue(Now)
                rsRs2!SYNStatus = "A"
                rsStadapro!Akt = "Neuheit"
                
                iStufe = 1038
                If gbTransfer(5) = True Then
                    rsRs2!GEFUEHRT = "J"
                Else
                    rsRs2!GEFUEHRT = "N" 'cGefuehrt
                    rsStadapro!GEFUEHRT = rsRs2!GEFUEHRT
                End If
                
                If gbTransfer(3) = True Then
                    rsRs2!MINBEST = cMinBest
                Else
                    rsRs2!MINBEST = 0
                End If
                
                If gbTagAkt = True Then
                    rsRs2!BEZEICH = UCase(rsrs1!BEZEICH)
                Else
                    rsRs2!BEZEICH = rsrs1!BEZEICH
                End If
                iStufe = 1039
                
                rsRs2!RABATT_OK = cRabatt_OK
                rsRs2!BONUS_OK = "J"
                rsRs2!UMS_OK = "J"
                rsRs2!PREISSCHU = cPreiSchutz
                rsRs2!LASTDATE = DateValue(Now)
                rsRs2!LASTTIME = TimeValue(Now)
                rsRs2!ekpr = rsrs1!lekpr
                rsRs2!ETIMERK = "N"
                rsRs2!AWM = "98"
                rsStadapro!FARBNR = 98
                
            End If
            
            iStufe = 104
            
            cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLinr
            Set rsRs3 = gdBase.OpenRecordset(cSQL)
            
            If Not IsNull(rsrs1!artnr) Then
                rsStadapro!artnr = rsrs1!artnr
            Else
                rsStadapro!artnr = ""
            End If
            rsRs2!artnr = rsrs1!artnr
                         
            If Not IsNull(rsrs1!BEZEICH) Then
                rsStadapro!BEZEICH = rsrs1!BEZEICH
            Else
                rsStadapro!BEZEICH = ""
            End If
            
            
                  
            iStufe = 105
            
            If rsRs2!linr <> rsrs1!linr Then
                cSQL = "Delete from Artlief where artnr = " & rsrs1!artnr & " "
                cSQL = cSQL & " and LINR = " & rsRs2!linr & " "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            If Not IsNull(rsrs1!linr) Then
                rsRs2!linr = rsrs1!linr
            End If
                              
            If gbTransfer(8) = True Then
                rsRs2!LPZ = rsrs1!LPZ
            Else
                rsRs2!LPZ = rsRs2!LPZ
            End If
            rsRs2!LIBESNR = rsrs1!LIBESNR
                               
            If Not IsNull(rsRs2!lekpr) Then
                dLEKPR = rsRs2!lekpr
            Else
                dLEKPR = 0
            End If
            cWert = Format$(dLEKPR, "####0.00")
            rsStadapro!LEK_ALT = cWert

            If Not IsNull(rsrs1!lekpr) Then
                dLEKPR = rsrs1!lekpr
            Else
                dLEKPR = 0
            End If
            cWert = Format$(dLEKPR, "####0.00")

            If gbTransfer(7) = True Then
                rsStadapro!LEK_NEW = cWert
                rsRs2!lekpr = rsrs1!lekpr
            Else
                rsStadapro!LEK_NEW = rsStadapro!LEK_ALT
                rsRs2!lekpr = rsRs2!lekpr
            End If
                   
            If (gbTransfer(1) = True) And Not bPreisschutz Then
                If rsRs2!KVKPR1 <> dKVkPr1Neu Then
                    If MSFlexGrid2.CellBackColor = vbYellow Then
                        rsStadapro!AUTOKALK = True
                    End If
                    
                    '** Bei Preisänderung ETIDRU füllen **
                    If rsStadapro!Akt = "Änderung" Then
                        rsStadapro!Akt = "Preisänderung"
                    End If
                    
                    If Not bPreisschutz Then
                        SchreibeEtiDruWKL11 rsrs1, rsRs2, dKVkPr1Neu, cArtNr
                    End If
                    
                End If
            End If
            
            iStufe = 106
                        
            If Not IsNull(rsRs2!vkpr) Then
                dVkPrAlt = rsRs2!vkpr
            Else
                dVkPrAlt = 0
            End If
            cWert = Format$(dVkPrAlt, "####0.00")
            rsStadapro!VKPR_ALT = cWert

            If Not IsNull(rsRs2!KVKPR1) Then
                dKVkPrAlt = rsRs2!KVKPR1
            Else
                dKVkPrAlt = 0
            End If
            cWert = Format$(dKVkPrAlt, "####0.00")
            rsStadapro!KVK_ALT = cWert

            If Not IsNull(rsrs1!vkpr) Then
                dVkPrNeu = rsrs1!vkpr
            Else
                dVkPrNeu = 0
            End If
            
            cWert = Format$(dKarVKPr, "######0.00")
            
            If gbTransfer(6) = False Then
                rsStadapro!VKPR_NEW = cWert
            Else
                cWert = Format$(dVkPrNeu, "######0.00")
                rsStadapro!VKPR_NEW = cWert
                rsRs2!vkpr = dKarVKPr
            End If
            
            If rsStadapro!Akt = "Neuheit" Then
                rsRs2!KVKPR1 = dKVkPr1Neu
                rsRs2!vkpr = dKarVKPr
                cWert = Format$(dKVkPr1Neu, "####0.00")
                rsStadapro!KVK_NEW = cWert
            Else
                If (gbTransfer(1) = True) And Not bPreisschutz Then
                    rsRs2!KVKPR1 = dKVkPr1Neu
                    cWert = Format$(dKVkPr1Neu, "####0.00")
                    rsStadapro!KVK_NEW = cWert
                Else
                    ctmp = Format$(dKVkPrAlt, "####0.00")
                    rsStadapro!KVK_NEW = ctmp
                End If
            End If
            
            iStufe = 107
            
            If gbTransfer(9) = True Then
                rsRs2!AGN = rsrs1!AGN
            Else
                rsRs2!AGN = rsRs2!AGN
            End If
            
            If gbTransfer(10) = True Then
                rsRs2!PGN = rsrs1!PGN
            Else
                rsRs2!PGN = rsRs2!PGN
            End If
            
            iStufe = 108
            
            If rsrs1!EAN <> rsRs2!EAN Then
                If Not IsNull(rsRs2!EAN2) Then
                    rsRs2!EAN3 = rsRs2!EAN2
                End If
                rsRs2!EAN2 = rsRs2!EAN
            End If
            rsRs2!EAN = rsrs1!EAN
            
            If gbTransfer(2) = True Then rsRs2!MINMEN = rsrs1!MINMEN
            If rsStadapro!Akt = "Neuheit" And (gbTransfer(2) = False) Then rsRs2!MINMEN = 1
        
            iStufe = 109
            
            rsRs2!MWST = rsrs1!MWST
        
            If gbTransfer(4) = True Or (rsStadapro!Akt = "Neuheit") Then
                rsRs2!NOTIZEN = rsrs1!NOTIZEN
            End If
        
            rsRs2!INHALT = rsrs1!INHALT
            rsRs2!INHALTBEZ = rsrs1!INHALTBEZ
            rsRs2!GRUNDPREIS = rsrs1!GRUNDPREIS

            rsRs2.Update
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            
            iStufe = 110
            
            If rsRs3.EOF Then
                rsRs3.AddNew
                rsRs3!SYNStatus = "A"
                rsRs3!MINMEN = Val(cMinMen)
            Else
                rsRs3.Edit
                rsRs3!SYNStatus = "E"
                
                If gbTransfer(2) = True Then
                    rsRs3!MINMEN = Val(cMinMen)
                Else
                    
                End If
            End If
            
            iStufe = 1101
            
            rsRs3!artnr = Val(cArtNr)
            rsRs3!linr = Val(cLinr)
            rsRs3!LIBESNR = cLiBesNr
            cLEKPR = fnMoveComma2Point$(cLEKPR)
            rsRs3!lekpr = Val(cLEKPR)
            
            'RKZ check
            rsRs3!RKZ = rsrs1!RKZ

            If rsRs3!RKZ = "J" Then
                If Not IsNull(rsRs3!EXDAT) Then
                    If CLng(rsRs3!EXDAT) = 0 Then
                        rsRs3!EXDAT = DateValue(Now)
                    End If
                Else
                    rsRs3!EXDAT = DateValue(Now)
                End If
            Else
                rsRs3!EXDAT = 0
            End If
            
            rsRs3.Update
            rsRs3.Close: Set rsRs3 = Nothing
            
            
        End If
        iStufe = 1109
        rsStadapro.Update
    Next lcount
    
    rsStadapro.Close
    
    
    
    Dim sSQL As String
    If NewTableSuchenDBKombi("ARTEAN2", gdBase) Then
        sSQL = "Update Artikel inner join ARTEAN2 on Artikel.artnr = ARTEAN2.artnr "
        sSQL = sSQL & " set Artikel.EAN2 = ARTEAN2.EAN2 "
        gdBase.Execute sSQL, dbFailOnError
        
        loeschNEW "ARTEAN2", gdBase
    End If
    
    If NewTableSuchenDBKombi("ARTEAN3", gdBase) Then
        sSQL = "Update Artikel inner join ARTEAN3 on Artikel.artnr = ARTEAN3.artnr "
        sSQL = sSQL & " set Artikel.EAN3 = ARTEAN3.EAN3 "
        gdBase.Execute sSQL, dbFailOnError
        
        loeschNEW "ARTEAN3", gdBase
    End If
    
    
    
    Dim rsrs As DAO.Recordset
    
    If NewTableSuchenDBKombi("ALTINFO" & srechnertab, gdBase) Then
    
        If Datendrin("ALTINFO" & srechnertab, gdBase) Then
        
            'eans raus bei alt artikel
            'Bestand und KVKpr von alt artikel übernehmen
            
            sSQL = "Update Artikel inner join ALTINFO" & srechnertab & " on Artikel.artnr = ALTINFO" & srechnertab & ".neuartnr "
            sSQL = sSQL & " set Artikel.Bestand = ALTINFO" & srechnertab & ".Bestand "
            sSQL = sSQL & " , Artikel.KVKPR1 = ALTINFO" & srechnertab & ".KVKPR1 "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update Artikel inner join ALTINFO" & srechnertab & " on Artikel.artnr = ALTINFO" & srechnertab & ".altartnr "
            sSQL = sSQL & " set Artikel.EAN ='' "
            sSQL = sSQL & " , Artikel.EAN2 = '' "
            sSQL = sSQL & " , Artikel.EAN3 = '' "
            sSQL = sSQL & " , Artikel.Bestand = 0 "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update artean_k inner join ALTINFO" & srechnertab & " on artean_k.artnr = ALTINFO" & srechnertab & ".altartnr "
            sSQL = sSQL & " set artean_k.EAN ='' "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Delete * from artean_k where artean_k.EAN ='' "
            gdBase.Execute sSQL, dbFailOnError
            
            
            'Historiendaten umschreiben?
        
        End If
    End If
    
    
    
    
    iStufe = 111
    BringFarbeInsSpiel "STADAPRO", gdBase
    
    iStufe = 112
    schreibeProtokollStamda "Fertig Artikel wurden erfolgreich eingelesen"
    pbrUebernahme.Visible = False
   
    iStufe = 113
    rsrs1.Close: Set rsrs1 = Nothing
    
    loesch "Master"
    iStufe = 114
    Kill cPfad & "MASTER.MDX"
    
    iStufe = 115
    ArtikeluebernahmeMaster2
    iStufe = 116
    
    ArtikeluebernahmeMaster1 = True
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ArtikeluebernahmeMaster1"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten. " & iStufe

        Fehlermeldung1
        Resume Next
    End If
End Function
Private Sub ArtikeluebernahmeMaster2()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim cPfad2      As String
    Dim cSQL        As String
    Dim cLinr       As String
    Dim rsrs1        As Recordset
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    'hier Großlieferantendaten einlesen
    'jetzt die Datei MASTER2.DBF mit den Zweit-EKPR einlesen
    '***** ist die Datei überhaupt vorhanden ? *****
    
    If tableSuchenDBKombi("Master2", 1) = True Then
    
        cSQL = "Select distinct(LINR) from Master2"
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            If Not IsNull(rsrs1!linr) Then
                cLinr = Trim(rsrs1!linr)
            Else
                cLinr = ""
            End If
        End If
        
        rsrs1.Close: Set rsrs1 = Nothing
        
        pbrUebernahme.Max = 50
        pbrUebernahme.Visible = True
        
                                   'alle REWE-Daten löschen
        cSQL = "Delete from ARTLIEF where LINR = " & cLinr  'variabel
        gdBase.Execute cSQL, dbFailOnError

        pbrUebernahme.value = 10
        'Tabelle ARTLIEF packen
        loeschNEW "TEMPXXXX", gdBase
        cSQL = "Select * into TEMPXXXX from ARTLIEF"
        gdBase.Execute cSQL, dbFailOnError
        
        pbrUebernahme.value = 20
        loeschNEW "ARTLIEF", gdBase
        
        pbrUebernahme.value = 30
        cSQL = "Select * into ARTLIEF from TEMPXXXX"
        gdBase.Execute cSQL, dbFailOnError
        
        pbrUebernahme.value = 40
        loeschNEW "TEMPXXXX", gdBase

        cSQL = "Insert into artlief Select Artnr,Linr,LEKPR,LIBESNR,MINMENGE as MINMEN, 'A' as SYNSTATUS from Master2  "
        gdBase.Execute cSQL, dbFailOnError
         
        cSQL = "Create Index LINR on ARTLIEF (LINR)"
        gdBase.Execute cSQL, dbFailOnError

        pbrUebernahme.value = 50
        cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
        gdBase.Execute cSQL, dbFailOnError

        pbrUebernahme.value = 10
        cSQL = "Create Index ARTLINR on ARTLIEF (ARTNR, LINR)"
        gdBase.Execute cSQL, dbFailOnError

        pbrUebernahme.value = 20
        cSQL = "Create Index LIBESNR on ARTLIEF (LIBESNR)"
        gdBase.Execute cSQL, dbFailOnError

        pbrUebernahme.value = 30
        'Originäre REWE-Artikel wieder nach ARTLIEF stellen!
                        
        cSQL = "Insert into ARTLIEF Select "
        cSQL = cSQL & " ARTNR, LINR, LIBESNR, LEKPR, MINMEN, 'E' as SYNSTATUS from ARTIKEL "
        cSQL = cSQL & "where LINR = " & cLinr
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    loesch "MASTER2"
    Kill cPfad & "MASTER2.MDX"

    
    pbrUebernahme.value = 40
    pbrUebernahme.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ArtikeluebernahmeMaster2"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

        Fehlermeldung1
        
    End If
End Sub
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String

    Screen.MousePointer = 11

    Select Case index
        Case Is = 0
            Me.Refresh
        
            If ArtikeluebernahmeMaster1 = False Then
                Exit Sub
            End If
            
            voreinstellungspeichernE11C
            
            Frame4.Visible = False
            Frame9.Visible = False
            
            

            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            
            Kill cPfad & "IN\MASTER!.000"
            
            File1.Pattern = "M*!.*"
            File1.Refresh
            
            'also
            File2.Refresh
            
            '********Zeige frame1
            If File1.ListCount > 0 Then
                File1.Selected(0) = True
                
                Label6.Caption = "Bitte wählen Sie eine Datei aus!"
                Label6.Refresh
            Else
                Label6.Caption = "Bitte wählen Sie eine Datei aus!"
                Label6.Refresh
            End If
            
            Label9.Visible = True
            Label9.Caption = "...Schritt 1 erledigt..." & vbCrLf & "...Schritt 2 erledigt..." & vbCrLf & "...Schritt 3 erledigt..." & vbCrLf & "...Schritt 4 erledigt..." & vbCrLf & "..........fertig.........."
            Label9.Refresh
            
            Frame1.Visible = True
           
        Case Is = 1  '** Zurück **
            Frame4.Visible = False
            Frame0.Visible = True
            
        Case Is = 2  '** Beenden **
            schreibeProtokollStamda "Stammdaten einlesen wurde beendet"
            Unload frmWKL11
        Case Is = 3  '** Beenden **
            schreibeProtokollStamda "Stammdaten einlesen wurde beendet"
            Unload frmWKL11
        Case Is = 4  '** Beenden **
            schreibeProtokollStamda "Stammdaten einlesen wurde beendet"
            Unload frmWKL11
        Case Is = 5  '** Runden **
            Rundenwkl11
        Case Is = 6  '** zurücksetzen **
            MerkeMarkierteLiNrWKL11
        Case 7
            Screen.MousePointer = 0
            frmWKL49.Show 1
        Case 8 'Artikel entfernen
        
            If MSFlexGrid2.RowSel > 1 Then
                FlexGrid_Update MSFlexGrid2
                
            Else
                
            End If
            
            Dim lrow As Long
            Dim lcol As Long

            lrow = MSFlexGrid2.Row
            lcol = MSFlexGrid2.Col

            MSFlexGrid2.Col = lcol
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.SetFocus
            
        Case 9 'Kalk und Rund
        
            If Text2(3).Text <> "" And Text2(0).Text <> "" Then
                Screen.MousePointer = 0
                MsgBox "Bitte entscheiden Sie sich für eine Kalkulationsvariante!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
        
            If Text2(3).Text <> "" Then 'Aufschlag in prozent auf listenvk
                If IsNumeric(Text2(3).Text) Then
                    KalkandRund "LISTENVK"
                End If
            
            ElseIf Text2(0).Text <> "" Then 'Aufschlagsfaktor auf listenek
            
                If IsNumeric(Text2(0).Text) Then
                    KalkandRund "LISTENEK"
                End If
                
            Else
                KalkandRund "RUNDEN"
            End If
            
            Faerbewegenpreisunter MSFlexGrid2, SpaltennummerKVKALT, SpaltennummerKVKNEU
            
            
        Case 10 'Ausnahme
            Frame3.Visible = True
            ZeigNOKALKL
            
        Case 11 'Hilfe
            
            gsHelpstring = "Stammdaten einlesen \ Lieferanten von der Kalkulation ausschließen"
            frmWKL110.Show 1
            
        Case 12 'Kalk und Rund
        
            If Text2(3).Text <> "" And Text2(0).Text <> "" Then
                Screen.MousePointer = 0
                MsgBox "Bitte entscheiden Sie sich für eine Kalkulationsvariante!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
        
            If Text2(3).Text <> "" Then 'Aufschlag in prozent auf listenvk
                If IsNumeric(Text2(3).Text) Then
                    nurKalk "LISTENVK"
                Else
                
                End If
            ElseIf Text2(0).Text <> "" Then 'Aufschlagsfaktor auf listenek
                If IsNumeric(Text2(0).Text) Then
                    nurKalk "LISTENEK"
                Else
                
                End If
            Else
                nurKalk "RUNDEN"
            End If
            
            Faerbewegenpreisunter MSFlexGrid2, SpaltennummerKVKALT, SpaltennummerKVKNEU
            
        Case 13 'LEK-Abschlag
        
            RechneLEK
        

    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command3_Click"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub RechneLEK()
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dLEKPR          As Double
    Dim dAufschlag      As Double
    Dim cArtNr          As String
   
    MSFlexGrid2.Redraw = False
    
    If IsNumeric(Text2(2).Text) = True Then
        dAufschlag = CDbl(Text2(2).Text) * -1
        
        MSFlexGrid2.Row = 0
        For i = 1 To MSFlexGrid2.Rows - 1
        
            MSFlexGrid2.Row = i
            MSFlexGrid2.Col = SpaltennummerArtnr
            cArtNr = MSFlexGrid2.Text
                        
            MSFlexGrid2.Row = i
            MSFlexGrid2.Col = SpaltennummerLEKPR
            If Not Len(MSFlexGrid2.Text) = 0 Then
                If IsNumeric(MSFlexGrid2.Text) Then
                    dLEKPR = CDbl(MSFlexGrid2.Text)
                
                    If dLEKPR <> 0 Then
                        dLEKPR = dLEKPR + ((dLEKPR * dAufschlag) / 100)
                        MSFlexGrid2.Text = Format(dLEKPR, "###,##0.00")
                    End If
                End If
            End If
        Next i
    End If
         
    MSFlexGrid2.Redraw = True
    
'    anzeige "normal", "", Label22(2)

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RechneLEK"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Faerbewegenpreisunter(gridx As MSFlexGrid, spaltePreisAlt As Byte, spaltePreisNeu As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j               As Integer
    Dim dpreisAlt       As Double
    Dim dpreisNeu       As Double
    Dim lAnz            As Long
    
    lAnz = 0
    With gridx
        .Redraw = False
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spaltePreisAlt
            
            If .Text <> "" Then
                If IsNumeric(.Text) Then
                    dpreisAlt = CDbl(.Text)
                End If
            End If
            
            .Row = j
            .Col = spaltePreisNeu
            
            If .Text <> "" Then
                If IsNumeric(.Text) Then
                    dpreisNeu = CDbl(.Text)
                End If
            End If
            
            
            
            If dpreisAlt > 0 Then
                If dpreisAlt <> dpreisNeu Then
                
                    'hey ne Preisveränderung
                    lAnz = lAnz + 1
                    .Col = spaltePreisNeu
                    .CellBackColor = &HFF00FF
                Else
                    .Col = spaltePreisNeu
                    .CellBackColor = &H80000005
                   
                End If
            End If
        Next j
        .Redraw = True
    End With
    
    Label4(35).Caption = "Preisänderungen (" & lAnz & ")"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Faerbewegenpreisunter"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Rundenwkl11()
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dZahl           As Double
    Dim sStelle         As String
    Dim sZahl           As String
    
    sStelle = TxtRunden.Text
    
    MSFlexGrid2.Row = 0
    For i = 1 To MSFlexGrid2.Rows - 1
        MSFlexGrid2.Row = i
        MSFlexGrid2.Col = SpaltennummerKVKNEU
        If Not Len(MSFlexGrid2.Text) = 0 Then
            If Len(Trim(sStelle)) = 1 Then
                MSFlexGrid2.Text = Left(MSFlexGrid2.Text, Len(MSFlexGrid2.Text) - 1) & TxtRunden.Text
            ElseIf Len(Trim(sStelle)) = 2 Then
                MSFlexGrid2.Text = Left(MSFlexGrid2.Text, Len(MSFlexGrid2.Text) - 2) & TxtRunden.Text
            Else
            
            End If
        Else
            MSFlexGrid2.Text = "0,00"
        End If
    Next i
    
    MSFlexGrid2.Refresh
        
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Rundenwkl11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR
    
    loeschDat Datmark
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function Datmark() As String
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cPfad As String

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Datmark = ""
    
    If File2.ListIndex < 0 Then
        
    Else
         For lcount = 0 To File2.ListCount - 1
            If File2.Selected(lcount) = True Then
                Datmark = cPfad & File2.list(lcount)
                Exit Function
            End If
        Next lcount
    End If
    
    If File1.ListIndex < 0 Then
        
    Else
         For lcount = 0 To File1.ListCount - 1
            If File1.Selected(lcount) = True Then
                Datmark = cPfad & "IN\" & File1.list(lcount)
                
            End If
        Next lcount
    End If
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Datmark"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen: Fremdformate ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub File1_Click()
    On Error GoTo LOKAL_ERROR
    
    Label9.Caption = ""
    Label9.Refresh
    
    If IsNumeric(Mid(File1.list(File1.ListIndex), 2, 6)) Then
        Label24.Caption = ermLiefBez(CLng(Mid(File1.list(File1.ListIndex), 2, 6)))
        Label24.Refresh
    Else
    
        
    
        Label24.Caption = "Wichtig!! Wochendatei(Master...) in der numerischen Reihenfolge einlesen!!"
        Label24.Refresh
        
    End If
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File1_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub File1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command1_Click 0
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    
    
End Sub
Private Sub File2_Click()
    On Error GoTo LOKAL_ERROR
    
    Label9.Caption = ""
    Label9.Refresh
    
    If IsNumeric(Mid(File2.list(File2.ListIndex), 2, 6)) Then
        Label24.Caption = ermLiefBez(CLng(Mid(File2.list(File2.ListIndex), 2, 6)))
        Label24.Refresh
    Else
        Label24.Caption = "Wichtig!! Wochendatei(Master...) in der numerischen Reihenfolge einlesen!!"
        Label24.Refresh
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File2_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub File2_DBLClick()
    On Error GoTo LOKAL_ERROR
    
    Command6_Click

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File2_DBLClick"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub File2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    
        Dim iRet As Integer
        Dim cPfad As String
        Dim cLBSatz As String
    
        cPfad = gcDBPfad
        If Right(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        cLBSatz = File2.list(File2.ListIndex)
        
        Select Case KeyCode
            Case Is = 46    '//Del
                iRet = MsgBox("Wollen Sie wirklich diese Datei löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    Kill cPfad & cLBSatz
                    File2.Refresh
                End If
        End Select

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    
        Dim iRet As Integer
        Dim cPfad As String
        Dim cLBSatz As String
    
        cPfad = gcDBPfad
        If Right(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        cPfad = cPfad & "in\"
        
        cLBSatz = File1.list(File1.ListIndex)
        
        Select Case KeyCode
            Case Is = 46    '//Del
                iRet = MsgBox("Wollen Sie wirklich diese Datei löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    Kill cPfad & cLBSatz
                    File1.Refresh
                End If
        End Select

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub loeschDat(cLBSatz As String)
    On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    
    If cLBSatz = "" Then
        Exit Sub
    End If
    
    iRet = MsgBox("Wollen Sie wirklich diese Datei löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        Kill cLBSatz
        File1.Refresh
        File2.Refresh
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loeschDat"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub Form_Activate()

'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
'   If Command1(4).Enabled Then
'       Command1(4).Enabled = False
'       Pause (1)
'       Command1_Click 4
'   End If
'Odayy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    Dim cPfad As String

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Kill cPfad & "IN\MASTER!.000"
    
    loeschNEW "Master", gdBase
    loeschNEW "EANKoml1", gdBase
    loeschNEW "MASTEMP", gdBase

    Frame4.Visible = False
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Unload"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(21).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case index
        
        Case Is = 21 'Lieferantenübersicht
        
            URLGoTo Me.hwnd, "http://kisslive.de/sortimente/parfuemerie/lieferanten.html"

    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If index = 21 Then
        Label1(21).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub MSFlexGrid1_EnterCell()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row > 0 And MSFlexGrid1.Col > 2 Then
        MSFlexGrid1.CellBackColor = &HC00000
        MSFlexGrid1.CellForeColor = vbYellow
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_EnterCell"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_LeaveCell()
    On Error GoTo LOKAL_ERROR

    If MSFlexGrid1.Row > 0 And MSFlexGrid1.Col > 2 Then
        MSFlexGrid1.CellBackColor = vbWhite
        MSFlexGrid1.CellForeColor = vbBlack
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid2.Row > 1 Then
'        Command2_Click 0
    Else
        sortierenGrid MSFlexGrid2
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid2.Row
    lcol = MSFlexGrid2.Col
    
    If MSFlexGrid2.Col = SpaltennummerAGN Or MSFlexGrid2.Col = SpaltennummerGEFUEHRT _
    Or MSFlexGrid2.Col = SpaltennummerRABATT_OK Or MSFlexGrid2.Col = SpaltennummerPREISSCHU _
    Or MSFlexGrid2.Col = SpaltennummerMWST Or MSFlexGrid2.Col = SpaltennummerKVKPR1 Then
    
        If iKeypress = 0 And KeyCode <> vbKeyBack Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            MSFlexGrid2.Text = ""
        End If
        iKeypress = iKeypress + 1
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lcol     As Long
    Dim lrow     As Long
    Dim cZeichen As String
    Dim cFeld    As String

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)

    cFeld = MSFlexGrid2.Text

    Select Case KeyAscii
        Case Is = 8
            If Len(cFeld) > 0 Then
                cFeld = Left$(cFeld, Len(cFeld) - 1)
            End If
        Case Else
            cFeld = cFeld & Chr$(KeyAscii)
    End Select

    MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, MSFlexGrid2.Col) = cFeld
    MSFlexGrid2.Refresh

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    If KeyCode = vbKeyF2 Then
        lrow = MSFlexGrid2.Row
        lcol = MSFlexGrid2.Col
        
        
        
        gsARTNR = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, SpaltennummerArtnr)
        If Left(gsARTNR, 1) = "X" Then
            gsARTNR = Right(gsARTNR, Len(gsARTNR) - 2)
        End If
        
        If gsARTNR <> "" Then
            If IsNumeric(gsARTNR) Then
                frmWKL10.Show 1
                Me.Refresh
                Screen.MousePointer = 0
            End If
        End If
        gsARTNR = ""
    End If

    '** 7= VKPR, 9=KVKPR_Neu, 14=MINBEST, 15=GEFÜHRT, 16=RABATT_OK, 17=PREISSCHUTZ **
    If MSFlexGrid2.Col = SpaltennummerAGN Or MSFlexGrid2.Col = SpaltennummerGEFUEHRT _
    Or MSFlexGrid2.Col = SpaltennummerRABATT_OK Or MSFlexGrid2.Col = SpaltennummerPREISSCHU _
    Or MSFlexGrid2.Col = SpaltennummerMWST Or MSFlexGrid2.Col = SpaltennummerKVKPR1 Then
        Select Case KeyCode
            Case Is = 46    '** Delete **
                MSFlexGrid2.Text = ""
        End Select
    End If

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub MSFlexGrid2_Click()
    On Error GoTo LOKAL_ERROR

    If MSFlexGrid2.Col = SpaltennummerArtnr Then
        If Left(MSFlexGrid2.Text, 1) = "X" Then
            MSFlexGrid2.Text = Right(MSFlexGrid2.Text, Len(MSFlexGrid2.Text) - 2)
        Else
            MSFlexGrid2.Text = "X " & MSFlexGrid2.Text
        End If
    ElseIf MSFlexGrid2.Col = SpaltennummerVKPR Then
        MSFlexGrid2.SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz As Long
    Dim lrow As Long
    
    lrow = MSFlexGrid1.Row
    MSFlexGrid1.Col = 0
    If MSFlexGrid1.Text = "X" Then
        MSFlexGrid1.Text = ""
    Else
        MSFlexGrid1.Text = "X"
    End If
    
    lAnz = ermX
    If lAnz = 1 Then
        anzeige "normal", "Es ist ein Lieferant ausgewählt.", Label5
    Else
        anzeige "normal", "Es sind " & lAnz & " Lieferanten ausgewählt.", Label5
    End If
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = lrow
    If IsNumeric(MSFlexGrid1.Text) Then
        Text1(0).Text = MSFlexGrid1.Text
    End If
    
    
    'die Bezeichnung auch
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Row = lrow
    
    Text1(1).Text = MSFlexGrid1.Text
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermX() As Byte
    On Error GoTo LOKAL_ERROR
    
    ermX = 0
    
    Dim lrow    As Long
    Dim lRows   As Long
    
    MSFlexGrid1.Redraw = False
    
    lRows = MSFlexGrid1.Rows - 1
    
    For lrow = 1 To lRows
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.Col = 0
    
        If MSFlexGrid1.Text = "X" Then
           ermX = ermX + 1
        End If
        
    Next lrow
    
    MSFlexGrid1.Redraw = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermX"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub DeKompMasterDateiWKL11()
    On Error GoTo LOKAL_ERROR
    
    Dim lfail   As Long
    Dim lRet    As Long
    Dim cQuelle As String
    Dim cZiel   As String
    Dim cPfad   As String
    Dim cdatei  As String
    Dim ctmp    As String
    Dim lPos    As String
    Dim iStufe  As Integer
    Dim iFileNr As Integer
    Dim Task$, hProcess&, result&
    Dim t       As Integer
    Dim sSQL    As String
    
    '* Arbeitsschritte: *******************************
    '* 1. Umbenennen der Wochen-Master-Datei        ***
    '* 2. Entpacken der Master-Datei                ***
    '* 3. Kopieren von MLISRT und MASTER            ***
    '**************************************************
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    loeschNEW "Master", gdBase
    Kill cPfad & "IN\MASTER!.000"
    
    iStufe = 0
    '*** UMBENENNEN ***
    If File1.ListCount = 0 Then
        MsgBox "Keine komprimierte Master-Datei zum Entpacken gefunden!" & vbCrLf & "Suche nach entpackter Master-Datei...", vbCritical, "STOP!"
        
        iFileNr = FreeFile
        Open cPfad & "DATENUPD.TXT" For Binary As #iFileNr
        lPos = LOF(iFileNr) + 1
        ctmp = Format$(Now, "DD.MM.YYYY HH:MM:SS") & " Keine MASTER! vorhanden!" & vbCrLf
        Put #iFileNr, lPos, ctmp
        Close iFileNr

        GoTo NEXT_STEP
    ElseIf File1.ListCount = 1 Then
        cdatei = File1.list(0)
    ElseIf File1.ListIndex < 0 Then
        MsgBox "Bitte eine Datei zum Dekomprimieren auswählen!", vbCritical, "STOP!"
        File1.SetFocus
        Exit Sub
    Else
        cdatei = File1.list(File1.ListIndex)
    End If
    
    
    cdatei = UCase$(cdatei)
    
    iStufe = 1
    iFileNr = FreeFile
    Open cPfad & "DATENUPD.TXT" For Binary As #iFileNr
    lPos = LOF(iFileNr) + 1
    ctmp = Format$(Now, "DD.MM.YYYY HH:MM:SS") & " " & cdatei & vbCrLf
    Put #iFileNr, lPos, ctmp
    Close iFileNr
    
    iStufe = 2
    If cdatei <> "MASTER!.000" Then
        cQuelle = cPfad & "IN\" & cdatei
        cZiel = cPfad & cdatei
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
        If lRet = 0 Then
            MsgBox "Konnte " & cQuelle & " nicht kopieren!", vbCritical, "STOP!"
        End If
        Kill cPfad & "IN\MASTER!.000"
        
        iStufe = 3
        Name cPfad & "IN\" & cdatei As cPfad & "IN\MASTER!.000"
    End If
    
    Kill cPfad & "IN\*.DBF"
    
    '*** ENTPACKEN ***
    
    Zip_Unzip "ICHAG", cPfad & "IN", cPfad & "IN\MASTER!.000", txtStatus
    
    If FileExists(cPfad & "IN\MASTER!.000") = False Then
        Exit Sub
    End If
        
    '*** KOPIEREN ***
    
NEXT_STEP:
    iFileNr = FreeFile
    Open cPfad & "IN\MLISRT.DBF" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\MLISRT.DBF"
        GoTo NEXT_STEP1
    Else
        Close iFileNr
        
    End If
    loesch "MLISRT"
        
    'unterscheidung 3
    
    cQuelle = cPfad & "IN\MLISRT.DBF"
    cZiel = cPfad & "MLISRT.DBF"
    
    sSQL = "Select * into MLISRT from MLISRT IN '" & cPfad & "IN" & "' 'dbase IV;'"
    gdBase.Execute sSQL, dbFailOnError
        
NEXT_STEP1:
    iFileNr = FreeFile
    Open cPfad & "IN\MASTER.DBF" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\MASTER.DBF"
        GoTo NEXT_STEP1a
    Else
        Close iFileNr
    End If
    
    loeschNEW "Master", gdBase
    
    cQuelle = cPfad & "IN\MASTER.DBF"
    cZiel = cPfad & "MASTER.DBF"
    
    sSQL = "Select * into master from master IN '" & cPfad & "IN" & "' 'dbase IV;'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from master where PGN = 60 "
    gdBase.Execute sSQL, dbFailOnError
        

    Kill cQuelle
    
    cQuelle = cPfad & "IN\MASTER.txt"
    cZiel = cPfad & "MASTER.txt"
    
    lRet = CopyFile(cQuelle, cZiel, lfail)
    If lRet = 0 Then
'        MsgBox "Konnte " & cQuelle & " nicht kopieren!", vbCritical, "STOP!"
    End If
    
NEXT_STEP1a:
    iFileNr = FreeFile
    Open cPfad & "IN\MASTER2.DBF" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\MASTER2.DBF"
        GoTo NEXT_STEP3a
    Else
        Close iFileNr
    End If
    
    loeschNEW "MASTER2", gdBase
    
    cQuelle = cPfad & "IN\MASTER2.DBF"
    cZiel = cPfad & "MASTER2.DBF"
    
    sSQL = "Select * into MASTER2 from MASTER2 IN '" & cPfad & "IN" & "' 'dbase IV;'"
    gdBase.Execute sSQL, dbFailOnError
    
    Kill cQuelle
    
NEXT_STEP3a:
    iFileNr = FreeFile
    Open cPfad & "IN\LIEFKURZ.DBF" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\LIEFKURZ.DBF"
        GoTo NEXT_STEP2
    Else
        Close iFileNr
    End If
    
    loeschNEW "LIEFKURZ", gdBase
    
    cQuelle = cPfad & "IN\LIEFKURZ.DBF"
    cZiel = cPfad & "LIEFKURZ.DBF"
    
    sSQL = "Select * into LIEFKURZ from LIEFKURZ IN '" & cPfad & "IN" & "' 'dbase IV;'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Lisrt inner join LIEFKURZ on Lisrt.linr = Liefkurz.Linr "
    sSQL = sSQL & " set Lisrt.Kuerzel = Liefkurz.kuerzel "
    gdBase.Execute sSQL, dbFailOnError
    
    Kill cQuelle
    
NEXT_STEP2:
    File1.Pattern = "*.*"
    File1.Refresh
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3376 Or err.Number = 3043 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DeKompMasterDateiWKL11"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Function DeKompStada() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lfail   As Long
    Dim lRet    As Long
    Dim cQuelle As String
    Dim cZiel   As String
    Dim cPfad   As String
    Dim cdatei  As String
    Dim ctmp    As String
    Dim lPos    As String
    Dim iStufe  As Integer
    Dim iFileNr As Integer
    
    DeKompStada = False
    
    '* Arbeitsschritte: *******************************
    '* 1. Umbenennen der Wochen-Master-Datei        ***
    '* 2. Entpacken der Master-Datei                ***
    '* 3. Kopieren von MLISRT und MASTER            ***
    '**************************************************
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Kill cPfad & "IN\MASTER!.000"
    
'    File1.Pattern = "M*!.*"
'    File1.Refresh
    
    iStufe = 0
    '*** UMBENENNEN ***
    If File1.ListCount = 0 Then
        MsgBox "Keine Stammdaten zum Entpacken gefunden!", vbInformation, "Winkiss Hinweis:"
        
        iFileNr = FreeFile
        Open cPfad & "DATENUPD.TXT" For Binary As #iFileNr
        lPos = LOF(iFileNr) + 1
        ctmp = Format$(Now, "DD.MM.YYYY HH:MM:SS") & " Keine MASTER! vorhanden!" & vbCrLf
        Put #iFileNr, lPos, ctmp
        Close iFileNr

        Exit Function
    ElseIf File1.ListCount = 1 Then
        cdatei = UCase$(File1.list(0))
    ElseIf File1.ListIndex < 0 Then
        MsgBox "Bitte Stammdaten auswählen!", vbInformation, "Winkiss Hinweis:"
        File1.SetFocus
        Exit Function
    Else
        cdatei = UCase$(File1.list(File1.ListIndex))
    End If
    
    iStufe = 1
    iFileNr = FreeFile
    Open cPfad & "DATENUPD.TXT" For Binary As #iFileNr
    lPos = LOF(iFileNr) + 1
    ctmp = Format$(Now, "DD.MM.YYYY HH:MM:SS") & " " & cdatei & vbCrLf
    Put #iFileNr, lPos, ctmp
    Close iFileNr
    
    iStufe = 2
    If cdatei <> "MASTER!.000" Then
        cQuelle = cPfad & "IN\" & cdatei
        cZiel = cPfad & cdatei
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
        If lRet = 0 Then
            MsgBox "Konnte " & cQuelle & " nicht kopieren!", vbInformation, "Winkiss Hinweis:"
        End If
        Kill cPfad & "IN\MASTER!.000"
        
        cQuelle = cPfad & "IN\" & cdatei
        cZiel = cPfad & "IN\MASTER!.000"
        
        lRet = CopyFile(cQuelle, cZiel, lfail)
        If lRet = 0 Then
            MsgBox "Konnte " & cQuelle & " nicht kopieren!", vbInformation, "Winkiss Hinweis:"
        End If
    End If
    
    Kill cPfad & "IN\*.CSV"
    
    '*** ENTPACKEN ***
    
    Zip_Unzip "ICHAG", cPfad & "IN", cPfad & "IN\MASTER!.000", txtStatus
    
    If FileExists(cPfad & "IN\MASTER!.000") = False Then
        Exit Function
    End If
       
   
    '*** KOPIEREN ***
    iFileNr = FreeFile
    Open cPfad & "IN\MLISRT.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\MLISRT.CSV"
    Else
        Close iFileNr
    End If
    
    cQuelle = cPfad & "IN\MLISRT.CSV"
    csvImport "MLISRT", gdBase, cQuelle, Label9
    Kill cQuelle
    
    iFileNr = FreeFile
    Open cPfad & "IN\LINBEZ.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\LINBEZ.CSV"
    Else
        Close iFileNr
    End If
    
    cQuelle = cPfad & "IN\LINBEZ.CSV"
    csvImport "MLINBEZ", gdBase, cQuelle, Label9
    Kill cQuelle
        
    iFileNr = FreeFile
    Open cPfad & "IN\MASTER.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\MASTER.CSV"
    Else
        Close iFileNr
    End If

    cQuelle = cPfad & "IN\MASTER.CSV"
    csvImport "MASTER", gdBase, cQuelle, Label9
    
    
    Dim sSQL As String
    
    sSQL = "Delete from Master where PGN = 60 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Kill cQuelle
    
    
    
    iFileNr = FreeFile
    Open cPfad & "IN\LIEFKURZ.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\LIEFKURZ.CSV"
    Else
        Close iFileNr
    End If

    cQuelle = cPfad & "IN\LIEFKURZ.CSV"
    csvImport "LIEFKURZ", gdBase, cQuelle, Label9
    Kill cQuelle
    
    
    If NewTableSuchenDBKombi("LIEFKURZ", gdBase) Then
        sSQL = "Update Lisrt inner join LIEFKURZ on Lisrt.linr = Liefkurz.Linr "
        sSQL = sSQL & " set Lisrt.Kuerzel = Liefkurz.kuerzel "
        gdBase.Execute sSQL, dbFailOnError
        
        loeschNEW "LIEFKURZ", gdBase
    End If
    
    'Artean2
    iFileNr = FreeFile
    Open cPfad & "IN\ARTEAN2.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\ARTEAN2.CSV"
    Else
        Close iFileNr
    End If

    cQuelle = cPfad & "IN\ARTEAN2.CSV"
    csvImport "ARTEAN2", gdBase, cQuelle, Label9
    Kill cQuelle
    

    'Artean3
    iFileNr = FreeFile
    Open cPfad & "IN\ARTEAN3.CSV" For Binary As #iFileNr
    lPos = LOF(iFileNr)
    If lPos = 0 Then
        Close iFileNr
        Kill cPfad & "IN\ARTEAN3.CSV"
    Else
        Close iFileNr
    End If

    cQuelle = cPfad & "IN\ARTEAN3.CSV"
    csvImport "ARTEAN3", gdBase, cQuelle, Label9
    Kill cQuelle
    
    

    If NewTableSuchenDBKombi("MASTER", gdBase) Then
        If Datendrin("MASTER", gdBase) Then
            File1.Pattern = "*.*"
            File1.Refresh
    
            Kill cPfad & "IN\" & cdatei
            DeKompStada = True
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DeKompStada"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function


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
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim btxt As Boolean
    Dim iFileNr As Integer
    Dim lRet As Long
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case index
        Case 0 To 2, 5
            Select Case index
                Case Is = 0 'alle Artikel hier auf geführtcheck achten
                If CG.value = vbChecked Then
                    reportbildschirm "dWKL11a", "aWKL11e"
                Else
                    reportbildschirm "dWKL11a", "aWKL11a"
                End If
                Case Is = 1 'neue Artikel
                    reportbildschirm "dWKL11b", "aWKL11b"
                Case Is = 2 'Preisänderungen hier auf geführtcheck achten
                
                If CG.value = vbChecked Then
                    reportbildschirm "dWKL11a", "aWKL11f"
                Else
                    reportbildschirm "dWKL11a", "aWKL11c"
                End If
                
                Case Is = 5 'Preisänderungen durch automatische Kalkulation
                    reportbildschirm "dWKL11c", "aWKL11d"
            End Select
        
        Case Is = 3     '** Schließen **
            Frame7.Visible = False: CG.Visible = False: Label26(1).Visible = True
        Case Is = 4     '** Wochenänderung txt - Datei **
            lRet = Shell("WRITE.EXE " & cPfad & "Master.txt", vbMaximizedFocus)
        Case 6
            If NewTableSuchenDBKombi("LPZT", gdBase) Then
                checkDieLinien2
            Else
                anzeige "rot", "Erst die Stammdaten einlesen, dann Protokoll drucken.", Label9
            End If
        
    End Select
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub Command5_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim cLBSatz     As String
    Dim cTitel      As String
    Dim cTitel2     As String
    Dim lAnzSeiten  As Long
    Dim iRet        As Integer
    Dim cSQL        As String
    
    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0     '** Drucken **

            loeschNEW "DRU_LISTE", gdBase
            
            cSQL = "Create Table DRU_LISTE "
            cSQL = cSQL & "(FELD1 Text(200), TITEL Text(200), TITEL2 Text(200) )"
            schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
            
            cTitel2 = "Übernahme-Protokoll vom " & Format$(Now, "DD.MM.YYYY")
            cTitel = List1.list(0)
            cSQL = "Insert into DRU_LISTE ( TITEL, TITEL2 ) values ('" & cTitel & "', '" & cTitel2 & "' ) "
            schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
            
            For lcount = 0 To List2.ListCount - 1
                Label8(0).Caption = Trim$(Str$(lcount + 1))
                Label8(0).Refresh
                cLBSatz = List2.list(lcount)
                cSQL = "Insert into DRU_LISTE ( TITEL, TITEL2, FELD1 ) values ('" & cTitel & "', '" & cTitel2 & "', '" & cLBSatz & "' ) "
                schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
            Next lcount
            
            reportbildschirm "WKL034", "aWKL11"
            


        Case Is = 1     '** Schließen **
            With Frame7
                .Visible = True

            End With
            With Frame1


                .Visible = True
            End With
            Frame8.Visible = False
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        
    End If
End Sub
Private Sub Command6_Click()

    On Error GoTo LOKAL_ERROR
    Dim sDatei As String
    Screen.MousePointer = 11
    Label9.Caption = ""
    Label9.Refresh
    
    If File2.ListIndex < 0 Then
        Screen.MousePointer = 0
        Label6.Caption = "Bitte eine Datei in der Liste auswählen!"
        Label6.Refresh
        File2.SetFocus
    Else
        
        KopiereDateiWKL11
        
        File2.Refresh
        File1.Refresh
        DoEvents
        
        If File1.ListCount > 0 Then
            File1.Selected(0) = True
        End If

        Frame7.Visible = False
        Command1(0).SetFocus

        Label6.Caption = "Bitte eine Datei in der Liste auswählen!"
        Label6.Refresh
                
    End If
        
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command7_Click(index As Integer)

    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0
            If NewTableSuchenDBKombi("MLINBEZ", gdBase) Then
                If Datendrin("MLINBEZ", gdBase) Then
                    AlleLinBezCSV
                Else
                    AlleLinBezNehmenWKL11
                End If
            Else
                AlleLinBezNehmenWKL11
            End If
    
        Case Is = 1
            DruckeHinweiseStammdatenWKL11
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub AlleLinBezNehmenWKL11()
    On Error GoTo LOKAL_ERROR
    
    Dim lfail       As Long
    Dim lcount      As Long
    Dim lRet        As Long
    Dim lDateIN     As Date
    Dim lDate       As Date
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfadIN     As String
    Dim cZiel       As String
    Dim cQuelle     As String
    Dim cSQL        As String
    Dim cSQL1       As String
    Dim cLinrIN     As String
    Dim ctmp        As String
    Dim cTempo      As String
    Dim i           As Integer
    Dim bgefunden   As Boolean
    Dim bDatOk      As Boolean
    Dim rsrs        As Recordset
    Dim rsRsIN      As Recordset
    Dim dbIN        As Database
    
    bDatOk = False
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
   '**********************************LINBEZ suchen
    For i = 0 To File1.ListCount - 1
        
        bgefunden = False
        If UCase(File1.list(i)) = "LINBEZ.DBF" Then
            cdatei = UCase$(File1.list(i))
            bgefunden = True
            Exit For
        Else
'            MsgBox "Keine Datei mit Linienbezeichnungen vorhanden!", vbCritical, "STOP!"
            bgefunden = False
        End If
    Next i
    
    If bgefunden = True Then
    
'    If NewTableSuchenDBKombi("MLINBEZ", gdBase) Then
    
        cPfadIN = cPfad & "IN\"    '**LINBEZ.DBF"
        Set dbIN = OpenDatabase(cPfadIN, False, False, "dBase IV;")
        
        cSQL = "Select LINR,LINBEZEICH from LINBEZ where LINR = 300200"
        Set rsRsIN = dbIN.OpenRecordset(cSQL)
        
        
        If NewTableSuchenDBKombi("LINBEZ", gdBase) Then
        
            cSQL = "Select LINR,LINBEZEICH from LINBEZ where LINR = 300200"
            Set rsrs = gdBase.OpenRecordset(cSQL)
        
            '** LINR 300200 ist immer vorhanden **
            If Not rsRsIN.EOF Then
                If Not IsNull(rsRsIN!LINBEZEICH) Then
                    cTempo = (rsRsIN!LINBEZEICH)
                    lDateIN = CDate(cTempo)
        
                Else
                    lDateIN = 0
                End If
                
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LINBEZEICH) Then
                        lDate = CDate(rsrs!LINBEZEICH)
                        If lDate < lDateIN Then
                            rsrs.Close: Set rsrs = Nothing
                            bDatOk = True
                        Else
                            rsrs.Close: Set rsrs = Nothing
                            bDatOk = False
                        End If
                    Else
                        lDate = 0
                    End If
                Else
                    rsrs.Close: Set rsrs = Nothing
                    bDatOk = True
                End If
            
            
            End If
        Else
            bDatOk = True
            
        End If
    End If
    
    
    If bgefunden = True And bDatOk = True Then
        pbrLinbez.Max = 100
        pbrLinbez.Visible = True
        pbrLinbez.value = 20
        
        'Hat sich etwas geändert
        'Sind Linien raus
        'sind neue Linien drin
        
        checkDieLinien
        
        lbl9.Caption = "Liniendatei wird aktualisiert..."
        lbl9.Refresh
        
        loeschNEW "LINBTEMP", gdBase
        
        cSQL = "Create Table LINBTEMP "
        cSQL = cSQL & "( LINR long"
        cSQL = cSQL & ", LINBEZEICH Text(30)"
        cSQL = cSQL & ", LPZ integer"
        cSQL = cSQL & ", MARKER Text(1) "
        cSQL = cSQL & ", KUERZEL Text(5) "
        cSQL = cSQL & ", MARKE Text(20) "
        cSQL = cSQL & ", SORTI integer" 'neu
        cSQL = cSQL & " ) "
        gdBase.Execute cSQL, dbFailOnError
        
        If NewTableSuchenDBKombi("LINBEZ", gdBase) Then
            cSQL = "Insert into LINBTEMP Select * from LinBez "
            cSQL = cSQL & " WHERE LINR BETWEEN 500000 AND 699999 or LPZ > 800 "
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        pbrLinbez.value = 40
        
        loeschNEW "TPLINBEZ", gdBase
        
        cSQL = "Select * into tplinbez from linbez IN '" & cPfad & "IN" & "' 'dbase IV;'"
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 60
        
        cSQL = "Insert into LINBTEMP Select * from TpLinBez "
        cSQL = cSQL & " WHERE LINR not BETWEEN 500000 AND 699999 "
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 80
        
        cSQL = "Update LINBTEMP inner join LINBEZ on  LINBTEMP.linr = linBez.linr  "
        cSQL = cSQL & " and LINBTEMP.lpz = LINBEZ.LPZ set LINBTEMP.KUERZEL = linbez.KUERZEL  "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "LINBEZ", gdBase
        
        cSQL = "Select * into LINBEZ from LINBTEMP"
        cSQL = cSQL & " order by LINR, LPZ "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "update linbez set kuerzel = '' , Marke = '' where linbezeich = '/////'"
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 100
        pbrLinbez.Visible = False
    End If
    
    cSQL = "update linbez set sorti = lpz where sorti is null"
    gdBase.Execute cSQL, dbFailOnError
    
    Kill cPfad & "IN\MASTER.TXT"
    Kill cPfad & "IN\MLISRT.DBF"
    
    EinzelnLisrtNehmenWKL11
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 75 Then
        Resume Next
    ElseIf err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "AlleLinBezNehmenWKL11"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Sub AlleLinBezCSV()
    On Error GoTo LOKAL_ERROR

    Dim lDateIN     As Date
    Dim lDate       As Date
    Dim cSQL        As String
    Dim bDatOk      As Boolean
    Dim rsrs        As Recordset
    Dim rsRsIN      As Recordset
    
    bDatOk = False
    
    If NewTableSuchenDBKombi("MLINBEZ", gdBase) Then

        cSQL = "Select LINR,LINBEZEICH from MLINBEZ where LINR = 300200"
        Set rsRsIN = gdBase.OpenRecordset(cSQL)
        
        If NewTableSuchenDBKombi("LINBEZ", gdBase) Then
        
            cSQL = "Select LINR,LINBEZEICH from LINBEZ where LINR = 300200"
            Set rsrs = gdBase.OpenRecordset(cSQL)
        
            '** LINR 300200 ist immer vorhanden **
            If Not rsRsIN.EOF Then
                If Not IsNull(rsRsIN!LINBEZEICH) Then
                    lDateIN = CDate(rsRsIN!LINBEZEICH)
                Else
                    lDateIN = 0
                End If
                
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LINBEZEICH) Then
                        lDate = CDate(rsrs!LINBEZEICH)
                        If lDate < lDateIN Then
                            rsrs.Close: Set rsrs = Nothing
                            bDatOk = True
                        Else
                            rsrs.Close: Set rsrs = Nothing
                            bDatOk = False
                        End If
                    Else
                        lDate = 0
                    End If
                Else
                    rsrs.Close: Set rsrs = Nothing
                    bDatOk = True
                End If
            End If
        Else
            bDatOk = True
            
        End If
    End If
    
    If bDatOk = True Then
        pbrLinbez.Max = 100
        pbrLinbez.Visible = True
        pbrLinbez.value = 20
        
        'Hat sich etwas geändert
        'Sind Linien raus
        'sind neue Linien drin
        
        checkDieLinien
        
        lbl9.Caption = "Liniendatei wird aktualisiert..."
        lbl9.Refresh
        
        loeschNEW "LINBTEMP", gdBase
        
        cSQL = "Create Table LINBTEMP "
        cSQL = cSQL & "( LINR long"
        cSQL = cSQL & ", LINBEZEICH Text(30)"
        cSQL = cSQL & ", LPZ integer"
        cSQL = cSQL & ", MARKER Text(1) "
        cSQL = cSQL & ", KUERZEL Text(5) "
        cSQL = cSQL & ", MARKE Text(20) "
        cSQL = cSQL & ", SORTI integer" 'neu
        cSQL = cSQL & " ) "
        gdBase.Execute cSQL, dbFailOnError
        
        If NewTableSuchenDBKombi("LINBEZ", gdBase) Then
            cSQL = "Insert into LINBTEMP Select * from LinBez "
            cSQL = cSQL & " WHERE LINR BETWEEN 500000 AND 699999 or LPZ > 800 "
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        pbrLinbez.value = 40
        
        loeschNEW "TPLINBEZ", gdBase
        
        cSQL = "Select * into tplinbez from Mlinbez "
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 60
        
        cSQL = "Insert into LINBTEMP Select * from TpLinBez "
        cSQL = cSQL & " WHERE LINR not BETWEEN 500000 AND 699999 "
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 80
        
        cSQL = "Update LINBTEMP inner join LINBEZ on  LINBTEMP.linr = linBez.linr  "
        cSQL = cSQL & " and LINBTEMP.lpz = LINBEZ.LPZ set LINBTEMP.KUERZEL = linbez.KUERZEL  "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "LINBEZ", gdBase
        
        cSQL = "Select * into LINBEZ from LINBTEMP"
        cSQL = cSQL & " order by LINR, LPZ "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "update linbez set kuerzel = '' , Marke = '' where linbezeich = '/////'"
        gdBase.Execute cSQL, dbFailOnError
        
        pbrLinbez.value = 100
        pbrLinbez.Visible = False
    End If
    
    cSQL = "update linbez set sorti = lpz where sorti is null"
    gdBase.Execute cSQL, dbFailOnError
    
    EinzelnLisrtNehmenWKL11
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 75 Then
        Resume Next
    ElseIf err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "AlleLinBezCSV"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Private Sub checkDieLinien()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String

    loeschNEW "LPZT", gdBase
        
    cSQL = "Create Table LPZT "
    cSQL = cSQL & "( LINR long"
    cSQL = cSQL & ", LINBEZEICH Text(30)"
    cSQL = cSQL & ", LPZ integer"
    cSQL = cSQL & ", MARKER Text(1) "
    cSQL = cSQL & ", KUERZEL Text(5) "
    cSQL = cSQL & ", MARKE Text(20) "
    cSQL = cSQL & ", SORTI integer" 'neu
    cSQL = cSQL & " )"
    gdBase.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("LINBEZ", gdBase) Then
        cSQL = "Insert into LPZT Select * from LinBez "
        cSQL = cSQL & " WHERE LINR BETWEEN 0 AND 499999 or LPZ < 599 "
        cSQL = cSQL & " and trim(LINBEZEICH) <> '/////'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkDieLinien"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub checkDieLinien2()
On Error GoTo LOKAL_ERROR

    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim rsRs2           As Recordset
    Dim cLinr           As String
    Dim clpz            As String
    Dim lcount          As Long
    Dim j               As Long
    Dim cdat            As String
    
    Screen.MousePointer = 11
    loeschNEW "LPZT2", gdBase
        
    cSQL = "Create Table LPZT2 "
    cSQL = cSQL & "( LINR long"
    cSQL = cSQL & ", LINBEZ Text(35)"
    cSQL = cSQL & ", LINBEZEICH Text(30)"
    cSQL = cSQL & ", LPZ integer"
    cSQL = cSQL & ", MARKER Text(1) "
    cSQL = cSQL & ", SORTI INTEGER "
    cSQL = cSQL & ", KUERZEL Text(5) "
    cSQL = cSQL & ", MARKE Text(20) "
    cSQL = cSQL & ", STATUS Text(30) "
    cSQL = cSQL & ", DATUM Text(10)) "
    gdBase.Execute cSQL, dbFailOnError
    
    CheckIndex "LPZT", "LINR", "", gdBase
    CheckIndex "LPZT", "LPZ", "", gdBase
    CheckIndex "LPZT", "LINBEZEICH", "", gdBase
    CheckIndex "LinBez", "LINBEZEICH", "", gdBase
    
    cSQL = "Select * from LPZT "
    cSQL = cSQL & " WHERE LINR = 300200 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LINBEZEICH) Then
            cdat = rsrs!LINBEZEICH
        Else
            cdat = ""
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'neue
    cSQL = "Select * from LinBez "
    cSQL = cSQL & " WHERE LINR BETWEEN 0 AND 499999 or LPZ < 599 "
    cSQL = cSQL & " and trim(LINBEZEICH) <> '/////'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            j = lcount Mod 100
            If j = 0 Then
                anzeige "normal", CStr(lcount), Label9
            Else
                
            End If
            lcount = lcount - 1
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            Else
                cLinr = "0"
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                clpz = rsrs!LPZ
            Else
                clpz = 0
            End If
            
            If cLinr <> "0" Then
                cSQL = "Select * from LPZT "
                cSQL = cSQL & " WHERE LINR = " & cLinr & " and LPZ = " & clpz
                Set rsRs2 = gdBase.OpenRecordset(cSQL)
                If Not rsRs2.EOF Then
                    
                Else
                    cSQL = " Insert into LPZT2 select *, 'neue Linien' as STATUS,'" & cdat & "'as Datum from LINBEZ"
                    cSQL = cSQL & " WHERE LINR = " & cLinr & " and LPZ = " & clpz
                    cSQL = cSQL & " and trim(LINBEZEICH) <> '/////'"
                    gdBase.Execute cSQL, dbFailOnError
                End If
                rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "LINALT", gdBase
    cSQL = "Select * into LINALT from LINBEZ "
    cSQL = cSQL & " WHERE  trim(LINBEZEICH) = '/////'"
    gdBase.Execute cSQL, dbFailOnError
    
    'alte
    cSQL = "Select * from LPZT "
    cSQL = cSQL & " where trim(LINBEZEICH) <> '/////'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        
            j = lcount Mod 100
            If j = 0 Then
                anzeige "normal", CStr(lcount), Label9
            Else
                
            End If
            
            lcount = lcount - 1
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            Else
                cLinr = "0"
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                clpz = rsrs!LPZ
            Else
                clpz = 0
            End If
            
            If cLinr <> "0" Then
                cSQL = "Select * from LINALT "
                cSQL = cSQL & " WHERE LINR = " & cLinr & " and LPZ = " & clpz
                cSQL = cSQL & " and trim(LINBEZEICH) = '/////'"
                Set rsRs2 = gdBase.OpenRecordset(cSQL)
                If Not rsRs2.EOF Then
                
                    cSQL = " Insert into LPZT2 select *, 'gelöschte Linien' as STATUS ,'" & cdat & "'as Datum from LPZT"
                    cSQL = cSQL & " WHERE LINR = " & cLinr & " and LPZ = " & clpz
                    gdBase.Execute cSQL, dbFailOnError
                
                End If
                rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Update LPZT2 inner join lisrt on LPZT2.linr = lisrt.linr "
    cSQL = cSQL & "set LPZT2.LINBEZ = lisrt.liefbez "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "LINALT", gdBase
    
    Screen.MousePointer = 0
    
    reportbildschirm "dWKL11a", "aWKL11s"

    anzeige "normal", "", Label9
           
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkDieLinien2"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub DruckeHinweiseStammdatenWKL11() 'X
    On Error GoTo LOKAL_ERROR
    
    setzedrucker gcListenDrucker

    Printer.FontBold = True
    Printer.FontSize = 9
    Printer.Font = "Courier new"
    Printer.Print Text4.Text
    Printer.EndDoc

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeHinweiseStammdatenWKL11"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim bPfad   As Boolean
    Dim lRet    As Long
    Dim cPfad   As String
    Dim ctmp    As String
    Dim sAnzeigetext As String
    Dim sSQL As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    gbEinstell = False
    bPfad = True
    
    WKL11Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label1(0)
    
    BTNSteuersenkung.BackColor = vbYellow
    
    loesch "MASTER2"
    Kill cPfad & "MASTER2.MDX"
    
    Frame1.Visible = True
    Frame7.Visible = False
    
    Label14.Caption = ""
    Label14.Refresh
    
    Text3(1).Text = ermNextWochendatei
    Text3(0).Text = ""
    
    Drive1.Drive = gcDBPfad & "\IN"
    Dir1.Path = gcDBPfad & "\IN"
    If Not bPfad Then
        MsgBox "Das Verzeichnis " & UCase$(gcDBPfad) & "IN existiert nicht!" & vbCrLf & vbCrLf & "Stammdaten-Update nicht möglich!", vbCritical, "STOP!"
        Command1(0).Enabled = False
        Exit Sub
    End If
    
   
    
    sAnzeigetext = "Wichtig!! Wochendatei(Master...) in der numerischen Reihenfolge einlesen!!"
    
    Label24.Caption = sAnzeigetext
    Label24.Refresh
    
    
    
    
    
    Kill gcDBPfad & "\IN\" & "M1*!.*"
        
    File1.Path = gcDBPfad & "\IN"
    If Not bPfad Then
        MsgBox "Das Verzeichnis " & UCase$(gcDBPfad) & "IN existiert nicht!" & vbCrLf & vbCrLf & "Stammdaten-Update nicht möglich!", vbCritical, "STOP!"
        Command1(0).Enabled = False
        Exit Sub
    End If
    
    If File1.ListCount > 0 Then
        File1.Selected(0) = True
        Label6.Caption = "Bitte wählen Sie eine Datei aus!"
        Label6.Refresh
    Else
        Label6.Caption = "Bitte wählen Sie eine Datei aus!"
        Label6.Refresh
    End If
    
    If gbOLDSTADADEL = True Then
        LoescheUrAltMasterWKL11
    End If
    
    If CInt(gcFilNr) = 1 Then
        ctmp = "Achtung: Lesen Sie Ihre Stammdaten bitte im Programm 'Zentrale' ein." & vbCrLf
        ctmp = ctmp & "Änderungen an Artikeln werden nicht an die Filialen weitergeleitet."
        
        Label9.ForeColor = vbRed
        Label9.Caption = ctmp
        Label9.Refresh
    Else
        Label9.ForeColor = vbBlue
        Label9.Caption = ""
        Label9.Refresh
    End If
    
    If gbFtpYes Then
        LeseLIZENZ
        If gbLizenz Then
        
        
            If gbNOWOCHENDATEN = True Then
        
                Label26(0).Caption = "Wochendaten deaktiviert"
                Label26(0).ForeColor = vbRed
                Label26(0).Visible = True
            
                Command1(4).Enabled = False
                Command1(6).Enabled = True
                Command1(8).Enabled = True
            Else
                Label26(0).Caption = "Lizenz liegt vor"
                Label26(0).ForeColor = glS1
                Label26(0).Visible = True
                
                Command1(4).Enabled = True
                Command1(6).Enabled = True
                Command1(8).Enabled = True
            End If
        Else
            Label26(0).Caption = "keine Lizenz"
            Label26(0).ForeColor = vbRed
            Label26(0).Visible = True
            
            Command1(4).Enabled = False
            Command1(6).Enabled = False
            Command1(8).Enabled = False
        End If
    
    Else
        Label26(0).Caption = "keine FTP Einstellungen"
        Label26(0).ForeColor = vbRed
        Label26(0).Visible = True
    End If
    
    If NewTableSuchenDBKombi("E11C", gdApp) Then
    
    
    
        If SpalteInTabellegefundenNEW("E11C", "LVKAUFSCHLAG", gdApp) = False Then
            SpalteAnfuegenNEW "E11C", "LVKAUFSCHLAG", "Text(10)", gdApp

            sSQL = "Update E11C set LVKAUFSCHLAG = '' "
            gdApp.Execute sSQL, dbFailOnError

        End If
        
        If SpalteInTabellegefundenNEW("E11C", "Bo11", gdApp) = False Then
            SpalteAnfuegenNEW "E11C", "Bo11", "BIT", gdApp

            sSQL = "Update E11C set bo11 = -1 "
            gdApp.Execute sSQL, dbFailOnError
        End If
    
    
    
        If SpalteInTabellegefundenNEW("E11C", "BO10", gdApp) Then
            voreinstellungladenE11C
        End If
    End If
    
    sAnzeigetext = "Aktualität der Stammdaten" & vbCrLf
    sAnzeigetext = sAnzeigetext & "Sie können sich unter Wochendaten einmal wöchentlich alle Artikeländerungen und Artikelneuheiten der letzten Woche abrufen." & vbCrLf
    sAnzeigetext = sAnzeigetext & "Haben Sie Änderungswünsche zu den Artikeldaten, so rufen Sie uns bitte unter 0511/955910 an oder schicken Sie uns ein Fax unter 0511/9559144, wir nehmen dann Einzeländerungen sofort vor und stellen Ihnen die Daten unter Tagesdaten gegen 18:00 Uhr des selben Tages bereit." & vbCrLf
    sAnzeigetext = sAnzeigetext & "Unter Lieferantendaten können Sie alle Artikeldaten eines bestimmten Lieferanten abrufen."
    Label26(1).Caption = sAnzeigetext
    Label26(1).Refresh
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 76 Then
        bPfad = False
        Resume Next
    ElseIf err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Load"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub voreinstellungspeichernE11C()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim bo0     As Integer
    Dim bo1     As Integer
    Dim bo2     As Integer
    Dim bo3     As Integer
    Dim bo4     As Integer
    Dim bo5     As Integer
    Dim bo6     As Integer
    Dim bo7     As Integer
    Dim bo8     As Integer
    Dim bo9     As Integer
    Dim bo10    As Integer
    Dim bo11    As Integer
   
    loeschNEW "E11C", gdApp
    CreateTable "E11C", gdApp
    
    If Check1(0).value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    If Check1(1).value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    If Check1(2).value = vbChecked Then
        bo2 = 0
    Else
        bo2 = -1
    End If
    
    If Check1(3).value = vbChecked Then
        bo3 = 0
    Else
        bo3 = -1
    End If
    
    If Check1(4).value = vbChecked Then
        bo4 = 0
    Else
        bo4 = -1
    End If
    
    If Check1(5).value = vbChecked Then
        bo5 = 0
    Else
        bo5 = -1
    End If
    
    If Check1(6).value = vbChecked Then
        bo6 = 0
    Else
        bo6 = -1
    End If
    
    If Check1(7).value = vbChecked Then
        bo7 = 0
    Else
        bo7 = -1
    End If
    
    If Check1(8).value = vbChecked Then
        bo8 = 0
    Else
        bo8 = -1
    End If
    
    If Check1(9).value = vbChecked Then
        bo9 = 0
    Else
        bo9 = -1
    End If
    
    If Check1(10).value = vbChecked Then
        bo10 = 0
    Else
        bo10 = -1
    End If
    
    If Check2.value = vbChecked Then
        bo11 = 0
    Else
        bo11 = -1
    End If
    
    
   
    sSQL = "Insert into E11C ( bo0,bo1,bo2,bo3,bo4,bo5,bo6,bo7,bo8,bo9,bo10,bo11,LVKAUFSCHLAG) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4 & " "
    sSQL = sSQL & "," & bo5 & "," & bo6 & "," & bo7 & "," & bo8 & "," & bo9 & "," & bo10 & "," & bo11 & ",'" & Text2(4).Text & "')"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE11C"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE11C()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdApp.OpenRecordset("E11C")
    If Not rs.EOF Then
    
        If rs!bo0 = True Then
            Check1(0).value = vbUnchecked
        Else
            Check1(0).value = vbChecked
        End If
        
        If rs!bo1 = True Then
            Check1(1).value = vbUnchecked
        Else
            Check1(1).value = vbChecked
        End If
        
        If rs!bo2 = True Then
            Check1(2).value = vbUnchecked
        Else
            Check1(2).value = vbChecked
        End If
        
        If rs!bo3 = True Then
            Check1(3).value = vbUnchecked
        Else
            Check1(3).value = vbChecked
        End If
        
        If rs!bo4 = True Then
            Check1(4).value = vbUnchecked
        Else
            Check1(4).value = vbChecked
        End If
        
        If rs!bo5 = True Then
            Check1(5).value = vbUnchecked
        Else
            Check1(5).value = vbChecked
        End If
        If rs!bo6 = True Then
            Check1(6).value = vbUnchecked
        Else
            Check1(6).value = vbChecked
        End If
        
        If rs!bo7 = True Then
            Check1(7).value = vbUnchecked
        Else
            Check1(7).value = vbChecked
        End If
        If rs!bo8 = True Then
            Check1(8).value = vbUnchecked
        Else
            Check1(8).value = vbChecked
        End If
        
        If rs!bo9 = True Then
            Check1(9).value = vbUnchecked
        Else
            Check1(9).value = vbChecked
        End If
        
        If rs!bo10 = True Then
            Check1(10).value = vbUnchecked
        Else
            Check1(10).value = vbChecked
        End If
        
        
        
        If Not IsNull(rs!LVKAUFSCHLAG) Then
            Text2(4).Text = rs!LVKAUFSCHLAG
        Else
            Text2(4).Text = ""
        End If
        
        If rs!bo11 = True Then
            Check2.value = vbUnchecked
        Else
            Check2.value = vbChecked
        End If
    
    End If
    rs.Close: Set rs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE11C"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function BISTDUPreisschutz(cArtNr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    BISTDUPreisschutz = True
    
    Dim sSQL As String
    Dim rsrs As Recordset
    If cArtNr <> "" Then
        If IsNumeric(cArtNr) Then
            sSQL = "Select * from Artikel where artnr  = " & cArtNr & " and Preisschu =  'J' "
            
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                BISTDUPreisschutz = False
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    End If
   
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BISTDUPreisschutz"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub KalkandRund(sArt As String)
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dKVKN           As Double
    Dim dLEKPR          As Double
    Dim dVkPr           As Double
    Dim dFaktor         As Double
    Dim dAufschlag      As Double
    Dim cLinr As String
    Dim cArtNr As String
    Dim lpreisszahler As Long
    
    MSFlexGrid2.Redraw = False
    
    lpreisszahler = 0
    Select Case sArt
        Case "RUNDEN"
    
            MSFlexGrid2.Row = 0
            For i = 1 To MSFlexGrid2.Rows - 1
                MSFlexGrid2.Row = i
                MSFlexGrid2.Col = SpaltennummerKVKNEU
                
                If Not Len(MSFlexGrid2.Text) = 0 Then
                    
                    dKVKN = MSFlexGrid2.Text
                    If dKVKN <> 0 Then
                        MSFlexGrid2.Text = Runden(dKVKN)
                    End If
                Else
                    MSFlexGrid2.Text = "0,00"
                End If
            Next i
            
            MSFlexGrid2.Refresh
    
        Case Is = "LISTENEK"
        
            If IsNumeric(Text2(0).Text) = True Then
                dFaktor = CDbl(Text2(0).Text)
                
                MSFlexGrid2.Row = 0
                For i = 1 To MSFlexGrid2.Rows - 1
                
                    MSFlexGrid2.Row = i
                    MSFlexGrid2.Col = SpaltennummerLINR
                    cLinr = MSFlexGrid2.Text
                    
                    If BISTDUINNOKALKL(cLinr) = False Then
                    
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerArtnr
                        cArtNr = MSFlexGrid2.Text
                        
                        If BISTDUPreisschutz(cArtNr) = True Then
                            MSFlexGrid2.Row = i
                            MSFlexGrid2.Col = SpaltennummerLEKPR
                            If Not Len(MSFlexGrid2.Text) = 0 Then
                                
                                If IsNumeric(MSFlexGrid2.Text) Then
                                    dLEKPR = CDbl(MSFlexGrid2.Text)
                                
                                    
                                    If dLEKPR <> 0 Then
                                    
                                        dKVKN = (dLEKPR * dFaktor)
                                        MSFlexGrid2.Col = SpaltennummerKVKNEU
                                        MSFlexGrid2.Text = Runden(dKVKN)

                                    End If
                                End If
                            Else
                                MSFlexGrid2.Col = SpaltennummerKVKNEU
                                MSFlexGrid2.Text = "0,00"
                            End If
                        Else
                            lpreisszahler = lpreisszahler + 1
                            MSFlexGrid2.Col = SpaltennummerKVKNEU
                            MSFlexGrid2.CellFontItalic = True
                            MSFlexGrid2.CellForeColor = vbRed
                        End If
                    Else 'wenigstens runden
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerKVKNEU
                        
                        If Not Len(MSFlexGrid2.Text) = 0 Then
                            
                            dKVKN = MSFlexGrid2.Text
                            If dKVKN <> 0 Then
                                MSFlexGrid2.Text = Runden(dKVKN)
                            End If
                        Else
                            MSFlexGrid2.Text = "0,00"
                        End If
                    
                    
                    End If
                Next i
                
                MSFlexGrid2.Refresh
            End If
            
        Case Is = "LISTENVK"
        
            If IsNumeric(Text2(3).Text) = True Then
                dAufschlag = CDbl(Text2(3).Text)
                
                
            
                MSFlexGrid2.Row = 0
                For i = 1 To MSFlexGrid2.Rows - 1
                
                    MSFlexGrid2.Row = i
                    MSFlexGrid2.Col = SpaltennummerLINR
                    cLinr = MSFlexGrid2.Text
                    
                    If BISTDUINNOKALKL(cLinr) = False Then
                        
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerArtnr
                        cArtNr = MSFlexGrid2.Text
                        
                        If BISTDUPreisschutz(cArtNr) = True Then
                            MSFlexGrid2.Row = i
                            MSFlexGrid2.Col = SpaltennummerVKPR
                            If Not Len(MSFlexGrid2.Text) = 0 Then
                                
                                If IsNumeric(MSFlexGrid2.Text) Then
                                    dVkPr = CDbl(MSFlexGrid2.Text)
                                
                                    If dVkPr <> 0 Then
                                        If Check6.value = vbChecked Then
                                            If dVkPr > 20 Then
                                                dKVKN = dVkPr + ((dVkPr * dAufschlag) / 100)
                                                MSFlexGrid2.Col = SpaltennummerKVKNEU
                                                MSFlexGrid2.Text = Runden(dKVKN)
                                            Else
                                                dKVKN = dVkPr
                                            End If
                                        Else
                                            dKVKN = dVkPr + ((dVkPr * dAufschlag) / 100)
                                            MSFlexGrid2.Col = SpaltennummerKVKNEU
                                            MSFlexGrid2.Text = Runden(dKVKN)
                                        End If
                                    End If
                                End If
                            Else
                                MSFlexGrid2.Col = SpaltennummerKVKNEU
                                MSFlexGrid2.Text = "0,00"
                            End If
                        Else
                            lpreisszahler = lpreisszahler + 1
                            MSFlexGrid2.Col = SpaltennummerKVKNEU
                            MSFlexGrid2.CellFontItalic = True
                            MSFlexGrid2.CellForeColor = vbRed
                        End If
                    Else 'wenigstens runden
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerVKPR
                        
                        If Not Len(MSFlexGrid2.Text) = 0 Then
                            
                            dKVKN = MSFlexGrid2.Text
                            If dKVKN <> 0 Then

                                MSFlexGrid2.Text = Runden(dKVKN)

                            End If
                        Else
                            MSFlexGrid2.Text = "0,00"
                        End If
                    
                    
                    End If
                Next i
                
                MSFlexGrid2.Refresh
            End If
    
    End Select
    
    MSFlexGrid2.Redraw = True
    
    If lpreisszahler > 0 Then
        anzeige "rot", lpreisszahler & " x Preisschutz ", Label22(2)
    Else
        anzeige "normal", "", Label22(2)
    End If
    

    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KalkandRund"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub nurKalk(sArt As String)
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dKVKN           As Double
    Dim dLEKPR          As Double
    Dim dVkPr           As Double
    Dim dFaktor         As Double
    Dim dAufschlag      As Double
    Dim cLinr As String
    Dim cArtNr As String
    Dim lpreisszahler As Long
    
    MSFlexGrid2.Redraw = False
    
    lpreisszahler = 0
    Select Case sArt
        Case "RUNDEN"
    
            MSFlexGrid2.Row = 0
            For i = 1 To MSFlexGrid2.Rows - 1
                MSFlexGrid2.Row = i
                MSFlexGrid2.Col = SpaltennummerKVKNEU
                
                If Not Len(MSFlexGrid2.Text) = 0 Then
                    
                    dKVKN = MSFlexGrid2.Text
                    If dKVKN <> 0 Then
                        MSFlexGrid2.Text = Format$(dKVKN, "####.00")
                    End If
                Else
                    MSFlexGrid2.Text = "0,00"
                End If
            Next i
            MSFlexGrid2.Refresh
        Case Is = "LISTENEK"
            If IsNumeric(Text2(0).Text) = True Then
                dFaktor = CDbl(Text2(0).Text)
                
                MSFlexGrid2.Row = 0
                For i = 1 To MSFlexGrid2.Rows - 1
                    MSFlexGrid2.Row = i
                    MSFlexGrid2.Col = SpaltennummerLINR
                    cLinr = MSFlexGrid2.Text
                    
                    If BISTDUINNOKALKL(cLinr) = False Then
                    
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerArtnr
                        cArtNr = MSFlexGrid2.Text
                        
                        If BISTDUPreisschutz(cArtNr) = True Then
                            MSFlexGrid2.Row = i
                            MSFlexGrid2.Col = SpaltennummerLEKPR
                            If Not Len(MSFlexGrid2.Text) = 0 Then
                                
                                If IsNumeric(MSFlexGrid2.Text) Then
                                    dLEKPR = CDbl(MSFlexGrid2.Text)
                                    If dLEKPR <> 0 Then
                                        dKVKN = (dLEKPR * dFaktor)

                                        MSFlexGrid2.Col = SpaltennummerKVKNEU
                                        MSFlexGrid2.Text = Format$(dKVKN, "####.00")
                                    End If
                                End If
                            Else
                                MSFlexGrid2.Col = SpaltennummerKVKNEU
                                MSFlexGrid2.Text = "0,00"
                            End If
                        Else
                            lpreisszahler = lpreisszahler + 1
                            MSFlexGrid2.Col = SpaltennummerKVKNEU
                            MSFlexGrid2.CellFontItalic = True
                            MSFlexGrid2.CellForeColor = vbRed
                        End If
                    Else 'wenigstens runden
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerKVKNEU
                        
                        If Not Len(MSFlexGrid2.Text) = 0 Then
                            dKVKN = MSFlexGrid2.Text
                            If dKVKN <> 0 Then
                                MSFlexGrid2.Text = Format$(dKVKN, "####.00")
                            End If
                        Else
                            MSFlexGrid2.Text = "0,00"
                        End If
                    End If
                Next i
                
                MSFlexGrid2.Refresh
            End If
            
        Case Is = "LISTENVK"
        
            If IsNumeric(Text2(3).Text) = True Then
                dAufschlag = CDbl(Text2(3).Text)
                
                
            
                MSFlexGrid2.Row = 0
                For i = 1 To MSFlexGrid2.Rows - 1
                
                    MSFlexGrid2.Row = i
                    MSFlexGrid2.Col = SpaltennummerLINR
                    cLinr = MSFlexGrid2.Text
                    
                    If BISTDUINNOKALKL(cLinr) = False Then
                        
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerArtnr
                        cArtNr = MSFlexGrid2.Text
                        
                        If BISTDUPreisschutz(cArtNr) = True Then
                            MSFlexGrid2.Row = i
                            MSFlexGrid2.Col = SpaltennummerVKPR
                            If Not Len(MSFlexGrid2.Text) = 0 Then
                                
                                If IsNumeric(MSFlexGrid2.Text) Then
                                    dVkPr = CDbl(MSFlexGrid2.Text)
                                
                                    If dVkPr <> 0 Then
                                    
                                        dKVKN = dVkPr + ((dVkPr * dAufschlag) / 100)
                                        MSFlexGrid2.Col = SpaltennummerKVKNEU
                                        MSFlexGrid2.Text = Format$(dKVKN, "####.00")

                                    End If
                                End If
                            Else
                                MSFlexGrid2.Col = SpaltennummerKVKNEU
                                MSFlexGrid2.Text = "0,00"
                            End If
                        Else
                            lpreisszahler = lpreisszahler + 1
                            MSFlexGrid2.Col = SpaltennummerKVKNEU
                            MSFlexGrid2.CellFontItalic = True
                            MSFlexGrid2.CellForeColor = vbRed
                        End If
                    Else 'wenigstens runden
                        MSFlexGrid2.Row = i
                        MSFlexGrid2.Col = SpaltennummerVKPR
                        
                        If Not Len(MSFlexGrid2.Text) = 0 Then
                            
                            dKVKN = MSFlexGrid2.Text
                            If dKVKN <> 0 Then

                                MSFlexGrid2.Text = Format$(dKVKN, "####.00")

                            End If
                        Else
                            MSFlexGrid2.Text = "0,00"
                        End If
                    
                    
                    End If
                Next i
                
                MSFlexGrid2.Refresh
            End If
    
    End Select
    
    MSFlexGrid2.Redraw = True
    
    If lpreisszahler > 0 Then
        anzeige "rot", lpreisszahler & " x Preisschutz ", Label22(2)
    Else
        anzeige "normal", "", Label22(2)
    End If
    

    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "nurKalk"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Function BISTDUINNOKALKL(cLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    BISTDUINNOKALKL = False
    
    Dim sSQL As String
    Dim rsrs As Recordset
    If cLinr <> "" Then
        If IsNumeric(cLinr) Then
            sSQL = "Select * from NOKALKL where  linr  = " & cLinr
            
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                BISTDUINNOKALKL = True
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    End If
   
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BISTDUINNOKALKL"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub FlexGrid_Update(oGrid As MSFlexGrid)
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
    
    
    If nRow > nRowSel Then
        lBig = nRow
        nDelRow = nRowSel - 1
    Else
        lBig = nRowSel
        nDelRow = nRow - 1
    End If
    
    
    Do While nDelRow < lBig
    
        nDelRow = nDelRow + 1
        
        If nDelRow > 1 Then
            .TextMatrix(nDelRow, SpaltennummerArtnr) = "entfernt"
        End If
       
    
    
    Loop
    
    
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_SelChange()
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    
    lrow = MSFlexGrid2.Row
    If lrow > 1 Then
        Label14.Caption = MSFlexGrid2.TextMatrix(lrow, SpaltennummerBEZEICH)
        Label14.Refresh
        
        Label15.Caption = ermittleLiefbez(CLng(MSFlexGrid2.TextMatrix(lrow, SpaltennummerLINR)))
        Label15.Refresh
    End If

    Exit Sub
LOKAL_ERROR:
    

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "MSFlexGrid2_SelChange"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Stammdaten einlesen auf."
    
        Fehlermeldung1
End Sub
Private Function ermittleLiefbez(linr As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Dim rs      As Recordset
    Dim sSQL    As String
    
    ermittleLiefbez = ""
    
    sSQL = "Select * from lisrt where linr = " & linr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!LIEFBEZ) Then
            ermittleLiefbez = rs!LIEFBEZ
        End If
    Else
        ermittleLiefbez = ""
    End If
    rs.Close: Set rs = Nothing
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleLiefbez"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Private Sub Text2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sAuswahlfeld As String
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case index
            Case Is = 1
                gF2Prompt.cFeld = "LINR"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text2(index).Text = gF2Prompt.cWahl
                    End If
                End If
        End Select
        Text2(index).SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR

    Text2(index).BackColor = glSelBack1
    Text2(index).SelStart = 0
    Text2(index).SelLength = Len(Text2(index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR

    Text3(index).BackColor = glSelBack1
    Text3(index).SelStart = 0
    Text3(index).SelLength = Len(Text3(index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case index
        Case 1 'liNR
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 0, 3, 2, 4 'kalk ,aufschlag in proz auf listenvk, auch ek abschlag
            cValid = "1234567890,-" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If

    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer, index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    cZeichen = Chr$(KeyAscii)
    KeyAscii = Asc(cZeichen)
    
    
    Select Case index
    
        Case 0, 1
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text2(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text3(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer, index As Integer)
On Error GoTo LOKAL_ERROR

Select Case index
    Case 0
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            
            If gF2Prompt.cWahl <> "" Then
                Text3(0).Text = gF2Prompt.cWahl
            End If
            Text3(0).SetFocus
            Command1_Click 8
        
        End If
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TxtRunden_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sProz           As Single
    Dim i               As Integer
    Dim dZahl           As Double
    
    
    If KeyCode = vbKeyF2 Then
        
        sProz = InputBox("Geben Sie bitte die Abschlagspanne ein!", "Winkiss Eingabe")
        
        MSFlexGrid2.Row = 0
        
        For i = 1 To MSFlexGrid2.Rows - 1
            MSFlexGrid2.Row = i
            MSFlexGrid2.Col = 6
            
            
            dZahl = MSFlexGrid2.Text
            dZahl = dZahl - (dZahl * sProz / 100)
            MSFlexGrid2.Col = 8
            
            MSFlexGrid2.Text = Format$(dZahl, "###,##0.00")
        Next i
        
        MSFlexGrid2.Refresh
        
   
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "TxtRunden_KeyUp"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        
    End If
End Sub
