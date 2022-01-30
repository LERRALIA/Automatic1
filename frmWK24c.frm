VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmWK24c 
   BackColor       =   &H00C0C000&
   Caption         =   "WINKISS automatische Rechnungserstellung"
   ClientHeight    =   8610
   ClientLeft      =   1215
   ClientTop       =   1590
   ClientWidth     =   11910
   Icon            =   "frmWK24c.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Starte mit Rechnungsnummer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   40
      TabIndex        =   22
      Top             =   3960
      Width           =   3495
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   240
         Width           =   1575
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "nächste ReNr"
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopien je Rechn:"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
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
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "höchste ReNr.:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Zusätze für ALLE Rechnungen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   40
      TabIndex        =   19
      Top             =   5640
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   1455
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   16
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   9
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "allgemeine Text-Hinweise"
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
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Pauschale für Porto / Verpackung:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "zu druckende Rechnung(en)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   3640
      TabIndex        =   17
      Top             =   0
      Width           =   8175
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
         Height          =   7155
         Left            =   120
         MultiSelect     =   1  '1 -Einfach
         TabIndex        =   18
         Top             =   240
         Width           =   7935
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   6600
         TabIndex        =   29
         Top             =   7680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "Schließen"
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   7680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "Drucke markierte Rechnungen"
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   3285
         TabIndex        =   31
         Top             =   7680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "Drucke alle Rechnungen"
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Vorgaben"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   28
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "Suche"
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, zzgl. MWSt"
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
         TabIndex        =   26
         Top             =   3480
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, ohne MWSt"
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
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "auf Netto berechnen, zzgl. MWSt"
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
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, inkl. MWSt"
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
         TabIndex        =   10
         Top             =   2760
         Value           =   -1  'True
         Width           =   3255
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   360
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   32
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "alle"
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "dieser Monat"
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         MousePointer    =   99
         BorderStyle     =   2
         ButtonStyle     =   2
         Caption         =   "letzter Monat"
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelpreise in Rechnung"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kaufdatum bis:"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kaufdatum von:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KdNr bis:"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KdNr von:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmWK24c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bsofortanDrucker As Boolean
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
Private Function fnPruefeEingabeDialogWK24c() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    
    fnPruefeEingabeDialogWK24c = 0
    
    cFeld = MaskEdBox1(0).Text
    If cFeld = "______" Then
        MaskEdBox1(0).Text = "1_____"
    End If
    
    cFeld = MaskEdBox1(1).Text
    If cFeld = "______" Then
        MaskEdBox1(1).Text = "999999"
    End If
    
    cFeld = MaskEdBox1(2).Text
    If cFeld = "__.__.____" Then
        MaskEdBox1(2).Text = "01.01.2000"
    Else
        If Not IsDate(cFeld) Then
            fnPruefeEingabeDialogWK24c = 3
            Exit Function
        End If
    End If
    
    cFeld = MaskEdBox1(3).Text
    If cFeld = "__.__.____" Then
        MaskEdBox1(3).Text = "31.12.2040"
    Else
        If Not IsDate(cFeld) Then
            fnPruefeEingabeDialogWK24c = 4
            Exit Function
        End If
    End If
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK24"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function fnHoleMaxReNrWK24c() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cReNr As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select max(val(SCHLUESSEL)) as MAXRENU from REKOPF "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!MAXRENU) Then
            cReNr = rsrs!MAXRENU
        Else
            cReNr = "0"
        End If
    Else
        cReNr = "0"
    End If
    rsrs.Close: Set rsrs = Nothing

    fnHoleMaxReNrWK24c = cReNr

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleMaxReNrWK24c"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub SucheKreditVerkaeufeWK24c()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    
    Dim lKdNrMin As Long
    Dim lKdNrMax As Long
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    
    Dim lKUNDNR As Long
    Dim dBetrag As Double
    Dim cKdName As String
    Dim cKdVorname As String
    Dim cKdOrt As String
    
    Dim ctmp As String
    Dim cLBSatz As String
    
    List1.Clear
    
    cFeld = MaskEdBox1(0).Text
    lKdNrMin = Val(cFeld)
    
    cFeld = MaskEdBox1(1).Text
    lKdNrMax = Val(cFeld)
    
    cFeld = MaskEdBox1(2).Text
    lDatVon = DateValue(cFeld)
    
    cFeld = MaskEdBox1(3).Text
    lDatBis = DateValue(cFeld)
    
    cSQL = "Select distinct KUNDNR from KREDIT "
    cSQL = cSQL & "where KUNDNR >= " & Trim$(Str$(lKdNrMin)) & " "
    cSQL = cSQL & "and KUNDNR <= " & Trim$(Str$(lKdNrMax)) & " "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kundnr) Then
                lKUNDNR = rsrs!Kundnr
            Else
                lKUNDNR = 0
            End If
            If lKUNDNR > 0 Then
                cSQL = "Select SUM(GVKPR) as BETRAG from KREDIT "
                cSQL = cSQL & "where KUNDNR = " & Trim$(Str$(lKUNDNR)) & " "
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lDatVon)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lDatBis)) & " "
                        
                Set rsRs2 = gdBase.OpenRecordset(cSQL)
                If Not rsRs2.EOF Then
                    rsRs2.MoveFirst
                    If Not IsNull(rsRs2!Betrag) Then
                        dBetrag = rsRs2!Betrag
                    Else
                        dBetrag = 0
                    End If
                Else
                    dBetrag = 0
                End If
                rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
                
                If dBetrag <> 0 Then
                    Set rsRs2 = gdBase.OpenRecordset("select * from Kunden where Kundnr = " & lKUNDNR)
                   
                    If Not rsRs2.EOF Then
                        If Not IsNull(rsRs2!name) Then
                            cKdName = rsRs2!name
                        Else
                            cKdName = ""
                        End If
                        If Not IsNull(rsRs2!vorname) Then
                            cKdVorname = rsRs2!vorname
                        Else
                            cKdVorname = ""
                        End If
                        If Not IsNull(rsRs2!STADT) Then
                            cKdOrt = rsRs2!STADT
                        Else
                            cKdOrt = ""
                        End If
                    End If
                    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
                End If
            End If
            
            ctmp = Trim$(Str$(lKUNDNR))
            ctmp = Space$(6 - Len(ctmp)) & ctmp
            cLBSatz = ctmp & " "
            
            cKdName = Trim$(cKdName)
            If cKdName <> "" Then
                ctmp = cKdName
            End If
            
            cKdVorname = Trim$(cKdVorname)
            If cKdVorname <> "" Then
                ctmp = ctmp & ", " & cKdVorname
            End If
            
            cKdOrt = Trim$(cKdOrt)
            If cKdOrt <> "" Then
                ctmp = ctmp & " / " & cKdOrt
            End If
            
            If Len(ctmp) > 55 Then
                ctmp = Left(ctmp, 55)
            Else
                ctmp = ctmp & Space$(55 - Len(ctmp))
            End If
            
            cLBSatz = cLBSatz & ctmp & " "
            
            ctmp = Format$(dBetrag, "#####0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & ctmp
            
            If dBetrag <> 0 Then
                List1.AddItem cLBSatz
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
    Fehler.gsFunktion = "SucheKreditVerkaeufeWK24c"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeRechnungenWK24c(lAnz As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sCrystaldat As String
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim ctmp As String
    Dim cKdnr As String
    Dim cSQL  As String
    Dim cInto As String
    Dim cIntoOFPO As String
    Dim rsrs As Recordset
    Dim lCopies As Long
    Dim lReDatum As Long
    Dim dSumme As Double
    Dim lcount As Long
    Dim lPos As Long
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim lrenr As Long
    Dim cZeichen As String
        
    'Zielfelder
    Dim cTitel As String
    Dim cAnrede As String
    Dim cAnredeTitel As String
    Dim cKdName1 As String
    Dim cKdName2 As String
    Dim cStrasse As String
    Dim cPlz As String
    Dim cOrt As String
    Dim cReNr As String
    Dim cReDatum As String
    Dim cReText As String
    Dim dReSumme As Double
    Dim dPortoVerp As Double
    Dim cKommentar As String
    Dim cPreisKz As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
        
    lDatVon = DateValue(MaskEdBox1(2).Text)
    lDatBis = DateValue(MaskEdBox1(3).Text)
        
    If Option1(0).Value = True Then
        gcRePreisKz = "B"
    ElseIf Option1(1).Value = True Then
        gcRePreisKz = "N"
    ElseIf Option1(2).Value = True Then
        gcRePreisKz = "O"
    ElseIf Option1(3).Value = True Then
        gcRePreisKz = "Z"
    End If
        
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            
            'Druck-Datei löschen und neu erzeugen
            loeschNEW "DRU_REKO", gdBase
            cSQL = "Create Table DRU_REKO "
            cSQL = cSQL & "(SCHLUESSEL Text(15)"
            cSQL = cSQL & ", KUNDNR LONG"
            cSQL = cSQL & ", ANREDE Text(35)"
            cSQL = cSQL & ", KDNAME1 Text(71)"
            cSQL = cSQL & ", KDNAME2 Text(71)"
            cSQL = cSQL & ", STRASSE Text(35)"
            cSQL = cSQL & ", PLZ Text(7)"
            cSQL = cSQL & ", ORT Text(35)"
            cSQL = cSQL & ", RENR Text(15)"
            cSQL = cSQL & ", REDATUM Datetime"
            cSQL = cSQL & ", RETEXT Text(100)"
            cSQL = cSQL & ", RESUMME double"
            cSQL = cSQL & ", STATUS Text(1)"
            cSQL = cSQL & ", PORTOVERP double"
            cSQL = cSQL & ", KOMMENTAR memo"
            cSQL = cSQL & ", PREISKZ Text(1)"
            cSQL = cSQL & ", STEUERNR Text(35)"
            
            cSQL = cSQL & ") "
            gdBase.Execute cSQL, dbFailOnError
            
            
            '*** Kundendaten lesen ***
                
            cTitel = ""
            cKdName1 = ""
            cKdName2 = ""
            cStrasse = ""
            cPlz = ""
            cOrt = ""
                
            ctmp = Text1.Text
            If ctmp = "" Then
                ctmp = "0"
            End If
            ctmp = fnMoveComma2Point$(ctmp)
            dPortoVerp = Val(ctmp)
            
            cKommentar = Text2.Text
        
            cKdnr = Trim$(Left(List1.list(lcount), 6))
            
            cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr
            Set rsrs = gdBase.OpenRecordset(cSQL)
            
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                
                If Not IsNull(rsrs!anrede) Then
                    cAnrede = rsrs!anrede
                Else
                    cAnrede = ""
                End If
                
                If Not IsNull(rsrs!titel) Then
                    cTitel = rsrs!titel
                Else
                    cTitel = ""
                End If
                
                cAnredeTitel = ""
                If cAnrede = "" Then
                    cAnredeTitel = cTitel
                Else
                    cAnredeTitel = cAnrede & Space(1) & cTitel
                End If
                
                If Not IsNull(rsrs!vorname) Then
                    cKdName1 = rsrs!vorname
                Else
                    cKdName1 = ""
                End If
            
                If Not IsNull(rsrs!name) Then
                    If cKdName1 <> "" Then
                        cKdName1 = cKdName1 & " " & rsrs!name
                    Else
                        cKdName1 = rsrs!name
                    End If
                Else
                    cKdName1 = cKdName1
                End If
                
                If Not IsNull(rsrs!firma) Then
                    cKdName2 = rsrs!firma
                Else
                    cKdName2 = ""
                End If
            
                If Not IsNull(rsrs!strasse) Then
                    cStrasse = rsrs!strasse
                Else
                    cStrasse = ""
                End If
            
                If Not IsNull(rsrs!Plz) Then
                    cPlz = rsrs!Plz
                Else
                    cPlz = ""
                End If
            
                If Not IsNull(rsrs!STADT) Then
                    cOrt = rsrs!STADT
                Else
                    cOrt = ""
                End If
            
            Else
            
            End If
            rsrs.Close: Set rsrs = Nothing
            
            lReDatum = Fix(Now)
            cReDatum = Trim$(Str$(lReDatum))
            
            '*** aktuelle Rechnungsnummer holen ***
            
            cReNr = Text3.Text
            lrenr = Val(cReNr)
            Text3.Text = Format$(lrenr, "##############0")
            cReNr = Text3.Text
            gcReNr = cReNr
                        
            cInto = "Insert into REKOPF "
            cIntoOFPO = "Insert into OFPO "
            cSQL = "( "
            cSQL = cSQL & "SCHLUESSEL"
            cSQL = cSQL & ", KUNDNR"
            cSQL = cSQL & ", ANREDE"
            cSQL = cSQL & ", KDNAME1"
            cSQL = cSQL & ", KDNAME2"
            cSQL = cSQL & ", STRASSE"
            cSQL = cSQL & ", PLZ"
            cSQL = cSQL & ", ORT"
            cSQL = cSQL & ", RENR"
            cSQL = cSQL & ", REDATUM"
            cSQL = cSQL & ", RETEXT"
            cSQL = cSQL & ", RESUMME"
            cSQL = cSQL & ", STATUS"
            cSQL = cSQL & ", PORTOVERP"
            cSQL = cSQL & ", KOMMENTAR"
            cSQL = cSQL & ", PREISKZ"
            cSQL = cSQL & ") values ("
            cSQL = cSQL & "'" & cReNr & "'"
            cSQL = cSQL & ", " & cKdnr & ""
            cSQL = cSQL & ", '" & cAnredeTitel & "'"
            cSQL = cSQL & ", '" & cKdName1 & "'"
            cSQL = cSQL & ", '" & cKdName2 & "'"
            cSQL = cSQL & ", '" & cStrasse & "'"
            cSQL = cSQL & ", '" & cPlz & "'"
            cSQL = cSQL & ", '" & cOrt & "'"
            cSQL = cSQL & ", '" & cReNr & "'"
            cSQL = cSQL & ", " & cReDatum & ""
            cSQL = cSQL & ", '" & cReText & "'"
            cSQL = cSQL & ", 0 "
            cSQL = cSQL & ", 'O' "
            cSQL = cSQL & ", " & Trim$(Str$(dPortoVerp)) & ""
            cSQL = cSQL & ", '" & cKommentar & "'"
            cSQL = cSQL & ", '" & gcRePreisKz & "'"
            cSQL = cSQL & ") "
            gdBase.Execute cInto & cSQL, dbFailOnError
            
            'OFPO auch, offene Postenliste
            gdBase.Execute cIntoOFPO & cSQL, dbFailOnError
            
            cInto = "Insert into DRU_REKO "
            gdBase.Execute cInto & cSQL, dbFailOnError
        
            dSumme = 0
                
            MoveDaten2RechnungWK24c cKdnr, cReNr, dSumme, lDatVon, lDatBis
            
            cSQL = "Update REKOPF set RESUMME = " & Trim$(Str$(dSumme)) & " "
            cSQL = cSQL & "where SCHLUESSEL = '" & cReNr & "' and REDATUM = " & cReDatum & " "
            gdBase.Execute cSQL, dbFailOnError
            
            'OFPO auch, offene Postenliste
            cSQL = "Update OFPO set RESUMME = " & Trim$(Str$(dSumme)) & " "
            cSQL = cSQL & "where SCHLUESSEL = '" & cReNr & "' and REDATUM = " & cReDatum & " "
            gdBase.Execute cSQL, dbFailOnError
            
            MoveRePos2DruRePosWK24c cReNr
            
            cSQL = "Delete from KREDIT where KUNDNR = " & cKdnr & " "
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lDatVon)) & " and ADATE <= " & Trim$(Str$(lDatBis)) & " "
            gdBase.Execute cSQL, dbFailOnError
            
            
            cSQL = "Update REKOPF set RESUMME = " & Trim$(Str$(dSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
            gdBase.Execute cSQL, dbFailOnError
            
            'OFPO auch, offene Postenliste
            cSQL = "Update OFPO set RESUMME = " & Trim$(Str$(dSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
            gdBase.Execute cSQL, dbFailOnError
            
            aktuali_newOFPO cReNr
            
            cSQL = "Update DRU_REKO set RESUMME = " & Trim$(Str$(dSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
            gdBase.Execute cSQL, dbFailOnError
                        
            cSQL = "Create Index SCHLUESSEL on DRU_REKO (SCHLUESSEL)"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index SCHLUESSEL on DRU_REPO (SCHLUESSEL)"
            gdBase.Execute cSQL, dbFailOnError
            
            Dim cNettoVoll  As String
            Dim cNettoErm   As String
            
            cNettoErm = Format(ermNettoERM, "#####0.00")
            cNettoErm = SwapStr(cNettoErm, ",", ".")
            
            cNettoVoll = Format(ermNettoVoll, "#####0.00")
            cNettoVoll = SwapStr(cNettoVoll, ",", ".")
            
            loeschNEW "NETTOS", gdBase
    
            cSQL = "Create Table NETTOS "
            cSQL = cSQL & "("
            cSQL = cSQL & " NettERM double"
            cSQL = cSQL & ", NettVol double"
            cSQL = cSQL & ") "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert Into NETTOS (NETTERM , NETTVOL) values ( " & cNettoErm & "," & cNettoVoll & " ) "
            gdBase.Execute cSQL, dbFailOnError
            
            Select Case gcRePreisKz
                Case Is = "N"
                    If Modul6.FindFile(gcDBPfad, "aWKL24cl.rpt") Then
                        sCrystaldat = "aWKL24cl"
                    Else
                        sCrystaldat = "aWKL24ca"
                    End If
                Case Is = "B"
                    If Modul6.FindFile(gcDBPfad, "aWKL24cm.rpt") Then
                        sCrystaldat = "aWKL24cm"
                    Else
                        sCrystaldat = "aWKL24cb"
                    End If
                Case Is = "O"
                    If Modul6.FindFile(gcDBPfad, "aWKL24cn.rpt") Then
                        sCrystaldat = "aWKL24cn"
                    Else
                        sCrystaldat = "aWKL24cc"
                    End If
                Case Is = "Z"
                    If Modul6.FindFile(gcDBPfad, "aWKL24co.rpt") Then
                        sCrystaldat = "aWKL24co"
                    Else
                        sCrystaldat = "aWKL24cd"
                    End If
                Case Else
                    If Modul6.FindFile(gcDBPfad, "aWKL24cl.rpt") Then
                        sCrystaldat = "aWKL24cl"
                    Else
                        
                        sCrystaldat = "aWKL24ca"
                    End If
            End Select
            
            If bsofortanDrucker Then
                bsofortanDrucker = False
                reportbildschirmToPrinter sCrystaldat
            Else
                reportbildschirm "dspez", sCrystaldat
            End If

            lCopies = Val(Text4.Text)
            If lCopies = 0 Then
                lCopies = 1
            End If
            
            lrenr = Val(cReNr)
            lrenr = lrenr + 1
            Text3.Text = Format$(lrenr, "##############0")
            Text3.Refresh
            DoEvents
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeRechnungenWK24c"
        Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
        Resume Next
    End If
End Sub
Private Function ermNettoERM() As String
    On Error GoTo LOKAL_ERROR
    
    ermNettoERM = "0"
    
    Dim cSQL As String
    Dim rs As Recordset
    
    cSQL = "Select Sum(GPreis)as maxi from DRU_REPO where MWST = 'E' "
    Set rs = gdBase.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermNettoERM = rs!maxi
            ermNettoERM = (ermNettoERM * 100) / (100 + gdMWStE)
        End If
    End If
    rs.Close: Set rs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNettoERM"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermNettoVoll() As String
    On Error GoTo LOKAL_ERROR
    
    ermNettoVoll = "0"
    
    Dim cSQL As String
    Dim rs As Recordset
    
    cSQL = "Select Sum(GPreis)as maxi from DRU_REPO where MWST = 'V' "
    Set rs = gdBase.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermNettoVoll = rs!maxi
            ermNettoVoll = (ermNettoVoll * 100) / (100 + gdMWStV)
        End If
    End If
   
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNettoVoll"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    Screen.MousePointer = 11
    
    PositionierenWK24c
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = "1"
    
    ctmp = fnHoleMaxReNrWK24c()
    Text3.Text = HoleNaechsteReNr
    
    Label1(8).Caption = ctmp
    bsofortanDrucker = False
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub PositionierenWK24c()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 0
    Frame1.Left = 40
    Frame1.Height = 3975
    Frame1.Width = 3495
    
    Frame2.Top = 0
    Frame2.Left = 3640
    Frame2.Height = 8295
    Frame2.Width = 8175
    
    Frame3.Top = 5640
    Frame3.Left = 40
    Frame3.Height = 2655
    Frame3.Width = 3495
    
    Frame4.Top = 3960
    Frame4.Left = 40
    Frame4.Height = 1575
    Frame4.Width = 3495
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK24c"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MaskEdBox1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = glSelBack1
    MaskEdBox1(Index).SelStart = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MaskEdBox1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    MaskEdBox1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lRet As Long
    Dim lAnz As Long
    Dim ctmp As String
    Dim cTmp2 As String
    Dim lPos As Long
    Dim cZeichen As String
    
    Dim lMon As Long
    Dim lJahr As Long
    Dim lDatum As Long
    Dim lDatBis As Long
    Dim lcount As Long
    Dim cMon As String
    Dim cJahr As String
    Dim cDatum As String
    
    Select Case Index
        Case Is = 0     'Suchen
            lRet = fnPruefeEingabeDialogWK24c()
            Select Case lRet
                Case Is = 0
                    SucheKreditVerkaeufeWK24c
                    
                Case Is = 1
                    MsgBox "Bitte die Untergrenze der Kundennummern angeben!", vbInformation, "Winkiss Hinweis:"
                    MaskEdBox1(0).SetFocus
                Case Is = 2
                    MsgBox "Bitte die Obergrenze der Kundennummern angeben!", vbInformation, "Winkiss Hinweis:"
                    MaskEdBox1(1).SetFocus
                Case Is = 3
                    MsgBox "Bitte ein gültiges VON-Datum eingeben!", vbInformation, "Winkiss Hinweis:"
                    MaskEdBox1(2).SetFocus
                Case Is = 4
                    MsgBox "Bitte ein gültiges BIS-Datum eingeben!", vbInformation, "Winkiss Hinweis:"
                    MaskEdBox1(3).SetFocus
            End Select
            
        Case Is = 1     'Schließen
            Unload frmWK24c
            
        Case Is = 2     'Drucke markierte Rechnungen
            If Trim$(Text3.Text) = "" Then
                MsgBox "Bitte die Start-Rechnungsnummer angeben!", vbInformation, "Winkiss Hinweis:"
                Text3.SetFocus
                Exit Sub
            Else
                ctmp = Text3.Text
                For lPos = 1 To Len(ctmp)
                    cZeichen = Mid(ctmp, lPos, 1)
                    If InStr("1234567890", cZeichen) = 0 Then
                        MsgBox "Für ein automatisches Hochzählen darf die Rechnungsnummer nur Ziffern enthalten!", vbInformation, "Winkiss Hinweis:"
                        Text3.SetFocus
                        Exit Sub
                    End If
                Next lPos
            End If
            lAnz = 0
            For lRet = 0 To List1.ListCount - 1
                If List1.Selected(lRet) = True Then
                    lAnz = lAnz + 1
                End If
            Next lRet
            If lAnz = 0 Then
                MsgBox "Bitte mindestens eine Rechnung auswählen!", vbInformation, "Winkiss Hinweis:"
                List1.SetFocus
                Exit Sub
            End If
            
            '*************************************************************
            '* Anzeige der gesetzten Parameter
            '*************************************************************
            ctmp = "Bitte überprüfen Sie die nachfolgenden Angaben:" & vbCrLf & vbCrLf
            
            ctmp = ctmp & "Arbeitsauftrag:     Drucke nur markierte Rechnungen" & vbCrLf & vbCrLf
            
            cTmp2 = Trim$(Str$(Val(MaskEdBox1(0).Text)))
            cTmp2 = Space$(6 - Len(cTmp2)) & cTmp2
            ctmp = ctmp & "Kundennummern von:  " & cTmp2 & " bis "
            cTmp2 = Trim$(Str$(Val(MaskEdBox1(1).Text)))
            cTmp2 = Space$(6 - Len(cTmp2)) & cTmp2
            ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
            ctmp = ctmp & "alle Einkäufe von:  " & MaskEdBox1(2).Text & " bis " & MaskEdBox1(3).Text & vbCrLf & vbCrLf
            If Option1(0).Value = True Then
                cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise enthalten MWSt-Beträge (BRUTTO)" & vbCrLf
                cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                cTmp2 = cTmp2 & "am Ende der Rechnung als 'inkl. MWSt' ausgewiesen"
            ElseIf Option1(1).Value = True Then
                cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl. MWSt. (NETTO)" & vbCrLf
                cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                cTmp2 = cTmp2 & "aus den einzelnen Positionen herausgerechnet und" & vbCrLf
                cTmp2 = cTmp2 & "am Ende der Rechnung wieder aufaddiert ('zzgl. MWSt')"
                
            ElseIf Option1(2).Value = True Then
                cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise OHNE MWSt. " & vbCrLf
                cTmp2 = cTmp2 & "d.h. es wird kein Wert für 'inkl. MWSt' " & vbCrLf
                cTmp2 = cTmp2 & "oder 'zzgl. MWSt' ermittelt"
            Else
                cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl MWSt." & vbCrLf
                cTmp2 = cTmp2 & "d.h. die Artikelpreise werden als Netto-Preise gesehen" & vbCrLf
                cTmp2 = cTmp2 & "und für jeden Artikel wird die MWSt. berechnet, " & vbCrLf
                cTmp2 = cTmp2 & "die am Ende der Rechnung aufaddiert wird ('zzgl. MWSt')"
                
            End If
            ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
            
            cTmp2 = "Starte mit ReNr.:   " & Text3.Text
            
            ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                    
            cTmp2 = "Porto + Verpackung: "
            If Trim$(Text1.Text) = "" Then
                cTmp2 = cTmp2 & "-"
            Else
                cTmp2 = cTmp2 & Text1.Text
            End If
            ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
            
            cTmp2 = "freier Text:        "
            If Trim$(Text2.Text) = "" Then
                cTmp2 = cTmp2 & "Nein"
            Else
                cTmp2 = cTmp2 & "Ja"
            End If
            
            ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
            
            cTmp2 = "Sind diese Angaben richtig?"
            
            ctmp = ctmp & cTmp2
            lRet = MsgBox(ctmp, vbQuestion + vbYesNo, "ÜBERPRÜFUNG")
            
            If lRet = vbYes Then
                lRet = MsgBox("Druckdaten ohne Druckvorschau ausdrucken?", vbQuestion + vbYesNo, "Winkiss Frage:")
                
                If lRet = vbYes Then
                    bsofortanDrucker = True
                Else
                    bsofortanDrucker = False
                End If
                
                DruckeRechnungenWK24c lAnz
                Command2_Click 0
            End If
            
        Case Is = 3     'Drucke alle Rechnungen
            If Trim$(Text3.Text) = "" Then
                MsgBox "Bitte die Start-Rechnungsnummer angeben!", vbCritical, "STOP!"
                Text3.SetFocus
                Exit Sub
            Else
                ctmp = Text3.Text
                For lPos = 1 To Len(ctmp)
                    cZeichen = Mid(ctmp, lPos, 1)
                    If InStr("1234567890", cZeichen) = 0 Then
                        MsgBox "Für ein automatisches Hochzählen darf die Rechnungsnummer nur Ziffern enthalten!", vbCritical, "STOP!"
                        Text3.SetFocus
                        Exit Sub
                    End If
                Next lPos
            End If
            lAnz = 0
            For lRet = 0 To List1.ListCount - 1
                List1.Selected(lRet) = True
                lAnz = lAnz + 1
            Next lRet
            If lAnz > 0 Then
                '*************************************************************
                '* Anzeige der gesetzten Parameter
                '*************************************************************
                ctmp = "Bitte überprüfen Sie die nachfolgenden Angaben:" & vbCrLf & vbCrLf
                
                ctmp = ctmp & "Arbeitsauftrag:     Drucke alle angezeigten Rechnungen" & vbCrLf & vbCrLf
                
                cTmp2 = Trim$(Str$(Val(MaskEdBox1(0).Text)))
                cTmp2 = Space$(6 - Len(cTmp2)) & cTmp2
                ctmp = ctmp & "Kundennummern von:  " & cTmp2 & " bis "
                cTmp2 = Trim$(Str$(Val(MaskEdBox1(1).Text)))
                cTmp2 = Space$(6 - Len(cTmp2)) & cTmp2
                ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                ctmp = ctmp & "alle Einkäufe von:  " & MaskEdBox1(2).Text & " bis " & MaskEdBox1(3).Text & vbCrLf & vbCrLf
                If Option1(0).Value = True Then
                    cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise enthalten MWSt-Beträge (BRUTTO)" & vbCrLf
                    cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                    cTmp2 = cTmp2 & "am Ende der Rechnung als 'inkl. MWSt' ausgewiesen"
                ElseIf Option1(1).Value = True Then
                    cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl. MWSt. (NETTO)" & vbCrLf
                    cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                    cTmp2 = cTmp2 & "aus den einzelnen Positionen herausgerechnet und" & vbCrLf
                    cTmp2 = cTmp2 & "am Ende der Rechnung wieder aufaddiert ('zzgl. MWSt')"
                    
                ElseIf Option1(2).Value = True Then
                    cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise OHNE MWSt. " & vbCrLf
                    cTmp2 = cTmp2 & "d.h. es wird kein Wert für 'inkl. MWSt' " & vbCrLf
                    cTmp2 = cTmp2 & "oder 'zzgl. MWSt' ermittelt"
                Else
                    cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl MWSt." & vbCrLf
                    cTmp2 = cTmp2 & "d.h. die Artikelpreise werden als Netto-Preise gesehen" & vbCrLf
                    cTmp2 = cTmp2 & "und für jeden Artikel wird die MWSt. berechnet, " & vbCrLf
                    cTmp2 = cTmp2 & "die am Ende der Rechnung aufaddiert wird ('zzgl. MWSt')"
                    
                End If
                ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                
                cTmp2 = "Starte mit ReNr.:   " & Text3.Text
                
                ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                        
                cTmp2 = "Porto + Verpackung: "
                If Trim$(Text1.Text) = "" Then
                    cTmp2 = cTmp2 & "-"
                Else
                    cTmp2 = cTmp2 & Text1.Text
                End If
                ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                
                cTmp2 = "freier Text:        "
                If Trim$(Text2.Text) = "" Then
                    cTmp2 = cTmp2 & "Nein"
                Else
                    cTmp2 = cTmp2 & "Ja"
                End If
                
                ctmp = ctmp & cTmp2 & vbCrLf & vbCrLf
                
                cTmp2 = "Sind diese Angaben richtig?"
                
                ctmp = ctmp & cTmp2
                lRet = MsgBox(ctmp, vbQuestion + vbYesNo, "ÜBERPRÜFUNG")
                If lRet = vbYes Then
                    lRet = MsgBox("Druckdaten ohne Druckvorschau ausdrucken?", vbQuestion + vbYesNo, "Winkiss Frage:")
                
                    If lRet = vbYes Then
                        bsofortanDrucker = True
                    Else
                        bsofortanDrucker = False
                    End If
                    DruckeRechnungenWK24c lAnz
                    Command2_Click 0
                End If
            End If
        
        Case Is = 4     'Alle Kunden
            MaskEdBox1(0).Text = "1_____"
            MaskEdBox1(1).Text = "999999"
            MaskEdBox1(2).SetFocus
        
        Case Is = 5     'dieser Monat
            lMon = Month(Now)
            lJahr = Year(Now)
            cMon = Trim$(Str$(lMon))
            cJahr = Trim$(Str$(lJahr))
            cMon = String$(2 - Len(cMon), "0") & cMon
            cDatum = "01." & cMon & "." & cJahr
            lDatum = DateValue(cDatum)
            
            MaskEdBox1(2).Text = cDatum
                        
            For lcount = 27 To 30
                lDatBis = lDatum + lcount
                cDatum = Format$(lDatBis, "DD.MM.YYYY")
                If Not IsDate(cDatum) Then
                    lDatBis = lDatBis - 1
                    cDatum = Format$(lDatBis, "DD.MM.YYYY")
                    Exit For
                End If
            Next lcount
            MaskEdBox1(3).Text = cDatum
            
        Case Is = 6     'letzter Monat
            lMon = Month(Now)
            lJahr = Year(Now)
            lMon = lMon - 1
            If lMon = 0 Then
                lMon = 12
                lJahr = lJahr - 1
            End If
            cMon = Trim$(Str$(lMon))
            cJahr = Trim$(Str$(lJahr))
            cMon = String$(2 - Len(cMon), "0") & cMon
            cDatum = "01." & cMon & "." & cJahr
            lDatum = DateValue(cDatum)
            
            MaskEdBox1(2).Text = cDatum
                        
            For lcount = 27 To 30
                lDatBis = lDatum + lcount
                cDatum = Format$(lDatBis, "DD.MM.YYYY")
                If Not IsDate(cDatum) Or Month(lDatBis) <> lMon Then
                    lDatBis = lDatBis - 1
                    cDatum = Format$(lDatBis, "DD.MM.YYYY")
                    Exit For
                End If
            Next lcount
            MaskEdBox1(3).Text = cDatum
            
        Case Is = 7     'nächste ReNr
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MoveRePos2DruRePosWK24c(cReNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "DRU_REPO", gdBase
    cSQL = "Create Table DRU_REPO "
    cSQL = cSQL & "(SCHLUESSEL Text(15)"
    cSQL = cSQL & ", KAUFDATUM Datetime"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", ANZAHL long"
    cSQL = cSQL & ", EPREIS double"
    cSQL = cSQL & ", GPREIS double"
    cSQL = cSQL & ", MWST Text(1)"
    cSQL = cSQL & ", PREISKZ Integer"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", Reihenf long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert Into DRU_REPO "
    cSQL = cSQL & "Select * from REPOS "
    cSQL = cSQL & " where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveRePos2DruRePosWK24c"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MoveDaten2RechnungWK24c(cKdnr As String, cReNr As String, dSumme As Double, lDatVon As Long, lDatBis As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim dWert As Double
    Dim lDatum As Long
    Dim cDatum As String
    Dim lartnr As Long
    Dim cBezeich As String
    Dim lAnzahl As Long
    Dim dEPreis As Double
    Dim dGPreis As Double
    Dim cMWST As String
    Dim lPreisKz As Long
        
    Dim cInto As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from KREDIT where KUNDNR = " & cKdnr & " and ADATE >= " & Trim$(Str$(lDatVon)) & " and ADATE <= " & Trim$(Str$(lDatBis))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzahl = rsrs.RecordCount
        rsrs.MoveFirst
        lAnzahl = 0
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Adate) Then
                lDatum = rsrs!Adate
                cDatum = Format$(lDatum, "DD.MM.YYYY")
            Else
                cDatum = Space$(10)
            End If
            
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = 0
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            
            If Not IsNull(rsrs!Menge) Then
                lAnzahl = rsrs!Menge
            Else
                lAnzahl = 0
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                dEPreis = rsrs!vkpr
            Else
                dEPreis = 0
            End If
            
            If Not IsNull(rsrs!GVKPR) Then
                dGPreis = rsrs!GVKPR
            Else
                dGPreis = 0
            End If
            dSumme = dSumme + dGPreis
            
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            Else
                cMWST = "V"
            End If
            
            If Not IsNull(rsrs!PREISKZ) Then
                lPreisKz = rsrs!PREISKZ
            Else
                lPreisKz = 0
            End If
            
            cInto = "Insert Into REPOS "
            cSQL = "("
            cSQL = cSQL & " SCHLUESSEL"
            cSQL = cSQL & ", KAUFDATUM"
            cSQL = cSQL & ", ARTNR"
            cSQL = cSQL & ", BEZEICH"
            cSQL = cSQL & ", ANZAHL"
            cSQL = cSQL & ", EPREIS"
            cSQL = cSQL & ", GPREIS"
            cSQL = cSQL & ", MWST"
            cSQL = cSQL & ", PREISKZ"
            cSQL = cSQL & ") values ("
            cSQL = cSQL & "'" & cReNr & "'"
            cSQL = cSQL & ", " & Trim$(Str$(lDatum)) & ""
            cSQL = cSQL & ", " & Trim$(Str$(lartnr)) & ""
            cSQL = cSQL & ", '" & cBezeich & "' "
            cSQL = cSQL & ", " & Trim$(Str$(lAnzahl)) & ""
            cSQL = cSQL & ", " & Trim$(Str$(dEPreis)) & ""
            cSQL = cSQL & ", " & Trim$(Str$(dGPreis)) & ""
            cSQL = cSQL & ", '" & cMWST & "' "
            cSQL = cSQL & ", " & Trim$(Str$(lPreisKz)) & ""
            cSQL = cSQL & ") "
        
            gdBase.Execute cInto & cSQL, dbFailOnError
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveDaten2RechnungWK24c"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    
    cValid = "1234567890," & Chr$(8)
    
    
    
    cZeichen = Chr$(KeyAscii)
    If cZeichen = "." Then
        cZeichen = ","
    End If
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)



    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        If InStr(Text1.Text, ",") > 0 And cZeichen = "," Then
            KeyAscii = 0
        End If
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR

    Text1.BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    If cZeichen <= "0" Or cZeichen >= "9" Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sammelrechnung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


