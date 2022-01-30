VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL18 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sortimentsanalyse"
   ClientHeight    =   8625
   ClientLeft      =   1605
   ClientTop       =   1440
   ClientWidth     =   11910
   Icon            =   "frmWKL18.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Tag             =   "Sortimen"
   Begin Threed.SSCommand cmdHelp 
      Height          =   600
      Left            =   11040
      TabIndex        =   94
      ToolTipText     =   "Winkiss Hilfe"
      Top             =   120
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   78
      BevelWidth      =   3
      Font3D          =   1
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "frmWKL18.frx":0442
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8520
      TabIndex        =   67
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   7
         Left            =   9240
         TabIndex        =   79
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   6000
         TabIndex        =   77
         Top             =   2520
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         MaxLength       =   3
         TabIndex        =   70
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         MaxLength       =   25
         TabIndex        =   69
         Top             =   3240
         Width           =   3735
      End
      Begin sevCommand3.Command Command2 
         Height          =   480
         Index           =   0
         Left            =   3000
         TabIndex        =   68
         Top             =   5040
         Width           =   2520
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
         Caption         =   "Daten speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label38 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   4
         Left            =   120
         TabIndex        =   90
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5880
         X2              =   5880
         Y1              =   5520
         Y2              =   2040
      End
      Begin VB.Label Label34 
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
         ForeColor       =   &H00404000&
         Height          =   855
         Left            =   360
         TabIndex        =   81
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenstellung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3240
         TabIndex        =   80
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C000&
         Caption         =   "Auflistung der Warengruppen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   6000
         TabIndex        =   78
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3240
         TabIndex        =   76
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Sortimentsliste bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Warengruppen-Nr:"
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
         Index           =   3
         Left            =   360
         TabIndex        =   72
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Warengruppenbezeichnung:"
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
         Left            =   360
         TabIndex        =   71
         Top             =   2880
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   2640
      TabIndex        =   58
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   6
         Left            =   9240
         TabIndex        =   63
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   7
         Left            =   4440
         TabIndex        =   60
         Top             =   5640
         Width           =   3135
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
         Caption         =   "Warengruppe löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   3360
         TabIndex        =   59
         Top             =   2520
         Width           =   5055
      End
      Begin VB.Label Label38 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   2
         Left            =   120
         TabIndex        =   88
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Wählen Sie eine Warengruppe!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3360
         TabIndex        =   66
         Top             =   4920
         Width           =   5055
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C000&
         Caption         =   "Auflistung der Warengruppen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   3360
         TabIndex        =   65
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenstellung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   2760
         TabIndex        =   64
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Vorgang: Warengruppe löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   2760
         TabIndex        =   62
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   2760
         TabIndex        =   61
         Top             =   1560
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   46
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   5
         Left            =   9240
         TabIndex        =   55
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   7
         Left            =   6960
         TabIndex        =   49
         Top             =   4920
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   6
         Left            =   6960
         TabIndex        =   48
         Top             =   4320
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
         Caption         =   "Bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   5
         Left            =   6960
         TabIndex        =   47
         Top             =   3720
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
         Caption         =   "Neu"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label38 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   7
         Left            =   120
         TabIndex        =   93
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Rechts
         Caption         =   "Name der Zusammenstellung: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   85
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Label Label24 
         Caption         =   "Zusammenstellung "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   57
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Zentriert
         Caption         =   "Zusammenstellung "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   10935
      End
      Begin VB.Label Label22 
         Caption         =   "Hier können Sie  Warengruppen, die Zusammenfassung von AGN´s, erstellen, löschen oder bearbeiten."
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
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   10935
      End
      Begin VB.Label Label21 
         Caption         =   "Warengruppen löschen"
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
         Left            =   2160
         TabIndex        =   53
         Top             =   4920
         Width           =   4575
      End
      Begin VB.Label Label20 
         Caption         =   "Warengruppen bearbeiten"
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
         Left            =   2160
         TabIndex        =   52
         Top             =   4320
         Width           =   4575
      End
      Begin VB.Label Label14 
         Caption         =   "Warengruppen erstellen"
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
         Left            =   2160
         TabIndex        =   51
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Wie möchten Sie vorgehen?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   50
         Top             =   2880
         Width           =   5415
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   480
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   10455
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
         Height          =   355
         Index           =   7
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   120
         ToolTipText     =   "Sind Ihre Lieferantenkürzel ordentlich gepflegt, so können Sie mit dem Kürzel arbeiten. Geben Sie z.B. Joop ein!"
         Top             =   4920
         Width           =   1215
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   16
         Left            =   8760
         TabIndex        =   113
         Top             =   4320
         Width           =   2535
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
         Caption         =   "Lieferanten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   15
         Left            =   5160
         TabIndex        =   112
         Top             =   3720
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "Artikel anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   14
         Left            =   6120
         TabIndex        =   111
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
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
         Caption         =   "TopCos"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   13
         Left            =   6120
         TabIndex        =   110
         Top             =   4920
         Visible         =   0   'False
         Width           =   2535
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
         Caption         =   "COSPAR"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   12
         Left            =   6120
         TabIndex        =   109
         Top             =   5520
         Visible         =   0   'False
         Width           =   2535
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
         Caption         =   "WIR FÜR SIE"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'Kein
         Height          =   1695
         Left            =   600
         TabIndex        =   102
         Top             =   3360
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "Vorjahr Zeitraum"
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
            TabIndex        =   108
            Top             =   1320
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vormonat"
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
            Left            =   120
            TabIndex        =   107
            Top             =   1080
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "aktueller Monat"
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
            TabIndex        =   106
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "aktuelles Jahr"
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
            TabIndex        =   104
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vorjahr"
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
            Left            =   120
            TabIndex        =   103
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Datum Voreinstellung"
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
            Left            =   120
            TabIndex        =   105
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   3240
         TabIndex        =   96
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   8
         Left            =   8760
         TabIndex        =   95
         Top             =   4920
         Width           =   2535
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
         Caption         =   "AGN Analyse"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   3
         Left            =   8760
         TabIndex        =   40
         Top             =   5520
         Width           =   2535
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
         Caption         =   "Auswertung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   2
         Left            =   8760
         TabIndex        =   37
         Top             =   6120
         Width           =   2535
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
         Height          =   360
         Index           =   3
         Left            =   1080
         TabIndex        =   34
         Top             =   2880
         Width           =   1575
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
         Height          =   360
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Top             =   2400
         Width           =   1575
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   20
         Left            =   2760
         TabIndex        =   114
         ToolTipText     =   "Kalender"
         Top             =   2400
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   21
         Left            =   2760
         TabIndex        =   115
         ToolTipText     =   "Kalender"
         Top             =   2880
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   17
         Left            =   8760
         TabIndex        =   116
         Top             =   3720
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "Bereinigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   6
         Left            =   4920
         TabIndex        =   119
         Top             =   4920
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "kein Lieferant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3600
         TabIndex        =   122
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant:"
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
         Left            =   3600
         TabIndex        =   121
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz beheben"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   118
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWKL18.frx":09AC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1575
         Index           =   1
         Left            =   7800
         TabIndex        =   117
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Zeitraum:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1215
         Index           =   0
         Left            =   4440
         TabIndex        =   97
         Top             =   2400
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label38 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   3
         Left            =   120
         TabIndex        =   89
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Vorgang: Auswertung"
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
         Left            =   1320
         TabIndex        =   41
         Top             =   6120
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bis:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         Caption         =   "Vorgang: Auswertung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2760
         TabIndex        =   36
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Zeitraum:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   35
         Top             =   2040
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10680
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   1
         Left            =   9240
         TabIndex        =   29
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   0
         Left            =   6000
         TabIndex        =   28
         Top             =   3960
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
      Begin VB.ListBox List12 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   3840
         TabIndex        =   27
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Label38 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   6
         Left            =   120
         TabIndex        =   92
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   735
         Left            =   2040
         TabIndex        =   31
         Top             =   4560
         Width           =   7455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "vorhandene Zusammenstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   26
         Top             =   1920
         Width           =   3975
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   525
      Index           =   4
      Left            =   9480
      TabIndex        =   24
      Top             =   7800
      Width           =   2175
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
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   10200
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   11
         Left            =   4200
         TabIndex        =   101
         Top             =   5520
         Width           =   2655
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
         Caption         =   "Cospar - Zusammenstellung erstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   10
         Left            =   1440
         TabIndex        =   100
         Top             =   5520
         Width           =   2655
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
         Caption         =   "TopCos - Zusammenstellung erstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   615
         Index           =   9
         Left            =   1440
         TabIndex        =   99
         Top             =   4800
         Width           =   2655
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
         Caption         =   "WIR FÜR SIE - Zusammenstellung erstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   4
         Left            =   9240
         TabIndex        =   42
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtProfilname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         MaxLength       =   35
         TabIndex        =   12
         Top             =   2400
         Width           =   3735
      End
      Begin sevCommand3.Command cmdNeu 
         Height          =   495
         Left            =   4800
         TabIndex        =   11
         Top             =   3000
         Width           =   1935
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
      Begin VB.Label Label38 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   87
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   44
         Top             =   3960
         Width           =   10095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   11175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "neue Zusammenstellung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   13
         Top             =   2040
         Width           =   3735
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   5
         Left            =   5520
         TabIndex        =   2
         Top             =   2280
         Width           =   615
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
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   4
         Left            =   5520
         TabIndex        =   3
         Top             =   4800
         Width           =   615
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
      Begin VB.ListBox List8 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   6360
         TabIndex        =   1
         Top             =   2280
         Width           =   4935
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   0
         Top             =   2280
         Width           =   5175
      End
      Begin sevCommand3.Command Command2 
         Height          =   480
         Index           =   3
         Left            =   9240
         TabIndex        =   4
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label38 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   5
         Left            =   120
         TabIndex        =   91
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   1320
         TabIndex        =   84
         Top             =   6120
         Width           =   7815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "akuelle Zusammenstellung:"
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
         Index           =   5
         Left            =   2760
         TabIndex        =   83
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C000&
         Caption         =   "Neu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   5400
         TabIndex        =   82
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "akuelle Warengruppe:"
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
         Index           =   9
         Left            =   2760
         TabIndex        =   74
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C000&
         Caption         =   "Neu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   5400
         TabIndex        =   73
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   5400
         TabIndex        =   45
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Sortimentsliste bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label3 
         Appearance      =   0  '2D
         BackColor       =   &H00C0C000&
         Caption         =   " Artikelgruppen der Warengruppe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   7
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelgruppen (Multi-Auswahl):"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   5175
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   20
      Top             =   3360
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
      Caption         =   "Neu"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   21
      Top             =   3960
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
      Caption         =   "Bearbeiten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   7800
      TabIndex        =   22
      Top             =   4560
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   7800
      TabIndex        =   23
      Top             =   5160
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
      Caption         =   "Auswerten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   8
      Left            =   7800
      TabIndex        =   98
      Top             =   5760
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
      Caption         =   "Testen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label38 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   86
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Eine bestehende Zusammenstellung auswerten."
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
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   7335
   End
   Begin VB.Label Label15 
      Caption         =   "Wie möchten Sie vorgehen?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   7335
   End
   Begin VB.Label Label16 
      Caption         =   "Eine neue Zusammenstellung erstellen."
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
      Left            =   240
      TabIndex        =   17
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Label Label17 
      Caption         =   "Eine bestehende Zusammenstellung bearbeiten."
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
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   7335
   End
   Begin VB.Label Label18 
      Caption         =   "Eine bestehende Zusammenstellung löschen."
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
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   7335
   End
   Begin VB.Label Label19 
      Caption         =   "Hier können Sie mehrere Artikelgruppen zu Warengruppen zusammenfassen und über verschiedene Zeiträume auswerten."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   11415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Sortiments - Analyse"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmWKL18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gcPname As String
Dim iPruef As Integer

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 6
            Text1_KeyUp 7, vbKeyF2, 0
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "agnfrei", gdBase
    loeschNEW "Agnana", gdBase
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
Private Sub WKL18Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 240
    Frame1.Height = 6735
    Frame1.Width = 11415
    
    Frame2.Top = 960
    Frame2.Left = 240
    Frame2.Height = 6735
    Frame2.Width = 11415
    
    Frame4.Top = 960
    Frame4.Left = 240
    Frame4.Height = 6735
    Frame4.Width = 11415
    
    Frame5.Top = 960
    Frame5.Left = 240
    Frame5.Height = 6735
    Frame5.Width = 11415
    
    Frame10.Top = 960
    Frame10.Left = 240
    Frame10.Height = 6735
    Frame10.Width = 11415
    
    Frame6.Top = 960
    Frame6.Left = 240
    Frame6.Height = 6735
    Frame6.Width = 11415
    
    
    
    Frame3.Top = 960
    Frame3.Left = 240
    Frame3.Height = 6735
    Frame3.Width = 11415

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL18Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DeleteAGNFromWGN(cWGN As String, cProfile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cLBSatz As String
    Dim cSQL As String
    
    If List8.ListCount = 0 Then
        Exit Sub
    End If
    
    
    If List8.ListIndex < 0 Then
        List4.AddItem List8.list(List8.TopIndex)
        cLBSatz = List8.list(List8.TopIndex)
        List8.RemoveItem List8.TopIndex
        
    Else
        cLBSatz = List8.list(List8.ListIndex)
        List8.RemoveItem List8.ListIndex
    End If
    
    cSQL = "Delete from WGNAGN where WGN = " & cWGN & " and AGN = " & Left$(cLBSatz, 5) & " and Pname = '" & cProfile & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DeleteAGNFromWGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Function fnPruefeDialogWGNDEFSta17%()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    Dim cLBSatz As String
    
    fnPruefeDialogWGNDEFSta17% = 0
    
    ctmp = Text4(0).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeDialogWGNDEFSta17% = 11
        Exit Function
    End If
    
    ctmp = Text4(1).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeDialogWGNDEFSta17% = 21
        Exit Function
    End If
    
    If Len(Text4(0).Text) = 1 Then
        Text4(0).Text = "  " & Text4(0).Text
    ElseIf Len(Text4(0).Text) = 2 Then
        Text4(0).Text = " " & Text4(0).Text
    End If
    
    For lcount = 0 To List2.ListCount - 1
        cLBSatz = List2.list(lcount)
        If Left$(cLBSatz, 3) = Text4(0).Text Then
            fnPruefeDialogWGNDEFSta17% = 12
            Exit Function
        End If
    Next lcount
    
    For lcount = 0 To List2.ListCount - 1
        cLBSatz = List2.list(lcount)
        If UCase(Trim(Right$(cLBSatz, Len(cLBSatz) - 3))) = UCase(Trim(Text4(1).Text)) Then
            fnPruefeDialogWGNDEFSta17% = 16
            Exit Function
        End If
    Next lcount

    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogWGNDEFSta17%"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub FuelleWGNLIST(cprofil As String, List6 As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cLBSatz As String
    
    List6.Clear
    
    cSQL = "Select * from WGNDEF where Pname = '" & cprofil & "' "
    cSQL = cSQL & " order by WGN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!WGN) Then
                cFeld = Trim$(Str$(rsrs!WGN))
            Else
                cFeld = "0"
            End If
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!WGNTEXT) Then
                cFeld = rsrs!WGNTEXT
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld
            
            List6.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleWGNLIST"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TestSchluessel()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim i As Integer
    
    cSQL = "Delete from WGNDEF where pname = 'Kiss Test'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from WGNAGN where pname = 'Kiss Test'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (111,'Pflege', 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (201,'Dekorative', 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (301,'Damendüfte', 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (401,'Herrendüfte', 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (999,'restliche AGN', 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 111 To 118
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (111," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 121 To 126
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (111," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (111,185, 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (111,313, 'Kiss Test')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 131 To 139
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (201," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 141 To 148
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (301," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 181 To 188
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (301," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 191 To 196
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (301," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 161 To 169
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (401," & i & ", 'Kiss Test')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    loeschNEW "agnfrei", gdBase
    
    cSQL = "select agn into agnfrei from agndbf where agn not in (select agn from WGNAGN where pname = 'Kiss Test' ) "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Insert into WGNAGN select "
    cSQL = cSQL & " 999 as wgn , agn , 'Kiss Test' as pname from agnfrei "
    gdBase.Execute cSQL, dbFailOnError
    
    HoleAuswertung18 "Kiss Test"
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TestSchluessel"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertAGN2WGN(cWGN As String, cProfile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cLBSatz As String
    Dim cSQL As String
    Dim bFound As Boolean
    Dim iFocuslast As Integer
    
    
    If List4.ListCount < 0 Then
        Label36.ForeColor = vbRed
        Label36.Caption = "Bitte eine oder mehrere Artikelgruppen auswählen!"
        Label36.Refresh
        List4.SetFocus
        Exit Sub
    End If
    
    bFound = False
    For lcount = 0 To List4.ListCount - 1
        If List4.Selected(lcount) Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If bFound = False Then
        Label36.ForeColor = vbRed
        Label36.Caption = "Bitte eine oder mehrere Artikelgruppen auswählen!"
        Label36.Refresh
        List4.SetFocus
        Exit Sub
    End If
    
    
    For lcount = 0 To List4.ListCount - 1
        If List4.Selected(lcount) Then

            cLBSatz = List4.list(lcount)

            cSQL = "Insert into WGNAGN (WGN, AGN,Pname) values (" & cWGN & ", " & Left$(cLBSatz, 5) & " , '" & cProfile & "') "
            gdBase.Execute cSQL, dbFailOnError
            List8.AddItem cLBSatz
        End If
        
    Next lcount
    
    bFound = True
    Do While bFound
        For lcount = 0 To List4.ListCount - 1
            If List4.Selected(lcount) Then
                List4.RemoveItem (lcount)
                iFocuslast = lcount
                Exit For
            End If
        Next lcount
        
        bFound = False
        For lcount = 0 To List4.ListCount - 1
            If List4.Selected(lcount) Then
                bFound = True
                Exit For
            End If
        Next lcount
    Loop
    
    If iFocuslast <= List4.ListCount - 1 Then
        List4.Selected(iFocuslast) = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertAGN2WGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub SchreibeWGN()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    iRet = fnPruefeDialogWGNDEFSta17%()
    
    Select Case iRet
        Case Is = 0     'ok
            NeuWgnSpeichern Trim(Label31.Caption)

            Text4(0).Text = ""
            Text4(1).Text = ""
            Text4(0).SetFocus
        Case Is = 11    'fehlende WGN
            Label34.ForeColor = vbRed
            Label34.Caption = "Bitte gegen Sie die Warengruppen-Nr an!"
            Label34.Refresh
            Text4(0).SetFocus
            Exit Sub
        Case Is = 12    'neue WGN bereits vorhanden
            Label34.ForeColor = vbRed
            Label34.Caption = "Diese Warengruppen-Nr ist bereits vorhanden!"
            Label34.Refresh
            Text4(0).SetFocus
            Exit Sub
        Case Is = 13    'Änderung auf bestehende WGN
            Label34.ForeColor = vbRed
            Label34.Caption = "Diese Warengruppen-Nr ist bereits vorhanden!"
            Label34.Refresh
            Text4(0).SetFocus
            Exit Sub
        Case Is = 16    'WGN- Text schon vorhanden
            Label34.ForeColor = vbRed
            Label34.Caption = "Diese Warengruppenbezeichnung ist bereits vorhanden!"
            Label34.Refresh
            Text4(1).SetFocus
            Exit Sub
        Case Is = 21    'fehlender Text
            Label34.ForeColor = vbRed
            Label34.Caption = "Bitte gegen Sie die Warengruppen-Bez an!"
            Label34.Refresh
            Text4(1).SetFocus
            Exit Sub
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeWGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SelektWarengruppe(cWGN As String, cprofil As String)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cLBSatz As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lcount As Long
    
    cSQL = "Select WGNAGN.*,agndbf.agtext from WGNAGN,agndbf "
    cSQL = cSQL & " where WGN = " & cWGN & " and Pname = '" & cprofil & "'"
    cSQL = cSQL & " and wgnagn.agn = agndbf.agn "
    List8.Clear
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AGN) Then
                ctmp = Trim$(Str$(rsrs!AGN))
            Else
                ctmp = "0"
            End If
            
            ctmp = Trim$(ctmp)
            ctmp = Space$(5 - Len(ctmp)) & ctmp
            cLBSatz = ctmp & " "
            If Not IsNull(rsrs!AGTEXT) Then
                ctmp = rsrs!AGTEXT
            Else
                ctmp = ""
            End If
            
            cLBSatz = cLBSatz & ctmp
            List8.AddItem cLBSatz
                
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SelektWarengruppe"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "cmdHelp_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cmdNeu_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sProfilname As String
    
    Screen.MousePointer = 11
    If txtProfilname.Text <> "" Then
        Label7.Caption = ""
        Label7.Refresh
        
        gcPname = txtProfilname.Text
        
        Label24.Caption = gcPname
        Label24.Refresh
        
        Label23.Caption = "Zusammenstellung Neu:"
        Label23.Refresh
        Frame1.Visible = True
        Frame3.Visible = False
    Else
    
        Label7.ForeColor = vbRed
        Label7.Caption = "Bitte einen Namen vergeben!"
        Label7.Refresh
        Screen.MousePointer = 0
        txtProfilname.SetFocus
        Exit Sub
    
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdNeu_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Profilfüllen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "Select distinct Pname from wgndef order by pname "
    Set rs = gdBase.OpenRecordset(sSQL)
    List12.Clear
    
    Do While Not rs.EOF
        List12.AddItem rs!pname
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Profilfüllen"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleAuswertung18(pname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iJahr               As Integer
    Dim cVon                As String
    Dim cBis                As String
    Dim cVonvj              As String
    Dim cBisvj              As String
    Dim lVon                As Long
    Dim lBis                As Long
    Dim cSQL                As String
    Dim cdabapfad           As String
    Dim rsrs                As Recordset
    Dim i                   As Integer
    Dim sFilialauswahl      As String
    Dim lKundenanzAkt       As Long
    Dim lKundenanzVJ        As Long
    Dim dLUGavg             As Double
    Dim lLinr               As Long
    Dim cLiefbez            As String
    
    If pname = "Kiss Test" Then
        Option1(4).Value = True
'        Text1(2).Text = "01.01.2008"
'        Text1(3).Text = "31.01.2008"
    End If
    
    If Len(Text1(2).Text) <> 10 Or Len(Text1(3).Text) <> 10 Then
        anzeige "rot", "Bitte geben Sie das Datum im richtigen Format ein z.B.: 03.11.20xx", Label2
        Exit Sub
    End If
    
    iJahr = Right$(Text1(2).Text, 4)
    iJahr = iJahr - 1
    cVon = Text1(2).Text
    cVonvj = Left$(Text1(2).Text, 6) & CStr(iJahr)
    
    iJahr = Right$(Text1(3).Text, 4)
    iJahr = iJahr - 1
    cBis = Text1(3).Text
    cBisvj = Left$(Text1(3).Text, 6) & CStr(iJahr)

    'Zeitraum aktuelles Jahr
    lVon = datumwandlung(cVon)
    lBis = datumwandlung(cBis)
    If lVon > lBis Then
        anzeige "rot", "Der Zeitraumbeginn ist größer als das Zeitraumende.", Label2
        
        Exit Sub
    End If
    
    lLinr = Val(Text1(7).Text)
    cLiefbez = ermLiefBez(lLinr)
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", Label2
    
    
    Datenbankwechsel
    Me.Refresh
    
    loeschNEW "FILAW", dabalokal
    CreateTableT2 "FILAW", dabalokal
    
    sFilialauswahl = gRegister.firma
    
    cSQL = "Insert into FILAW (FILWAHL) values ('" & sFilialauswahl & "') "
    dabalokal.Execute cSQL, dbFailOnError
    
    cdabapfad = gcPfad
    If Right$(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "lokal.mdb"
    
    
    If glLokalAktuZeit = 0 Then
               
    Else
        'Import WGNDEF und WGNAGN
        loeschNEW "WGNDEF", dabalokal
        cSQL = "Select WGNDEF.* INTO WGNDEF IN '" & cdabapfad & "' from WGNDEF "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "WGNAGN", dabalokal
        cSQL = "Select WGNAGN.* INTO WGNAGN IN '" & cdabapfad & "' from WGNAGN "
        gdBase.Execute cSQL, dbFailOnError
        'Import Ende
    End If
    
    loeschNEW "Sortmet", dabalokal
    CreateTable "SORTMET", dabalokal
    
    cSQL = "INSERT into SORTMET Select B.WGN, B.WGNTEXT "
    cSQL = cSQL & " ,0.0 as UMSakt"
    cSQL = cSQL & " ,0.0 as MENGakt "
    cSQL = cSQL & " ,0.0 as EKPRakt  "
    cSQL = cSQL & " from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ALLLUG", dabalokal
    CreateTable "ALLLUG", dabalokal
    
    cSQL = "Insert into ALLLUG Select  artikel.artnr  "
    cSQL = cSQL & " ,0.0 as LAW"
    cSQL = cSQL & " ,0 as AGN"
    cSQL = cSQL & " ,0 as Bestand"
    cSQL = cSQL & " ,1 as LUG "
    cSQL = cSQL & "  from artikel "
    
    If lLinr > 0 Then
        cSQL = cSQL & "  inner join artlief on artikel.artnr = artlief.artnr "
        cSQL = cSQL & "  where artlief.linr = " & lLinr
    End If
    
    dabalokal.Execute cSQL, dbFailOnError
    
    
    
    If NewTableSuchenDBKombi("ALLARTLU", dabalokal) Then
        cSQL = "Update ALLLUG inner join ALLARTLU on ALLLUG.artnr = ALLARTLU.artnummer "
        cSQL = cSQL & " Set ALLLUG.LUG = ALLARTLU.LUG "
        dabalokal.Execute cSQL, dbFailOnError
    End If
        
    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.agn = artikel.agn "
    cSQL = cSQL & " , alllug.bestand = artikel.bestand "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set  alllug.LAW = alllug.Bestand * artikel.ekpr "
    cSQL = cSQL & " where alllug.bestand <> 0"
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "AL1", dabalokal
    
    
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,AVG(LUG) as ALUG "
    cSQL = cSQL & " ,Sum(LAW) as ALAW "
    cSQL = cSQL & " into AL1 from ALLLUG A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where C.Pname = '" & pname & "'"
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError

    loeschNEW "ttem", dabalokal
    'mwst V aktueller Zeitraum
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,'V'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStV & ")  as NUMSakt"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    
    
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis

    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'mwst E aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,'E'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStE & ")  as NUMSakt"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'mwst O aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,'O'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStO & ")  as NUMSakt"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis

    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select WGN "
    cSQL = cSQL & " ,SUM(UMSakt) as UMSaktS "
    cSQL = cSQL & " ,SUM(NUMSakt) as NUMSaktS "
    cSQL = cSQL & " ,SUM(MENGakt) as MENGaktS "
    cSQL = cSQL & " ,SUM(EKPRakt) as EKPRaktS "
    cSQL = cSQL & " ,AVG(NSP) as NSPM"
    cSQL = cSQL & " ,SUM(NSE) as NSES"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende

    cSQL = " Update SORTMET inner join  ttemsum  "
    cSQL = cSQL & " on SORTMEt.WGN = ttemsum.WGN  "
    cSQL = cSQL & " set SORTMET.UMSAKT = ttemsum.UMSAKTS "
    cSQL = cSQL & " , SORTMET.NUMSAKT = ttemsum.NUMSAKTS "
    cSQL = cSQL & " ,SORTMET.MENGAKT = ttemsum.MENGAKTS "
    cSQL = cSQL & " ,SORTMET.EKPRAKT = ttemsum.EKPRAKTS "
    cSQL = cSQL & " ,SORTMET.NSP = ttemsum.NSPM "
    cSQL = cSQL & " ,SORTMET.NSE = ttemsum.NSES "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMET inner join  ttem "
    cSQL = cSQL & " on SORTMEt.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTMET.nettoUmsVakt = ttem.NUMSakt "
    cSQL = cSQL & " where mwst = 'V'"
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMET inner join  ttem "
    cSQL = cSQL & " on SORTMEt.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTMET.nettoUmsEakt = ttem.NUMSakt "
    cSQL = cSQL & " where mwst = 'E'"
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMET inner join  ttem "
    cSQL = cSQL & " on SORTMEt.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTMET.nettoUmsOakt = ttem.NUMSakt "
    cSQL = cSQL & " where mwst = 'O'"
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    'Ermittle Gesamt aktueller Zeitraum
    
    loeschNEW "sghead", dabalokal
    CreateTable "SGHEAD", dabalokal
    
    cSQL = "INSERT INTO SGHEAD Select "
    cSQL = cSQL & " SUM(PREIS) as GUMSAKT "
    cSQL = cSQL & " ,SUM(Menge) as GMENGAKT "
    cSQL = cSQL & " ,SUM(MENGE * EKPR) as GEKPRAKT  "
    cSQL = cSQL & " from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "tt", dabalokal
    'V
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStV & ") as GUMSAKT_V "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='V'"
    dabalokal.Execute cSQL, dbFailOnError

    Dim dnetto As Double
    dnetto = 0
    cSQL = "Select GUMSAKT_V as maxi from tt"

    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead set GUMSAKT_V = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'V ende

    loeschNEW "tt", dabalokal
    'E
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStE & ") as GUMSAKT_E "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='E'"
    dabalokal.Execute cSQL, dbFailOnError

    dnetto = 0
    cSQL = "Select GUMSAKT_E as maxi from tt"
    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead set GUMSAKT_E = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'E ende
    
    loeschNEW "tt", dabalokal
    'Ohne
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStO & ") as GUMSAKT_O "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='O'"
    dabalokal.Execute cSQL, dbFailOnError

    dnetto = 0
    cSQL = "Select GUMSAKT_O as maxi from tt"
    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead set GUMSAKT_O = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'Ohne ende
    
    cSQL = " Update sghead set GUMSAKT_NETTO = GUMSAKT_O + GUMSAKT_E + GUMSAKT_V "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Kundenzahl Vorjahresbereich
    lKundenanzAkt = 0
    cSQL = "Select sum(kunz1) as maxi from umsatz"
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lKundenanzAkt = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    cSQL = " Update sghead "
    cSQL = cSQL & " set sghead.vondate = " & lVon
    cSQL = cSQL & " ,sghead.bisdate = " & lBis
    dabalokal.Execute cSQL, dbFailOnError
    
    '**********************Vorjahresrechnungen
    
    'Zeitraum Vorjahr
    
    lVon = datumwandlung(cVonvj)
    lBis = datumwandlung(cBisvj)
    
    loeschNEW "Sortme", dabalokal
    CreateTable "SORTME", dabalokal
    
    cSQL = "INSERT into SORTME Select B.WGN, B.WGNTEXT "
    cSQL = cSQL & " ,0.0 as UMSvj"
    cSQL = cSQL & " ,0.0 as MENGvj "
    cSQL = cSQL & " ,0.0 as EKPRvj  "
    cSQL = cSQL & " from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    

    'vorjahr volle Mwst
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,'V'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStV & ")  as NUMSvj"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr e Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,'E'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStE & ")  as NUMSvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr o Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,'O'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStO & ")  as NUMSvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select WGN "
    cSQL = cSQL & " ,SUM(UMSvj) as UMSvjS "
    cSQL = cSQL & " ,SUM(NUMSvj) as NUMSvjS "
    cSQL = cSQL & " ,SUM(MENGvj) as MENGvjS "
    cSQL = cSQL & " ,SUM(EKPRvj) as EKPRvjS "
    cSQL = cSQL & " ,AVG(NSPvj) as NSPvjM"
    cSQL = cSQL & " ,SUM(NSEvj) as NSEvjS"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende
    
    cSQL = " Update SORTME inner join  ttemsum  "
    cSQL = cSQL & " on SORTME.WGN = ttemsum.WGN  "
    cSQL = cSQL & " set SORTME.UMSvj = ttemsum.UMSvjS "
    cSQL = cSQL & " , SORTME.NUMSvj = ttemsum.NUMSvjS "
    cSQL = cSQL & " ,SORTME.MENGvj = ttemsum.MENGvjS "
    cSQL = cSQL & " ,SORTME.EKPRvj = ttemsum.EKPRvjS "
    cSQL = cSQL & " ,SORTME.NSPvj = ttemsum.NSPvjM "
    cSQL = cSQL & " ,SORTME.NSEvj = ttemsum.NSEvjS "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTME inner join  ttem "
    cSQL = cSQL & " on SORTME.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTME.nettoUmsV = ttem.NUMSvj "
    cSQL = cSQL & " where mwst = 'V'"
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTME inner join  ttem "
    cSQL = cSQL & " on SORTME.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTME.nettoUmsE = ttem.NUMSvj "
    cSQL = cSQL & " where mwst = 'E'"
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTME inner join  ttem "
    cSQL = cSQL & " on SORTME.WGN = ttem.WGN  "
    cSQL = cSQL & " set SORTME.nettoUmsO = ttem.NUMSvj "
    cSQL = cSQL & " where mwst = 'O'"
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    cSQL = " Update SORTME "
    cSQL = cSQL & " set SORTME.UMSAKT = 0 "
    cSQL = cSQL & " , SORTME.NUMSAKT = 0 "
    cSQL = cSQL & " ,SORTME.MENGAKT = 0 "
    cSQL = cSQL & " ,SORTME.EKPRAKT = 0 "
    cSQL = cSQL & " ,SORTME.NSP = 0 "
    cSQL = cSQL & " ,SORTME.NSE = 0 "
    cSQL = cSQL & " ,SORTME.LUG = 0 "
    cSQL = cSQL & " ,SORTME.LAGERW = 0 "
    cSQL = cSQL & " ,SORTME.nettoUmsVakt = 0 "
    cSQL = cSQL & " ,SORTME.nettoUmsEakt = 0 "
    cSQL = cSQL & " ,SORTME.nettoUmsOakt = 0 "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTME inner join  SortmeT  "
    cSQL = cSQL & " on SORTME.WGN = SORTMET.WGN  "
    cSQL = cSQL & " set SORTME.UMSAKT = SORTMET.UMSAKT "
    cSQL = cSQL & " ,SORTME.NUMSAKT = SORTMET.NUMSAKT "
    cSQL = cSQL & " ,SORTME.MENGAKT = SORTMET.MENGAKT "
    cSQL = cSQL & " ,SORTME.EKPRAKT = SORTMET.EKPRAKT "
    cSQL = cSQL & " ,SORTME.NSP = SORTMET.NSP "
    cSQL = cSQL & " ,SORTME.NSE = SORTMET.NSE "
    cSQL = cSQL & " ,SORTME.nettoUmsVakt = SORTMET.nettoUmsVakt "
    cSQL = cSQL & " ,SORTME.nettoUmsEakt = SORTMET.nettoUmsEakt "
    cSQL = cSQL & " ,SORTME.nettoUmsOakt = SORTMET.nettoUmsOakt "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTME inner join  AL1  "
    cSQL = cSQL & " on SORTME.WGN = AL1.WGN  "
    cSQL = cSQL & " set SORTME.LUG = AL1.ALUG "
    cSQL = cSQL & " , SORTME.LAGERW = AL1.ALAW "
    dabalokal.Execute cSQL, dbFailOnError

    loeschNEW "Sortmet", dabalokal
    
    loeschNEW "sghead1", dabalokal
    CreateTable "SGHEAD1", dabalokal
    
    cSQL = "INSERT INTO SGHEAD1 Select "
    cSQL = cSQL & " SUM(PREIS) as GUMSVJ "
    cSQL = cSQL & " ,SUM(Menge) as GMENGVJ "
    cSQL = cSQL & " ,SUM(MENGE * EKPR) as GEKPRVJ  "
    cSQL = cSQL & "  from Kassjour A "
    cSQL = cSQL & " where ums_ok = 'J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    dabalokal.Execute cSQL, dbFailOnError
    
    
    loeschNEW "tt", dabalokal
    'V
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStV & ") as GUMSvj_V "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='V'"
    dabalokal.Execute cSQL, dbFailOnError

    
    dnetto = 0
    cSQL = "Select GUMSvj_V as maxi from tt"

    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead1 set GUMSvj_V = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'V ende

    loeschNEW "tt", dabalokal
    'E
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStE & ") as GUMSvj_E "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='E'"
    dabalokal.Execute cSQL, dbFailOnError

    dnetto = 0
    cSQL = "Select GUMSvj_E as maxi from tt"
    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead1 set GUMSvj_E = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'E ende
    
    loeschNEW "tt", dabalokal
    'Ohne
    cSQL = "Select "
    cSQL = cSQL & " SUM(Preis * 100) /(100 + " & gdMWStO & ") as GUMSvj_O "
    cSQL = cSQL & " into tt from Kassjour A "
    cSQL = cSQL & " where ums_ok ='J'"
    If lLinr > 0 Then
        cSQL = cSQL & " and A.linr = " & lLinr
    End If
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " and MWST ='O'"
    dabalokal.Execute cSQL, dbFailOnError

    dnetto = 0
    cSQL = "Select GUMSvj_O as maxi from tt"
    Set rsrs = dabalokal.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dnetto = rsrs!maxi
        End If
    End If
    rsrs.Close

    cSQL = " Update sghead1 set GUMSvj_O = '" & dnetto & "'"
    dabalokal.Execute cSQL, dbFailOnError
    'Ohne ende
    
    cSQL = " Update sghead1 set GUMSvj_NETTO = GUMSvj_O + GUMSvj_E + GUMSvj_V "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    'Kundenzahl Vorjahresbereich
    lKundenanzVJ = 0
    cSQL = "Select sum(kunz1) as maxi from umsatz"
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lKundenanzVJ = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    dLUGavg = ermavgLUG

    cSQL = " Update sghead1 "
    cSQL = cSQL & " set sghead1.vondatevj = " & lVon
    cSQL = cSQL & " ,sghead1.bisdatevj = " & lBis
    cSQL = cSQL & " ,sghead1.Kundenakt = " & lKundenanzAkt
    cSQL = cSQL & " ,sghead1.Kundenvj = " & lKundenanzVJ
    cSQL = cSQL & " ,sghead1.LUGAVG = '" & dLUGavg & "'"
    cSQL = cSQL & " ,sghead1.ZUNAME = '" & pname & "'"
    dabalokal.Execute cSQL, dbFailOnError
    

    reportbildschirmApp "", "aWKL18a"

    anzeige "normal", "", Label2
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleAuswertung18"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub HoleAuswertung18Linr(pname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iJahr               As Integer
    Dim cVon                As String
    Dim cBis                As String
    Dim cVonvj              As String
    Dim cBisvj              As String
    Dim lVon                As Long
    Dim lBis                As Long
    Dim cSQL                As String
    Dim cdabapfad           As String
    Dim rsrs                As Recordset
    Dim i                   As Integer
    Dim sFilialauswahl      As String
    Dim lKundenanzAkt       As Long
    Dim lKundenanzVJ        As Long
    Dim dLUGavg             As Double
    
'    If pname = "Kiss Test" Then
'        Option1(4).Value = True
'        Text1(2).Text = "01.01.2011"
'        Text1(3).Text = "31.01.2011"
'    End If
    
    If Len(Text1(2).Text) <> 10 Or Len(Text1(3).Text) <> 10 Then
        anzeige "rot", "Bitte geben Sie das Datum im richtigen Format ein z.B.: 03.11.2014", Label2
        Exit Sub
    End If
    
    iJahr = Right$(Text1(2).Text, 4)
    iJahr = iJahr - 1
    cVon = Text1(2).Text
    cVonvj = Left$(Text1(2).Text, 6) & CStr(iJahr)
    
    iJahr = Right$(Text1(3).Text, 4)
    iJahr = iJahr - 1
    cBis = Text1(3).Text
    cBisvj = Left$(Text1(3).Text, 6) & CStr(iJahr)

    'Zeitraum aktuelles Jahr
    lVon = datumwandlung(cVon)
    lBis = datumwandlung(cBis)
    If lVon > lBis Then
        anzeige "rot", "Der Zeitraumbeginn ist größer als das Zeitraumende.", Label2
        
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", Label2
    
    
    Datenbankwechsel
    Me.Refresh
    
    loeschNEW "FILAW", dabalokal
    CreateTableT2 "FILAW", dabalokal
    
    sFilialauswahl = gRegister.firma
    
    cSQL = "Insert into FILAW (FILWAHL) values ('" & sFilialauswahl & "') "
    dabalokal.Execute cSQL, dbFailOnError
    
    cdabapfad = gcPfad
    If Right$(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "lokal.mdb"
    
    
    If glLokalAktuZeit = 0 Then
               
    Else
        'Import WGNDEF und WGNAGN
        loeschNEW "WGNDEF", dabalokal
        cSQL = "Select WGNDEF.* INTO WGNDEF IN '" & cdabapfad & "' from WGNDEF "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "WGNAGN", dabalokal
        cSQL = "Select WGNAGN.* INTO WGNAGN IN '" & cdabapfad & "' from WGNAGN "
        gdBase.Execute cSQL, dbFailOnError
        'Import Ende
    End If
    

    
    loeschNEW "ttem", dabalokal
    'mwst V aktueller Zeitraum
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'V'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStV & ")  as NUMSakt"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by a.Linr, c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    
    'mwst E aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'E'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStE & ")  as NUMSakt"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by a.Linr, c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'mwst O aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'O'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStO & ")  as NUMSakt"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis

    cSQL = cSQL & " group by a.Linr, c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select LINR,WGN "
    cSQL = cSQL & " ,SUM(UMSakt) as UMSaktS "
    cSQL = cSQL & " ,SUM(NUMSakt) as NUMSaktS "
    cSQL = cSQL & " ,SUM(MENGakt) as MENGaktS "
    cSQL = cSQL & " ,SUM(EKPRakt) as EKPRaktS "
    cSQL = cSQL & " ,AVG(NSP) as NSPM"
    cSQL = cSQL & " ,SUM(NSE) as NSES"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by LINR,WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende
    
    
    
    
    loeschNEW "SORTMETLINR", dabalokal
    CreateTableT2 "SORTMETLINR", dabalokal


    cSQL = "INSERT into SORTMETLINR Select LINR, '' as Liefbez, WGN, '' as WGNTEXT "
    cSQL = cSQL & " ,UMSaktS as UMSakt"
    cSQL = cSQL & " ,MENGaktS as MENGakt "
    cSQL = cSQL & " ,EKPRaktS as EKPRakt  "
    cSQL = cSQL & " from ttemsum "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    loeschNEW "tte", dabalokal
    cSQL = " Select  B.WGN, B.WGNTEXT "
    cSQL = cSQL & " into tte from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMETLINR s inner join  tte t "
    cSQL = cSQL & " on s.WGN = t.WGN   "
    cSQL = cSQL & " set s.WGNTEXT = t.WGNTEXT "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    
    
    loeschNEW "ttem", dabalokal
    
    
   
    
    '**********************Vorjahresrechnungen
    
    'Zeitraum Vorjahr
    
    lVon = datumwandlung(cVonvj)
    lBis = datumwandlung(cBisvj)
        
    loeschNEW "ttem", dabalokal
    

    'vorjahr volle Mwst
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'V'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStV & ")  as NUMSvj"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by a.linr,c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr e Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'E'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStE & ")  as NUMSvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by a.linr,c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr o Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,A.Linr "
    cSQL = cSQL & " ,'O'as MWST "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " , SUM(A.Preis * 100) /(100 + " & gdMWStO & ")  as NUMSvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
    cSQL = cSQL & " group by a.linr,c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select LINR,WGN "
    cSQL = cSQL & " ,SUM(UMSvj) as UMSvjS "
    cSQL = cSQL & " ,SUM(NUMSvj) as NUMSvjS "
    cSQL = cSQL & " ,SUM(MENGvj) as MENGvjS "
    cSQL = cSQL & " ,SUM(EKPRvj) as EKPRvjS "
    cSQL = cSQL & " ,AVG(NSPvj) as NSPvjM"
    cSQL = cSQL & " ,SUM(NSEvj) as NSEvjS"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by LINR,WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende
    
    loeschNEW "SORTMELINR", dabalokal
    CreateTableT2 "SORTMELINR", dabalokal

    cSQL = "INSERT into SORTMELINR Select LINR, '' as Liefbez, WGN, '' as WGNTEXT "
    cSQL = cSQL & " ,UMSvjS as UMSvj "
    cSQL = cSQL & " ,NUMSvjS as NUMSvj "
    cSQL = cSQL & " ,MENGvjS as MENGvj "
    
    cSQL = cSQL & " ,EKPRvjS as EKPRvj "
    cSQL = cSQL & " ,NSPvjM as NSPvj "
    cSQL = cSQL & " ,NSEvjS as NSEvj "
    
    cSQL = cSQL & " from ttemsum "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "tte", dabalokal
    cSQL = " Select  B.WGN, B.WGNTEXT "
    cSQL = cSQL & " into tte from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMELINR s inner join  tte t "
    cSQL = cSQL & " on s.WGN = t.WGN   "
    cSQL = cSQL & " set s.WGNTEXT = t.WGNTEXT "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    'hier neu 30.07.14
    cSQL = "INSERT into SORTMELINR Select LINR, '' as Liefbez, WGN, '' as WGNTEXT "
    cSQL = cSQL & " from SortmeTLINR where SORTMEtLINR.linr not in (Select linr from SORTMELINR)  "
    dabalokal.Execute cSQL, dbFailOnError
    'neu ende
    
    
    cSQL = " Update SORTMELINR s "
    cSQL = cSQL & " set s.UMSAKT = 0 "
    cSQL = cSQL & " , s.NUMSAKT = 0 "
    cSQL = cSQL & " ,s.MENGAKT = 0 "
    cSQL = cSQL & " ,s.EKPRAKT = 0 "
    cSQL = cSQL & " ,s.NSP = 0 "
    cSQL = cSQL & " ,s.NSE = 0 "
    cSQL = cSQL & " ,s.LUG = 0 "
    cSQL = cSQL & " ,s.LAGERW = 0 "
    cSQL = cSQL & " ,s.nettoUmsVakt = 0 "
    cSQL = cSQL & " ,s.nettoUmsEakt = 0 "
    cSQL = cSQL & " ,s.nettoUmsOakt = 0 "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMELINR s inner join  SortmeTLINR T  "
    cSQL = cSQL & " on s.WGN = t.WGN and s.LINR = t.LINR "
    cSQL = cSQL & " set s.UMSAKT = t.UMSAKT "
    cSQL = cSQL & " ,s.NUMSAKT = t.NUMSAKT "
    cSQL = cSQL & " ,s.MENGAKT = t.MENGAKT "
    cSQL = cSQL & " ,s.EKPRAKT = t.EKPRAKT "
    cSQL = cSQL & " ,s.NSP = t.NSP "
    cSQL = cSQL & " ,s.NSE = t.NSE "
    cSQL = cSQL & " ,s.nettoUmsVakt = t.nettoUmsVakt "
    cSQL = cSQL & " ,s.nettoUmsEakt = t.nettoUmsEakt "
    cSQL = cSQL & " ,s.nettoUmsOakt = t.nettoUmsOakt "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMELINR s inner join LISRT T on "
    cSQL = cSQL & " s.linr = t.linr "
    cSQL = cSQL & " SET s.liefbez = t.liefbez "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    

    loeschNEW "SortmeTLINR", dabalokal
    

    reportbildschirmApp "", "aWKL18b"

    anzeige "normal", "", Label2
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleAuswertung18Linr"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub HoleAuswertungWFS(pname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iJahr               As Integer
    Dim cVon                As String
    Dim cBis                As String
    Dim cVonvj              As String
    Dim cBisvj              As String
    Dim lVon                As Long
    Dim lBis                As Long
    Dim cSQL                As String
    Dim cdabapfad           As String
    Dim rsrs                As Recordset
    Dim i                   As Integer
    Dim sFilialauswahl      As String
    Dim lKundenanzAkt       As Long
    Dim lKundenanzVJ        As Long
    Dim dLUGavg             As Double
    
    If pname = "Kiss Test" Then
    
        If Len(Text1(2).Text) <> 10 Or Len(Text1(3).Text) <> 10 Then
    
    
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(3).Text = "31.12." & Year(Now) - 1
        
        End If
        
    End If
    
    If Len(Text1(2).Text) <> 10 Or Len(Text1(3).Text) <> 10 Then
        anzeige "rot", "Bitte geben Sie das Datum im richtigen Format ein z.B.: 03.11.2019", Label2
        Exit Sub
    End If
    
    iJahr = Right$(Text1(2).Text, 4)
    iJahr = iJahr - 1
    cVon = Text1(2).Text
    cVonvj = Left$(Text1(2).Text, 6) & CStr(iJahr)
    
    iJahr = Right$(Text1(3).Text, 4)
    iJahr = iJahr - 1
    cBis = Text1(3).Text
    cBisvj = Left$(Text1(3).Text, 6) & CStr(iJahr)

    'Zeitraum aktuelles Jahr
    lVon = datumwandlung(cVon)
    lBis = datumwandlung(cBis)
    If lVon > lBis Then
        anzeige "rot", "Der Zeitraumbeginn ist größer als das Zeitraumende.", Label2
        
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden ermittelt...", Label2
    
    
    Datenbankwechsel
    Me.Refresh
    

    loeschNEW "FILAW", dabalokal
    CreateTableT2 "FILAW", dabalokal
    
    sFilialauswahl = gRegister.firma

    
    Dim dEinkWertakt As Double
    Dim dEinkWertvj As Double
    Dim dEP As Double
    
    dEinkWertakt = 0
    dEinkWertvj = 0
    dEP = 0
    
    cSQL = "Insert into FILAW (FILWAHL) values ('" & sFilialauswahl & "') "
    dabalokal.Execute cSQL, dbFailOnError
    
    cdabapfad = gcPfad
    If Right$(cdabapfad, 1) <> "\" Then
        cdabapfad = cdabapfad & "\"
    End If
    cdabapfad = cdabapfad & "lokal.mdb"
    
    If glLokalAktuZeit = 0 Then
               
    Else
        'Import WGNDEF und WGNAGN
        loeschNEW "WGNDEF", dabalokal
        cSQL = "Select WGNDEF.* INTO WGNDEF IN '" & cdabapfad & "' from WGNDEF "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "WGNAGN", dabalokal
        cSQL = "Select WGNAGN.* INTO WGNAGN IN '" & cdabapfad & "' from WGNAGN "
        gdBase.Execute cSQL, dbFailOnError
        'Import Ende
    End If
    
    loeschNEW "SORTMET", dabalokal
    CreateTable "SORTMET", dabalokal
    
    cSQL = "INSERT into SORTMET Select B.WGN, B.WGNTEXT "
    cSQL = cSQL & " ,0.0 as UMSakt"
    cSQL = cSQL & " ,0.0 as MENGakt "
    cSQL = cSQL & " ,0.0 as EKPRakt  "
    cSQL = cSQL & " from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ALLLUG", dabalokal
    CreateTable "ALLLUG", dabalokal
    
    cSQL = "Insert into ALLLUG Select  artnr  "
    cSQL = cSQL & " ,0.0 as LAW "
    cSQL = cSQL & " ,0.0 as LAWVKNETTO "
    cSQL = cSQL & " ,0 as AGN "
    cSQL = cSQL & " ,0 as Bestand "
    cSQL = cSQL & " ,1 as LUG "
    cSQL = cSQL & "  from artikel "
    dabalokal.Execute cSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ALLARTLU", dabalokal) Then
        cSQL = "Update ALLLUG inner join ALLARTLU on ALLLUG.artnr = ALLARTLU.artnummer "
        cSQL = cSQL & " Set ALLLUG.LUG = ALLARTLU.LUG "
        dabalokal.Execute cSQL, dbFailOnError
    End If
        

    
    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.agn = artikel.agn "
    cSQL = cSQL & " , alllug.bestand = artikel.bestand "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.LAWVKNETTO = alllug.Bestand * (artikel.KVKPR1 * 100)/ (100 + " & gdMWStV & ") "
    cSQL = cSQL & " where artikel.MWST = 'V' "
    cSQL = cSQL & " and alllug.bestand <> 0"
    dabalokal.Execute cSQL, dbFailOnError

    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.LAWVKNETTO = alllug.Bestand * (artikel.KVKPR1 * 100)/ (100 + " & gdMWStE & ") "
    cSQL = cSQL & " where artikel.MWST = 'E' "
    cSQL = cSQL & " and alllug.bestand <> 0"
    dabalokal.Execute cSQL, dbFailOnError


    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.LAWVKNETTO = alllug.Bestand * artikel.KVKPR1 "
    cSQL = cSQL & " where artikel.MWST = 'O' "
    cSQL = cSQL & " and alllug.bestand <> 0"
    dabalokal.Execute cSQL, dbFailOnError

    cSQL = "Update ALLLUG inner join artikel on alllug.artnr = artikel.artnr "
    cSQL = cSQL & " set alllug.LAW = alllug.Bestand * artikel.ekpr "
    cSQL = cSQL & " where alllug.bestand <> 0"
    dabalokal.Execute cSQL, dbFailOnError
    
    
    loeschNEW "AL1", dabalokal
    
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,AVG(LUG) as ALUG "
    cSQL = cSQL & " ,Sum(LAW) as ALAW "
    cSQL = cSQL & " ,Sum(LAWVKNETTO) as ALAWVKNETTO "
    cSQL = cSQL & " into AL1 from ALLLUG A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where C.Pname = '" & pname & "'"
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError

    loeschNEW "ttem", dabalokal
    'mwst V aktueller Zeitraum
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"

    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'mwst E aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"

    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'mwst O aktueller Zeitraum
    cSQL = " Insert into ttem select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSakt "
    cSQL = cSQL & " ,SUM(A.Menge) as MENGakt "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRakt "
    cSQL = cSQL & " , AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSP"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSE"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on c.AGN = a.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"

    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select WGN "
    cSQL = cSQL & " ,SUM(UMSakt) as UMSaktS "
    cSQL = cSQL & " ,SUM(MENGakt) as MENGaktS "
    cSQL = cSQL & " ,SUM(EKPRakt) as EKPRaktS "
    cSQL = cSQL & " ,AVG(NSP) as NSPM"
    cSQL = cSQL & " ,SUM(NSE) as NSES"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende

    cSQL = " Update SORTMET inner join  ttemsum  "
    cSQL = cSQL & " on SORTMET.WGN = ttemsum.WGN  "
    cSQL = cSQL & " set SORTMET.UMSAKT = ttemsum.UMSAKTS "
    cSQL = cSQL & " ,SORTMET.MENGAKT = ttemsum.MENGAKTS "
    cSQL = cSQL & " ,SORTMET.EKPRAKT = ttemsum.EKPRAKTS "
    cSQL = cSQL & " ,SORTMET.NSP = ttemsum.NSPM "
    cSQL = cSQL & " ,SORTMET.NSE = ttemsum.NSES "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    'Ermittle Gesamt aktueller Zeitraum
    
    loeschNEW "sghead", dabalokal
    CreateTable "SGHEAD", dabalokal
    
    cSQL = "INSERT INTO SGHEAD Select "
    cSQL = cSQL & " SUM(PREIS) as GUMSAKT "
    cSQL = cSQL & " ,SUM(Menge) as GMENGAKT "
    cSQL = cSQL & " ,SUM(MENGE * EKPR) as GEKPRAKT  "
    cSQL = cSQL & " from Kassjour "
    cSQL = cSQL & " where ums_ok ='J'"
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis

    
    dabalokal.Execute cSQL, dbFailOnError
    
    'Kundenzahl Vorjahresbereich
    lKundenanzAkt = 0
    cSQL = "Select sum(kunz1) as maxi from umsatz"

    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lKundenanzAkt = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    cSQL = " Update sghead "
    cSQL = cSQL & " set sghead.vondate = " & lVon
    cSQL = cSQL & " ,sghead.bisdate = " & lBis
    dabalokal.Execute cSQL, dbFailOnError
    
    '**********************Vorjahresrechnungen
    
    'Zeitraum Vorjahr
    
    lVon = datumwandlung(cVonvj)
    lBis = datumwandlung(cBisvj)
    
    loeschNEW "SORTMEWFS", dabalokal
    CreateTableT2 "SORTMEWFS", dabalokal
    
    cSQL = "INSERT into SORTMEWFS Select B.WGN, B.WGNTEXT "
    cSQL = cSQL & " ,0.0 as UMSvj"
    cSQL = cSQL & " ,0.0 as MENGvj "
    cSQL = cSQL & " ,0.0 as EKPRvj  "
    cSQL = cSQL & " from WGNDEF B, WGNAGN C "
    cSQL = cSQL & " where C.WGN = B.WGN "
    cSQL = cSQL & " and B.Pname = '" & pname & "'"
    cSQL = cSQL & " group by B.WGN, B.WGNTEXT,B.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    'vorjahr volle Mwst
    cSQL = " Select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStV & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStV & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStV & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " into ttem from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'V'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr e Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStE & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStE & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStE & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'E'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'Vorjahr o Mwst
    cSQL = " Insert into ttem Select C.WGN "
    cSQL = cSQL & " ,SUM(A.PREIS) as UMSvj"
    cSQL = cSQL & " ,SUM(A.Menge) as MENGvj "
    cSQL = cSQL & " ,SUM(A.MENGE * A.EKPR) as EKPRvj  "
    cSQL = cSQL & " ,AVG( ((((A.Preis/(100 + " & gdMWStO & "))* 100) - (A.EKPR * A.Menge))* 100) / ((A.Preis/(100 + " & gdMWStO & "))* 100)) as NSPvj"
    cSQL = cSQL & " , SUM((A.Preis * 100) /(100 + " & gdMWStO & ") - (A.EKPR * A.Menge)) as NSEvj"
    cSQL = cSQL & " from Kassjour A left join WGNAGN C "
    cSQL = cSQL & " on A.AGN = C.AGN "
    cSQL = cSQL & " where A.ums_ok <> 'N'"
    cSQL = cSQL & " and MWST = 'O'"
    cSQL = cSQL & " and A.Preis <> 0"
    cSQL = cSQL & " and C.Pname = '" & pname & "'"
    cSQL = cSQL & " and A.ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"
    cSQL = cSQL & " group by c.WGN, c.Pname "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden
    loeschNEW "ttemsum", dabalokal
    
    cSQL = " Select WGN "
    cSQL = cSQL & " ,SUM(UMSvj) as UMSvjS "
    cSQL = cSQL & " ,SUM(MENGvj) as MENGvjS "
    cSQL = cSQL & " ,SUM(EKPRvj) as EKPRvjS "
    cSQL = cSQL & " ,AVG(NSPvj) as NSPvjM"
    cSQL = cSQL & " ,SUM(NSEvj) as NSEvjS"
    cSQL = cSQL & " into ttemsum from ttem "
    cSQL = cSQL & " group by WGN "
    dabalokal.Execute cSQL, dbFailOnError
    
    'summe bilden ende
    
    cSQL = " Update SORTMEWFS inner join  ttemsum  "
    cSQL = cSQL & " on SORTMEWFS.WGN = ttemsum.WGN  "
    cSQL = cSQL & " set SORTMEWFS.UMSvj = ttemsum.UMSvjS "
    cSQL = cSQL & " ,SORTMEWFS.MENGvj = ttemsum.MENGvjS "
    cSQL = cSQL & " ,SORTMEWFS.EKPRvj = ttemsum.EKPRvjS "
    cSQL = cSQL & " ,SORTMEWFS.NSPvj = ttemsum.NSPvjM "
    cSQL = cSQL & " ,SORTMEWFS.NSEvj = ttemsum.NSEvjS "
    dabalokal.Execute cSQL, dbFailOnError
    
    loeschNEW "ttem", dabalokal
    
    cSQL = " Update SORTMEWFS "
    cSQL = cSQL & " set SORTMEWFS.UMSAKT = 0 "
    cSQL = cSQL & " ,SORTMEWFS.MENGAKT = 0 "
    cSQL = cSQL & " ,SORTMEWFS.EKPRAKT = 0 "
    cSQL = cSQL & " ,SORTMEWFS.NSP = 0 "
    cSQL = cSQL & " ,SORTMEWFS.NSE = 0 "
    cSQL = cSQL & " ,SORTMEWFS.LUG = 0 "
    cSQL = cSQL & " ,SORTMEWFS.LAGERW = 0 "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMEWFS inner join  SORTMET  "
    cSQL = cSQL & " on SORTMEWFS.WGN = SORTMET.WGN  "
    cSQL = cSQL & " set SORTMEWFS.UMSAKT = SORTMET.UMSAKT "
    cSQL = cSQL & " ,SORTMEWFS.MENGAKT = SORTMET.MENGAKT "
    cSQL = cSQL & " ,SORTMEWFS.EKPRAKT = SORTMET.EKPRAKT "
    cSQL = cSQL & " ,SORTMEWFS.NSP = SORTMET.NSP "
    cSQL = cSQL & " ,SORTMEWFS.NSE = SORTMET.NSE "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMEWFS inner join  AL1  "
    cSQL = cSQL & " on SORTMEWFS.WGN = AL1.WGN  "
    cSQL = cSQL & " set SORTMEWFS.LUG = AL1.ALUG "
    cSQL = cSQL & " , SORTMEWFS.LAGERW = AL1.ALAW "
    cSQL = cSQL & " , SORTMEWFS.LAGERWVKNETTO = AL1.ALAWVKNETTO "
    dabalokal.Execute cSQL, dbFailOnError
    
    
    cSQL = " Update SORTMEWFS set LAGERWERTRAG = LAGERWVKNETTO - LAGERW "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMEWFS set LAGERWNSP = LAGERWERTRAG *100 /LAGERWVKNETTO  "
    cSQL = cSQL & " where LAGERWVKNETTO <> 0 "
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = " Update SORTMEWFS set UMSAKTNETTO = EKPRAKT + NSE "
    dabalokal.Execute cSQL, dbFailOnError
  
    dEP = dEinkWertakt - dEinkWertvj
    
    cSQL = " Update SORTMEWFS set DECKUNG = (UMSAKTNETTO - ('" & dEP & "' *100))/UMSAKTNETTO"
    dabalokal.Execute cSQL, dbFailOnError

    loeschNEW "SORTMET", dabalokal
    
    loeschNEW "sghead1", dabalokal
    CreateTable "SGHEAD1", dabalokal
    
    cSQL = "INSERT INTO SGHEAD1 Select "
    cSQL = cSQL & " SUM(PREIS) as GUMSVJ "
    cSQL = cSQL & " ,SUM(Menge) as GMENGVJ "
    cSQL = cSQL & " ,SUM(MENGE * EKPR) as GEKPRVJ  "
    cSQL = cSQL & "  from Kassjour "
    cSQL = cSQL & " where ums_ok = 'J'"
    cSQL = cSQL & " and ADATE between " & lVon & " and " & lBis
'    cSQL = cSQL & " and (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"
    dabalokal.Execute cSQL, dbFailOnError
    
    'Kundenzahl Vorjahresbereich
    lKundenanzVJ = 0
    cSQL = "Select sum(kunz1) as maxi from umsatz"
'    cSQL = cSQL & " where (filiale=" & Mid$(lstAusgFilialen.list(0), 1, InStr(1, lstAusgFilialen.list(0), " "))
'    For i = 1 To lstAusgFilialen.ListCount - 1
'        cSQL = cSQL & " or filiale=" & Mid$(lstAusgFilialen.list(i), 1, InStr(1, lstAusgFilialen.list(i), " "))
'    Next i
'    cSQL = cSQL & ")"
    cSQL = cSQL & " where DATUM between " & lVon & " and " & lBis
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lKundenanzVJ = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    'dLUGavg Mittelwert LUG
    dLUGavg = ermavgLUG
'    cSQL = "Select avg(LUG) as maxi from ALLARTLU"
'    Set rsrs = gdbMdb.OpenRecordset(cSQL)
'    If Not rsrs.EOF Then
'        If Not IsNull(rsrs!Maxi) Then
'            dLUGavg = rsrs!Maxi
'        End If
'    End If
'    rsrs.Close

    cSQL = " Update sghead1 "
    cSQL = cSQL & " set sghead1.vondatevj = " & lVon
    cSQL = cSQL & " ,sghead1.bisdatevj = " & lBis
    cSQL = cSQL & " ,sghead1.Kundenakt = " & lKundenanzAkt
    cSQL = cSQL & " ,sghead1.Kundenvj = " & lKundenanzVJ
    cSQL = cSQL & " ,sghead1.LUGAVG = '" & dLUGavg & "'"
    cSQL = cSQL & " ,sghead1.ZUNAME = '" & pname & "'"
    dabalokal.Execute cSQL, dbFailOnError
    

    reportbildschirmApp "", "aWKL18f"

    anzeige "normal", "", Label2
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleAuswertungWFS"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub NeuWgnSpeichern(cprofil As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cFeld1 As String
    Dim cFeld2 As String
    Dim bTrans As Boolean
    Dim lcount As Long
    
    cFeld1 = Text4(0).Text
    cFeld1 = Trim$(cFeld1)
    
    cFeld2 = Text4(1).Text
    cFeld2 = Trim$(cFeld2)
    
    cSQL = "Insert into WGNDEF (WGN, WGNTEXT, Pname) values (" & cFeld1 & ", '" & cFeld2 & "','" & cprofil & "' ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cFeld1 = Space$(3 - Len(cFeld1)) & cFeld1
    List2.AddItem cFeld1 & " " & cFeld2
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NeuWgnSpeichern"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
  Select Case Index
        Case Is = 20        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(2000, 4000), "DD.MM.YYYY")
        Case Is = 21        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(2000, 4000), "DD.MM.YYYY")
            'fertig
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
        
    Select Case Index
        Case Is = 0     'Speichern
            SchreibeWGN
        Case Is = 3     'Zurück
            profilnamespeichern
            Text4(0).Text = ""
            Text4(1).Text = ""
            List8.Clear
            Frame10.Visible = False
            
            Profilfüllen
            Frame2.Visible = True
        Case Is = 4     'AGN aus WGN entfernen
            DeleteAGNFromWGN Label30.Caption, Label11.Caption
            aktualiAGN Label11.Caption
        Case Is = 5     'AGN einer WGN zuordnen
            InsertAGN2WGN Label30.Caption, Label11.Caption
'            aktualiAGN Label11.Caption
        Case Is = 7 'Warengruppe löschen
            Select Case Command2(7).Caption
                Case Is = "Warengruppe löschen"
                    If List1.ListIndex < 0 Then
                        Label29.ForeColor = vbRed
                        Label29.Caption = "Bitte einen Eintrag auswählen!"
                        Label29.Refresh
                        
                        Exit Sub
                    End If
                
                    LoeschWGN List1, Label25.Caption, Left(List1.list(List1.ListIndex), 3)
                    Label29.Caption = "Wählen Sie eine Warengruppe!"
                    Label29.Refresh
                    
                Case Is = "Warengruppe bearbeiten"
                    If List1.ListIndex < 0 Then
                        Label29.ForeColor = vbRed
                        Label29.Caption = "Bitte einen Eintrag auswählen!"
                        Label29.Refresh
                        
                        Exit Sub
                    End If
                    
                    Label8.Caption = "Warengruppe Bearbeiten"
                    Label8.Refresh
                    
                    'WGN Nummer
                    Label30.Caption = Trim(Left(List1.list(List1.ListIndex), 3))
                    Label30.Refresh
                    
                    'WGN Bezeichnung
                    Label35.Caption = Trim(Right(List1.list(List1.ListIndex), Len(List1.list(List1.ListIndex)) - 3))
                    Label35.Refresh
                    
                    Label11.Caption = Trim(Label25.Caption)
                    Label11.Refresh
                    
                    'welche AGN´s stehen in diesem Profil noch zur Verfügung
                    aktualiAGN Label25.Caption
                    SelektWarengruppe Left(List1.list(List1.ListIndex), 3), Label25.Caption
                    List2DelList1
                    Frame10.Visible = True
                    Label36.Caption = ""
                    Label36.Refresh
                    Frame2.Visible = False
                    Frame1.Visible = False
            End Select
        
    End Select
            
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub aktualiAGN(cprofil As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim cLBSatz As String

    loeschNEW "agnAusw", gdBase
    
    cSQL = "Select distinct AGNDBF.AGN ,AGNDBF.agText into agnAusw from AGNDBF"
    cSQL = cSQL & " order by agndbf.AGN"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from agnAusw where agn in ( Select agn from WGNAGN where Pname = '" & cprofil & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    List4.Clear
    
    cSQL = "Select * from AGNAUSW order by AGN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AGN) Then
                ctmp = rsrs!AGN
            Else
                ctmp = ""
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(5 - Len(ctmp)) & ctmp
            cLBSatz = ctmp & " "
            If Not IsNull(rsrs!AGTEXT) Then
                ctmp = rsrs!AGTEXT
            Else
                ctmp = ""
            End If
            
            cLBSatz = cLBSatz & ctmp
            List4.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "aktualiAGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoeschWGN(list As ListBox, cprofil As String, cWGN As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Delete from WGNAGN where WGN = " & cWGN & " and Pname = '" & cprofil & "'"
    gdBase.Execute cSQL, dbFailOnError
    cSQL = "Delete from WGNDEF where WGN = " & cWGN & " and Pname = '" & cprofil & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    list.RemoveItem list.ListIndex
    list.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoeschWGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub profilnamespeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If gcPname <> "" Then
        sSQL = " Update wgndef set Pname = '" & gcPname & "'"
        sSQL = sSQL & " where Pname is null "
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Delete from wgndef "
        sSQL = sSQL & " where Pname is null "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Delete from wgnagn "
    sSQL = sSQL & " where Pname is null "
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "profilnamespeichern"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0 'Neu
            Label7.Caption = ""
            Label7.Refresh
            Frame3.Visible = True
            Label4.Caption = "Vorgang: Neu" & vbCrLf & vbCrLf
            Label4.Caption = Label4.Caption & "Vergeben Sie einen Namen für Ihre Zusammenstellung!"
            Label4.Refresh
            txtProfilname.Text = ""
            txtProfilname.SetFocus
            
        Case Is = 1 'Bea
            Frame5.Visible = True
            Label9.Caption = "Vorgang: Bearbeitung" & vbCrLf & vbCrLf
            Label9.Caption = Label9.Caption & "Wählen Sie eine Zusammenstellung aus!"
            Label9.Refresh
            Label12.Caption = ""
            Label12.Refresh
        Case Is = 2 'Löschen
            Frame5.Visible = True
            Label9.Caption = "Vorgang: Löschen" & vbCrLf & vbCrLf
            Label9.Caption = Label9.Caption & "Wählen Sie eine Zusammenstellung aus!"
            Label9.Refresh
            Label12.Caption = ""
            Label12.Refresh
        
        Case Is = 3 'Auswertung
            Frame5.Visible = True
            Label9.Caption = "Vorgang: Auswertung" & vbCrLf & vbCrLf
            Label9.Caption = Label9.Caption & "Wählen Sie eine Zusammenstellung aus!"
            Label9.Refresh
            Label12.Caption = ""
            Label12.Refresh
            
        Case Is = 4 'schließen
            Unload frmWKL18
            
            
        Case Is = 5 'WGN Neu
            Frame4.Visible = True
            
            Label5.Caption = "Vorgang: Warengruppe neu"
            Label5.Refresh
            Label31.Caption = Label24.Caption
            Label31.Refresh
            Label34.Caption = ""
            Label34.Refresh
            Text4(0).Text = ""
            Text4(1).Text = ""
            Text4(0).SetFocus
            'listeWgnsfüllen
            FuelleWGNLIST Label31.Caption, List2
            
            
        Case Is = 6 'WGN Bearbeiten
            Command2(7).Caption = "Warengruppe bearbeiten"
            Command2(7).Refresh
            Label26.Caption = "Vorgang: Warengruppe bearbeiten"
            Label26.Refresh
            Label25.Caption = Label24.Caption
            Label25.Refresh
            Frame2.Visible = True
            
            'listeWgnsfüllen
            FuelleWGNLIST Label25.Caption, List1
            
        Case Is = 7 'WGN Löschen
            Command2(7).Caption = "Warengruppe löschen"
            Command2(7).Refresh
            Label26.Caption = "Vorgang: Warengruppe löschen"
            Label26.Refresh
            Label25.Caption = Label24.Caption
            Label25.Refresh
            Frame2.Visible = True
            
            'listeWgnsfüllen
            FuelleWGNLIST Label25.Caption, List1
            
        Case Is = 8 'Testschlüssel erstellen
            TestSchluessel
            Profilfüllen
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command5_Click 4
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
        
    Select Case Index
    
        Case 17 'AGN Bereinigen
            AGN_Bereinigen
        Case Is = 7 'Zurück aus WGN Neu
            Frame4.Visible = False
        Case Is = 6 'Zurück aus WGN löschen
            Frame2.Visible = False
            Frame1.Visible = True
        Case Is = 4 'Zurück
            Frame3.Visible = False
        Case Is = 3 'Auswertung
            
            HoleAuswertung18 gcPname
        Case Is = 8 'Analyse
            
            AGNANALYSE
        Case 15
            ZeigeArtikelAGN
        Case 16 'Lieferantensicht Auswertung
            
            HoleAuswertung18Linr gcPname
            
            
        Case Is = 5 'Zurück aus WGN OberFrage
            Select Case Label23.Caption
                Case Is = "Zusammenstellung Bearbeiten:"
                    Frame1.Visible = False
                    Frame5.Visible = True
                Case Is = "Zusammenstellung Neu:"
                    Frame1.Visible = False
                    Frame3.Visible = True
            End Select
        Case Is = 2 'Zurück
            Frame5.Visible = True
            Frame6.Visible = False
        Case Is = 1 'Zurück
            Frame5.Visible = False
        Case Is = 0 'Übernehmen
            If List12.ListIndex < 0 Then
                Label12.ForeColor = vbRed
                Label12.Caption = "Bitte einen Eintrag auswählen!"
                Label12.Refresh
                
                Exit Sub
            End If
            
            Select Case Left(Label9.Caption, 10)
                Case Is = "Vorgang: B" 'Bearbeitung
                
                    Screen.MousePointer = 11
                    gcPname = Trim(List12.list(List12.ListIndex))

                    Label24.Caption = gcPname
                    Label24.Refresh
                    Label23.Caption = "Zusammenstellung Bearbeiten:"
                    Label23.Refresh
                    
                    Frame1.Visible = True
                    Frame5.Visible = False
                    Screen.MousePointer = 0
                Case Is = "Vorgang: L"
                    Label12.Caption = ""
                    Label12.Refresh
                    gcPname = Trim(List12.list(List12.ListIndex))
                    
                    sSQL = "Delete from wgndef where pname = '" & gcPname & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Delete from WGNagn where pname = '" & gcPname & "' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    Profilfüllen
                    
                Case Is = "Vorgang: A"
                    Label2.Caption = ""
                    Label2.Refresh
                    Label10.Caption = "Vorgang: Auswertung" & vbCrLf & vbCrLf
                    Label10.Caption = Label10.Caption & "Wählen Sie einen Zeitraum aus!"
                    Label10.Refresh
                    gcPname = Trim(List12.list(List12.ListIndex))
                    Frame6.Visible = True
                    Frame5.Visible = False
            End Select
            
        Case Is = 9 'Testschluessel erstellen
            WFSSchluessel
            Profilfüllen
        Case Is = 10 'Testschluessel erstellen
            TCSchluessel
            Profilfüllen
        Case Is = 11 'Testschluessel erstellen
            CPSchluessel
            Profilfüllen
        Case Is = 12
            HoleAuswertungWFS gcPname
        Case Is = 13
            HoleAuswertungWFS gcPname
        Case Is = 14
            HoleAuswertungWFS gcPname
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeArtikelAGN()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cVon As String
    Dim cBis As String
    Dim lBis As Long
    Dim lVon As Long
    
    Screen.MousePointer = 11
    
    loeschNEW "KASSNOAGN", gdBase
    CreateTableT2 "KASSNOAGN", gdBase
    
    sSQL = "Insert into KASSNOAGN Select distinct(artnr) as art ,bezeich , Kassjour.agn ,Linr,lpz,'' as VON ,'' as BIS  from Kassjour inner join agnana "
    sSQL = sSQL & " on Kassjour.agn = agnana.agn "
    
    If Len(Text1(2).Text) = 10 And Len(Text1(3).Text) = 10 Then
        cVon = Text1(2).Text
        cBis = Text1(3).Text
        lVon = datumwandlung(cVon)
        lBis = datumwandlung(cBis)
        
        sSQL = sSQL & " where Kassjour.ADATE between " & lVon & " and " & lBis
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    If Len(Text1(2).Text) = 10 And Len(Text1(3).Text) = 10 Then
        
        sSQL = "Update KASSNOAGN set VON =  '" & cVon & "'"
        sSQL = sSQL & " ,BIS =  '" & cBis & "'"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    
    Markenabgleich "KASSNOAGN", gdBase
    
    reportbildschirm "", "aWKL18g"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeArtikelAGN"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WFSSchluessel()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim i As Integer
    
    Dim sPname  As String
    sPname = "WIR FÜR SIE"
    
    cSQL = "Delete from WGNDEF where pname = '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from WGNAGN where pname =  '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (281,'Gesichtspflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (282,'Körperpflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (283,'Gesichtspflege Herren', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (284,'Dekorative Kosmetik', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (285,'Damendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (286,'Herrendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    
    For i = 111 To 118
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (281," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 121 To 126
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (282," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 131 To 139
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (284," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 141 To 148
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (285," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 161 To 163
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (286," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (283,164, '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 165 To 169
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (286," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 181 To 188
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (285," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WFSSchluessel"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TCSchluessel()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim i As Integer
    
    Dim sPname  As String
    sPname = "TopCos"
    
    cSQL = "Delete from WGNDEF where pname = '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from WGNAGN where pname =  '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (381,'Gesichtspflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (382,'Körperpflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (383,'Gesichtspflege Herren', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (384,'Dekorative Kosmetik', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (385,'Damendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (386,'Herrendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    

    
    For i = 111 To 118
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (381," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 121 To 126
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (382," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 131 To 139
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (384," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 141 To 148
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (385," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 161 To 163
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (386," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (383,164, '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 165 To 169
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (386," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 181 To 188
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (385," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TCSchluessel"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub CPSchluessel()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim i As Integer
    
    Dim sPname  As String
    sPname = "COSPAR"
    
    cSQL = "Delete from WGNDEF where pname = '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from WGNAGN where pname =  '" & sPname & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (481,'Gesichtspflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (482,'Körperpflege Damen', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (483,'Gesichtspflege Herren', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (484,'Dekorative Kosmetik', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (485,'Damendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WGNDEF (wgn,wgntext,pname) values (486,'Herrendüfte', '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 111 To 118
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (481," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 121 To 126
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (482," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 131 To 139
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (484," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 141 To 148
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (485," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 161 To 163
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (486," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (483,164, '" & sPname & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    For i = 165 To 169
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (486," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    For i = 181 To 188
        cSQL = "Insert into WGNAGN (wgn,AGN,pname) values (485," & i & ", '" & sPname & "')"
        gdBase.Execute cSQL, dbFailOnError
    Next i
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CPSchluessel"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    Dim i As Integer
    Dim cPfad As String
    Dim lRet As Long
    
    WKL18Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    For i = 0 To 7
        Label38(i).ForeColor = glU1
    Next i
    Profilfüllen
    
    cmdHelp.Visible = istHilfeda(Me.Tag)

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AGN_Bereinigen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset

    Screen.MousePointer = 11
    
    anzeige "normal", "Daten werden bereinigt, bitte warten...", Label2
    
    
    Dim cJahr As String
    Dim i As Integer
    
    loeschNEW "Agnana", gdBase
    
    sSQL = " Select distinct year(adate) as Jahrangabe into agnana from Kassjour "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("agnana", gdBase) = True Then
        
        Set rs = gdBase.OpenRecordset("Select * from agnana order by Jahrangabe asc")
        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                If Not IsNull(rs!Jahrangabe) Then
                    cJahr = rs!Jahrangabe
                    
                    For i = 1 To 12
                    
                        anzeige "normal", "Jahr: " & cJahr & " Monat: " & i & ", Daten werden bereinigt, bitte warten...", Label2
                        
                        
                        sSQL = "Update Kassjour inner join Artikel on Kassjour.artnr = Artikel.artnr "
                        sSQL = sSQL & " set Kassjour.AGN = Artikel.AGN where artikel.agn > 0"
                        
                        sSQL = sSQL & " and year(adate) = " & cJahr & ""
                        sSQL = sSQL & " and month(adate) = " & i & ""
                        gdBase.Execute sSQL, dbFailOnError
                    Next i
                    
                    
                    
                    
                    
                End If
                rs.MoveNext
            Loop
        End If
        rs.Close: Set rs = Nothing
    End If
    
    
    
    
    
    
    
    anzeige "normal", "Fertig!", Label2

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AGN_Bereinigen"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AGNANALYSE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    Dim ctmp As String
    ctmp = "Diese Artikelgruppennummern(AGN) wurden in der Kassenjournaldatei ohne Übereinstimmung"
    ctmp = ctmp & " in der AGN - Tabelle ermittelt. Die Auswertung kann daher nicht zum korrektem Ergebnis führen."
    
    Screen.MousePointer = 11
    
    loeschNEW "Agnana", gdBase
    
    sSQL = " Select distinct (AGN) into agnana from Kassjour "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Delete from agnana where agn in (Select agn from AGNdbf )"
    gdBase.Execute sSQL, dbFailOnError
    
    List3.Clear
    
    If Datendrin("agnana", gdBase) = True Then
        
        Set rs = gdBase.OpenRecordset("agnana", dbOpenTable)
        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                If Not IsNull(rs!AGN) Then
                    List3.AddItem rs!AGN
                End If
                rs.MoveNext
            Loop
            List3.Visible = True
            Command6(15).Visible = True
            Label1(0).Visible = True
            Label1(0).Caption = ctmp
            Label1(0).Refresh
        End If
        rs.Close: Set rs = Nothing
        
    Else
        Label1(0).Visible = False
        Command6(15).Visible = False
        List3.Visible = False
    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AGNANALYSE"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2DelList1()
On Error GoTo LOKAL_ERROR

Dim lCount1         As Long
Dim lCount2         As Long
Dim sVergleich      As String
    
For lCount1 = 0 To List8.ListCount - 1
    sVergleich = List8.list(lCount1)
    For lCount2 = 0 To List4.ListCount - 1
        If List4.list(lCount2) = sVergleich Then
            List4.RemoveItem lCount2
            Exit For
        End If

    Next lCount2
Next lCount1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2DelList1"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List12_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command6_Click 0
    End If
    
    If KeyCode = vbKeyEscape Then
        Command6_Click 1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List12_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List8_dblClick()
    On Error GoTo LOKAL_ERROR
    
    Command2_Click 4
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List8_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 4     'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(2).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(3).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(2).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(3).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    Case 2
                        Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    Case Else
                        Text1(3).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
        Case Is = 5     'ak monat
            Text1(2).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case Is = 8     'vorjahrzr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM") & "." & Year(Now) - 1
        Case Is = 6     'vorjahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(3).Text = "31.12." & Year(Now) - 1
        Case Is = 7     'ak jahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Index = 7 Then
        LiefKuerzelAufloesung Label1(4), Text1(7)
    End If
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 7 Then
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            
            gF2Prompt.cFeld = "LINR"
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    Label1(4).Caption = gF2Prompt.cWert
                    Label1(4).Refresh
                End If
            End If
           
            Text1(Index).SetFocus
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text4(Index).BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text4(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtProfilname_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    txtProfilname.BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtProfilname_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtProfilname_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdNeu_Click
    End If
    
    If KeyCode = vbKeyEscape Then
        Command6_Click 4
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtProfilname_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtProfilname_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtProfilname.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtProfilname_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Sortimentsanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
