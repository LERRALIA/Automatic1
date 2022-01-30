VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWK16a 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "WE aus Bestellung"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmWK16a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "kumulieren"
      Height          =   255
      Left            =   6960
      TabIndex        =   85
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fraArtAnfuegen 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Caption         =   "Artikel Anfügen ( Bitte mindestens ein Feld ausfüllen ) :"
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
      Height          =   1095
      Left            =   3600
      TabIndex        =   44
      Top             =   720
      Width           =   3255
      Begin VB.TextBox Text6 
         Height          =   288
         Left            =   2880
         TabIndex        =   75
         Top             =   1680
         Width           =   3732
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   11175
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   2
         Left            =   6960
         TabIndex        =   52
         Top             =   6960
         Width           =   2172
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   49
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   2880
         TabIndex        =   48
         Top             =   960
         Width           =   2652
      End
      Begin VB.TextBox Text4 
         Height          =   288
         Left            =   2880
         TabIndex        =   47
         Top             =   1320
         Width           =   3732
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   0
         Left            =   9240
         TabIndex        =   46
         Top             =   1080
         Width           =   2172
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
         Caption         =   " Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   1
         Left            =   9240
         TabIndex        =   45
         Top             =   6960
         Width           =   2172
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
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   51
         Top             =   2040
         Width           =   11172
      End
      Begin VB.Label label8 
         BackColor       =   &H00808000&
         Caption         =   "Lieferantenbestellnummer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   76
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel anfügen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         TabIndex        =   63
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Artikelnummer :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         Caption         =   "EAN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "Artikelbezeichnung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   -2520
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox txtStatus 
         Height          =   255
         Left            =   2640
         TabIndex        =   71
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox picprogress 
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
         Left            =   3840
         ScaleHeight     =   75
         ScaleWidth      =   6915
         TabIndex        =   70
         Top             =   650
         Visible         =   0   'False
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   3
         Left            =   3360
         TabIndex        =   66
         Top             =   6360
         Width           =   2760
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
         Caption         =   "Zwischenspeichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8705
         _Version        =   393216
         BackColorSel    =   10485760
         ForeColorSel    =   65535
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   2
         Left            =   6120
         TabIndex        =   43
         Top             =   6360
         Width           =   2760
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
         Caption         =   "Artikel anfügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   8640
         TabIndex        =   32
         Top             =   120
         Width           =   2175
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
         Caption         =   "Geliefert + Berechnet auf Bestellt setzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   0
         Left            =   6360
         TabIndex        =   31
         Top             =   120
         Width           =   2265
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
         Caption         =   "Geliefert + Berechnet auf 0 setzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Left            =   5160
         TabIndex        =   30
         Top             =   120
         Width           =   1200
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Left            =   3840
         TabIndex        =   29
         Top             =   120
         Width           =   1305
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
         Caption         =   "Speichern"
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
         Height          =   480
         Left            =   2520
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   120
         Width           =   1215
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   1
         Left            =   8880
         TabIndex        =   9
         Top             =   6360
         Width           =   2775
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
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   6360
         Width           =   3240
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
         Caption         =   "Lieferung übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   78
         Top             =   5880
         Width           =   11415
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10920
         MouseIcon       =   "frmWK16a.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWK16a.frx":074C
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem Scanpal einlesen möchten"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   6120
         Width           =   11415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Zeilenrabatt über alles:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   9000
      TabIndex        =   68
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin sevCommand3.Command Command6 
      Height          =   300
      Left            =   10920
      TabIndex        =   67
      Top             =   240
      Visible         =   0   'False
      Width           =   855
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
      Caption         =   "Suche"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
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
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   11775
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   14
         Left            =   10410
         TabIndex        =   61
         Top             =   0
         Width           =   675
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   12
         Left            =   9000
         TabIndex        =   60
         Top             =   0
         Width           =   700
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   15
         Left            =   11080
         TabIndex        =   59
         Top             =   0
         Width           =   670
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   13
         Left            =   9700
         TabIndex        =   58
         Top             =   0
         Width           =   705
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   11
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   10
         Left            =   7320
         TabIndex        =   21
         Top             =   0
         Width           =   720
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
         Caption         =   ","
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   9
         Left            =   6600
         TabIndex        =   20
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   8
         Left            =   5880
         TabIndex        =   19
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   7
         Left            =   5160
         TabIndex        =   18
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   6
         Left            =   4440
         TabIndex        =   17
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   5
         Left            =   3720
         TabIndex        =   16
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   15
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Top             =   0
         Width           =   720
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
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   720
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
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   7
         Left            =   8040
         TabIndex        =   37
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   6
         Left            =   6600
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   5
         Left            =   5160
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   4
         Left            =   3720
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   3
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   1560
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12120
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10815
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   5640
         Width           =   7935
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Lieferantenbezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   82
            Top             =   120
            Width           =   3495
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   81
            Top             =   120
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
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
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   1815
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   6
         Left            =   8160
         TabIndex        =   77
         Top             =   4440
         Width           =   2535
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   5
         Left            =   8160
         TabIndex        =   74
         Top             =   3840
         Width           =   2535
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
         Caption         =   "Importieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   4
         Left            =   8160
         TabIndex        =   73
         Top             =   3240
         Width           =   2535
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
         Caption         =   "Exportieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   3
         Left            =   8160
         TabIndex        =   72
         Top             =   5040
         Visible         =   0   'False
         Width           =   2535
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
         Caption         =   "Kundenbestellungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Linien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   252
         Index           =   3
         Left            =   8160
         TabIndex        =   56
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Artikelnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   42
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Lieferantenbestellnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   41
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Bezeichnung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   40
         Top             =   840
         Value           =   -1  'True
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9120
         TabIndex        =   4
         Top             =   7200
         Width           =   2535
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
         Caption         =   "Schließe&n"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         Height          =   5100
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   39
         Top             =   480
         Width           =   7935
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
         Height          =   480
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   7935
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   8160
         TabIndex        =   3
         Top             =   2640
         Width           =   2535
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
         Caption         =   "&Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   8160
         TabIndex        =   2
         Top             =   2040
         Width           =   2535
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
         Caption         =   "Aus&wählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   7
         Left            =   10320
         TabIndex        =   86
         Top             =   120
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
         Picture         =   "frmWK16a.frx":0D2F
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelnummer / EAN"
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   84
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Wert aller Bestellungen"
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   83
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "sortiert nach"
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
         Left            =   8160
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   615
         Index           =   2
         Left            =   3480
         TabIndex        =   35
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
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
         Left            =   3480
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Bestellung vom ... bei:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
      End
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   11280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikelnummer / EAN"
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "WE aus Bestellung"
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
      TabIndex        =   62
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmWK16a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sZufall             As String
Dim gcWEdatei           As String
Dim cAnfuLinr           As String
Dim cSort               As String

Dim SpaltennummerArtnr As Byte
Dim SpaltennummerLEKPR As Byte
Dim SpaltennummerKVKPR1  As Byte
Dim SpaltennummerBESTELLT  As Byte
Dim SpaltennummerGELIEFERT  As Byte
Dim SpaltennummerBERECHNET  As Byte
Dim SpaltennummerBEZEICH As Byte
Dim SpaltennummerLIEFBETRAG As Byte
Dim SpaltennummerZEILEN_RAB   As Byte
Dim SpaltennummerZEILENWERT As Byte
Dim SpaltennummerRECHN_RAB  As Byte
Dim SpaltennummerRECHN_WERT  As Byte
Dim SpaltennummerSTCK_PREIS As Byte
Dim SpaltennummerLINR As Byte
Dim SpaltennummerLIBESNR As Byte
Dim SpaltennummerLPZ As Byte
Dim SpaltennummerMOPREIS As Byte
Dim SpaltennummerAWM As Byte

Dim gbAenderKVK As Boolean
Dim gbAenderLEKPR As Boolean
Dim gbAenderBEZEICH As Boolean
Dim gbAenderBESTELLT As Boolean
Dim gbAenderGELIEFERT As Boolean
Dim gbAenderBERECHNET As Boolean

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei      As String
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cDatum      As String
    Dim cLieferant  As String
    Dim iRet        As Integer
    Dim iZufall     As Integer
    Dim cTabelle    As String
    Dim dbBestell   As Database
    Dim cDatname    As String
    Dim sSQL        As String
    Dim lDatum      As Long
    Dim lZaehler    As Long
    Dim lcount      As Long
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case Index
        Case Is = 0
            voreinstellungspeichern
            
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
            Else
            
                cdatei = List2.list(List2.ListIndex)
                cLieferant = Mid(cdatei, 27, Len(cdatei) - 26)
                Label2(2).Caption = cLieferant
                cdatei = Trim(Left(cdatei, 13))
                gcWEdatei = ""
                gcWEdatei = cdatei
                
                Randomize
                iZufall = 0
                iZufall = Int((999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 999 generieren.
                sZufall = Str(iZufall)
                sZufall = "A" & Trim(sZufall)
                
                cTabelle = cdatei
                
                If NewTableSuchenDBKombi(cTabelle, gdBase) Then
                    HoleBestellDateiWK16a cdatei
                Else
                    cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    loeschNEW cdatei, gdBase
                    List2.RemoveItem List2.ListIndex
                    Screen.MousePointer = 0
                    Exit Sub
                
                End If
                
                Text5.Visible = True
                Label3(1).Visible = True
                Command6.Visible = True
                Check2.Visible = True
'                Text5.SetFocus
            End If
        Case Is = 1
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
            Else
                lZaehler = 0
                
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        lZaehler = lZaehler + 1
                    End If
                Next lcount
            
                If lZaehler > 1 Then
                    iRet = MsgBox("Wollen Sie die " & lZaehler & " Bestell-Dateien wirklich löschen?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                    
                        For lcount = 0 To List2.ListCount - 1
                            If List2.Selected(lcount) = True Then
                                cdatei = UCase$(Trim$(Left(List2.list(lcount), 13)))
                                
                                cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                loeschNEW cdatei, gdBase
                            End If
                        Next lcount
                    End If
                    
                    LeseInhaltWK15a
                
                Else
                
                    cdatei = UCase$(Trim$(Left(List2.list(List2.ListIndex), 13)))
                
                    iRet = MsgBox("Wollen Sie die Bestell-Datei " & vbCrLf & vbCrLf & cdatei & vbCrLf & vbCrLf & " wirklich löschen?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                        
                        cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                        gdBase.Execute cSQL, dbFailOnError
                        
                        cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                        gdBase.Execute cSQL, dbFailOnError
                        
                        loeschNEW cdatei, gdBase
                        
                        List2.RemoveItem List2.ListIndex
                    End If
                
                End If
            End If
        
        Case Is = 2
            Unload frmWK16a
        Case Is = 3 'Kundenbestellungen anzeigen
            KB "GELIEFERT", "INFORMIEREN"
            UpdateKuBestKUNDENSTATUS "INFORMIEREN", "GELIEFERT"
            Command1(3).Visible = False
        Case 4 'Exportieren
        
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
                Exit Sub
            Else
                cDatname = List2.list(List2.ListIndex)
                cDatname = Left(cDatname, 13)
                cDatname = UCase$(Trim$(cDatname))
            End If
            
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            cPfad = cPfad & "Bestell\"
            
            Screen.MousePointer = 0
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Datei speichern"
                .Filter = "Access - Dateien (*.mdb)|*.mdb"
                .FileName = cPfad & cDatname & ".mdb"
                .ShowSave
            End With
    
            cPfad = cdlopen.FileName

            If FileExists(cPfad) Then
                iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbNo Then
                    
                    Exit Sub
                Else
                    Kill cPfad
                End If
            End If
            
            Set dbBestell = CreateDatabase(cPfad, dbLangGeneral, dbVersion40)
            dbBestell.Close
            
            sSQL = "select * into " & cDatname & " in '" & cPfad & "' from " & cDatname
            gdBase.Execute sSQL, dbFailOnError
        
        Case 5 'Importieren
            
            Dim lAnzTable   As Long
        
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            cPfad = cPfad & "Bestell\"
            
            Screen.MousePointer = 0
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Datei importieren"
                .Filter = "Access - Dateien (*.mdb)|*.mdb"
                .FileName = cPfad & cDatname & ".mdb"
                .ShowOpen
            End With
    
            cPfad = cdlopen.FileName
            
            Set dbBestell = OpenDatabase(cPfad, False, False)
            dbBestell.TableDefs.Refresh
    
            lAnzTable = dbBestell.TableDefs.Count
    
            For lcount = 0 To lAnzTable - 1
                cDatname = dbBestell.TableDefs(lcount).name
                If Left(UCase(cDatname), 1) = "Q" Then
                    If NewTableSuchenDBKombi(cDatname, gdBase) Then
                        iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                        If iRet = vbNo Then
                            
                            Exit Sub
                        Else
                            loeschNEW cDatname, gdBase
                        End If
                    Else
                     
                    End If
                     
                    cPfad = gcDBPfad
                    If Right(cPfad, 1) <> "\" Then
                        cPfad = cPfad & "\"
                    End If
                     
                    TransferTab dbBestell, cPfad & "Kissdata.mdb", cDatname
                    
                    'BESTREST füllen
                    sSQL = "Delete from BESTREST where DATEINAME = '" & cDatname & ".DBF'"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    lDatum = Fix(Now)
                    cDatum = Trim$(Str$(lDatum))
            
                    sSQL = "Insert into BESTREST "
                    sSQL = sSQL & "Select LINR, "
                    sSQL = sSQL & "ARTNR, LEKPR, BESTVOR, '" & cDatname & ".DBF' as DATEINAME, "
                    sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                    sSQL = sSQL & " from " & cDatname & " where BESTVOR <> 0 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    dbBestell.Close
                    Screen.MousePointer = 0
                    neuFildatschreiben
                    Screen.MousePointer = 11
                    Exit For
                End If
            Next lcount
            LeseInhaltWK15a
            
        Case 6 'Drucken
        
            loeschNEW "PRINTQ", gdBase
            CreateTable "PRINTQ", gdBase
        
            If List2.ListCount > 0 Then
                For lcount = 0 To List2.ListCount - 1
                
                    cdatei = List2.list(lcount)
                    
                    cSQL = "Insert into PrintQ (Zeile) values ('" & cdatei & "')"
                    gdBase.Execute cSQL, dbFailOnError
                    
                Next lcount
                
                reportbildschirm "WKL029", "aWKL15ab"
            End If
        Case 7
            Screen.MousePointer = 0
            gsZSpalte = "Artnr"
            gsZSpalte1 = "BESTELLT"
            gsZSpalte2 = "GELIEFERT"
            gsZSpalte3 = "BERECHNET"
            gstab = "WEBEST"
            frmWKL36.Show 1
            'fertig
    End Select
    
    Screen.MousePointer = 0
    

err:
Exit Sub

LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub HoleBestellDateiWK16a(cdatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cTabelle    As String
    Dim rsrs        As Recordset
    Dim rsArtikel   As Recordset
    Dim rsTabelle   As Recordset
    Dim cArtNr      As String
    Dim ctmp        As String
    Dim cMitt       As String
    
    cTabelle = cdatei
    
    cAnfuLinr = Left(cTabelle, Len(cTabelle) - 1)
    cAnfuLinr = Right(cAnfuLinr, Len(cAnfuLinr) - 1)

    Set rsTabelle = gdBase.OpenRecordset(cTabelle, dbOpenTable)
    
    If Not rsTabelle.EOF Then
        rsTabelle.MoveFirst
        Do While Not rsTabelle.EOF
            If Not IsNull(rsTabelle!artnr) Then
                cArtNr = Trim(rsTabelle!artnr)
            End If
            cSQL = "Select * from Artikel where artnr = " & cArtNr
            Set rsArtikel = gdBase.OpenRecordset(cSQL)
            If Not rsArtikel.EOF Then
                rsArtikel.MoveFirst
                
                rsTabelle.Edit
                If Not IsNull(rsArtikel!KVKPR1) Then
                    rsTabelle!KVKPR1 = rsArtikel!KVKPR1
                End If
                
'                If Not IsNull(rsArtikel!vkpr) Then
'                    rsTabelle!vkpr = rsArtikel!vkpr
'                End If
                rsTabelle.Update
            End If
            
            cSQL = "Select * from Artlief where artnr = " & cArtNr
            cSQL = cSQL & " and Linr = " & cAnfuLinr
            Set rsArtikel = gdBase.OpenRecordset(cSQL)
            If Not rsArtikel.EOF Then
                rsArtikel.MoveFirst
                rsTabelle.Edit
                If Not IsNull(rsArtikel!lekpr) Then
                    rsTabelle!lekpr = rsArtikel!lekpr
                End If
                rsTabelle.Update
            End If
            
            rsArtikel.Close
        rsTabelle.MoveNext
        Loop
    End If
    rsTabelle.Close
    
    loeschNEW sZufall, gdBase
    
    cSQL = "Create Table " & sZufall
    cSQL = cSQL & " ( ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LEKPR Double"
    cSQL = cSQL & ", BESTELLT Long"
    cSQL = cSQL & ", GELIEFERT Long"
    cSQL = cSQL & ", BERECHNET Long"
    cSQL = cSQL & ", LIEFBETRAG Double"
    cSQL = cSQL & ", ZEILEN_RAB Double"
    cSQL = cSQL & ", ZEILENWERT Double"
    cSQL = cSQL & ", RECHN_RAB Double"
    cSQL = cSQL & ", RECHN_WERT Double"
    cSQL = cSQL & ", STCK_PREIS Double"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", KVKPR1 Double"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", MOPREIS Long"
    cSQL = cSQL & ", AWM Text(2)"
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index ARTNR on " & sZufall & " (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    If Option1(0).Value = True Then
        cSort = " order by MOPREIS, BEZEICH"
    ElseIf Option1(1).Value = True Then
        cSort = " order by MOPREIS, LIBESNR"
    ElseIf Option1(2).Value = True Then
        cSort = " order by MOPREIS, ARTNR"
    ElseIf Option1(3).Value = True Then
        cSort = " order by MOPREIS, LPZ"
    Else
        cSort = " order by MOPREIS, LPZ"
    End If
    
    If SpalteInTabellegefundenNEW(cTabelle, "bestvor", gdBase) Then
        cSQL = "Insert into " & sZufall
        cSQL = cSQL & " Select ARTNR, BEZEICH, LEKPR, BESTVOR as BESTELLT"
        cSQL = cSQL & ", BESTVOR as GELIEFERT, BESTVOR as BERECHNET"
        cSQL = cSQL & ", LEKPR * BESTVOR as LIEFBETRAG, 0 as ZEILEN_RAB"
        cSQL = cSQL & ", LEKPR * BESTVOR as ZEILENWERT, 0 as RECHN_RAB"
        cSQL = cSQL & ", LEKPR * BESTVOR as RECHN_WERT, LEKPR as STCK_PREIS"
        cSQL = cSQL & ", LINR, LIBESNR, KVKPR1,LPZ ,mopreis,awm "
        cSQL = cSQL & " from " & cTabelle & " where BESTVOR <> 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        'Mitteilung anzeigen
        anzeige "normal", "", label8(2)
        cSQL = "Select distinct(Mitteilung) as mitt from " & cTabelle
        Set rsTabelle = gdBase.OpenRecordset(cSQL)
    
        If Not rsTabelle.EOF Then
            If Not IsNull(rsTabelle!mitt) Then
                cMitt = Trim(rsTabelle!mitt)
                cMitt = SwapStr(cMitt, Chr(13), " ")
                cMitt = SwapStr(cMitt, Chr(10), " ")
                anzeige "normal", cMitt, label8(2)
            End If
        End If
        rsTabelle.Close
    Else
        cSQL = "Insert into " & sZufall
        cSQL = cSQL & " Select * from " & cTabelle & " where BESTELLT <> 0 "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Datendrin(sZufall, gdBase) Then
        zeigegrid cSort
        
        Frame2.Visible = True
        Frame0.Visible = True
        Frame1.Enabled = False
        
        MSFlexGrid1.Row = 2
        MSFlexGrid1.Col = SpaltennummerGELIEFERT
        
        Label0(0).Caption = "2"
        Label0(1).Caption = SpaltennummerGELIEFERT
        
        MSFlexGrid1.SetFocus
    Else
        cdatei = List2.list(List2.ListIndex)
        cdatei = Trim(Left(cdatei, 10))
        cdatei = UCase$(cdatei)
    
        cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW cdatei, gdBase
        LeseInhaltWK15a
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleBestellDateiWK16a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigegrid(corder As String)
On Error GoTo LOKAL_ERROR

    Tabcheck "WEBEST"
    FormatGridOverTablay "WEBEST"
    
    Dim j As Integer
    
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
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With
    
    FuellenMShFlex1WKL16a corder
    
    ermittlespalten16a
    
    FaerbenHGrid MSFlexGrid1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    FaerbenRedwenn0 MSFlexGrid1
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
    
Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigegrid"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub FaerbenRedwenn0(grid As MSHFlexGrid)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Dim lBestellt As Long
    Dim lGeliefert As Long
    
    With grid
        .Redraw = False
    
        For i = 2 To .Rows - 1
            .Row = i
            .Col = SpaltennummerBESTELLT
            lBestellt = Val(.Text)
            
            .Col = SpaltennummerGELIEFERT
            lGeliefert = Val(.Text)
            
            If lGeliefert < lBestellt Then
                .Col = SpaltennummerGELIEFERT
                .CellBackColor = vbRed
                
                .Col = SpaltennummerBERECHNET
                .CellBackColor = vbRed
                
            Else
                .Col = SpaltennummerGELIEFERT
                .CellBackColor = vbWhite
                
                .Col = SpaltennummerBERECHNET
                .CellBackColor = vbWhite
            End If
        Next i
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenRedwenn0"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten16a()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "LEKPR"
                SpaltennummerLEKPR = i
            Case Is = "KVKPR1"
                SpaltennummerKVKPR1 = i
            Case Is = "BESTELLT"
                SpaltennummerBESTELLT = i
            Case Is = "GELIEFERT"
                SpaltennummerGELIEFERT = i
            Case Is = "BERECHNET"
                SpaltennummerBERECHNET = i
            Case Is = "BEZEICH"
                SpaltennummerBEZEICH = i
            Case Is = "LIEFBETRAG"
                SpaltennummerLIEFBETRAG = i
            Case Is = "ZEILEN_RAB"
                SpaltennummerZEILEN_RAB = i
            Case Is = "ZEILENWERT"
                SpaltennummerZEILENWERT = i
            Case Is = "RECHN_RAB"
                SpaltennummerRECHN_RAB = i
            Case Is = "RECHN_WERT"
                SpaltennummerRECHN_WERT = i
            Case Is = "STCK_PREIS"
                SpaltennummerSTCK_PREIS = i
            Case Is = "LINR"
                SpaltennummerLINR = i
            Case Is = "LIBESNR"
                SpaltennummerLIBESNR = i
            Case Is = "LPZ"
                SpaltennummerLPZ = i
            Case Is = "MOPREIS"
                SpaltennummerMOPREIS = i
    
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten16a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKL16a(corder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    
    Set rsrs = gdBase.OpenRecordset("Select * from " & sZufall & corder)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    With MSFlexGrid1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            .Rows = lrow + 1
            .Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "LEK", "KVK", "Lieferbetrag", "Zeilenrabatt", "Zeilenwert", "Rechn.Rabatt", "Rechn.Wert", "Stückpreis"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")

                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert


                    End Select
                    
                    If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                        aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                    End If
                    
                End If
            Next i
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.5
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
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL16a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdAnfuegen_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            bAnfuegen = False
            artikel_suchen

        Case Is = 1
            fraArtAnfuegen.Visible = False
        Case Is = 2
            bAnfuegen = True
            ArtikelAnfuegenWKL15a
    End Select
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdAnfuegen_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub artikel_suchen()
    On Error GoTo LOKAL_ERROR:
    
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim sFeld       As String
    Dim rec         As Recordset
    Dim cBezeichSuche As String
    
    loeschNEW "anfue", gdBase

    sSQL = "select artikel.artnr"
    sSQL = sSQL & ", artikel.bezeich "
    sSQL = sSQL & ", artikel.ean "
    sSQL = sSQL & ", artlief.libesnr "
    sSQL = sSQL & ", artikel.kvkpr1 "
    sSQL = sSQL & ", artlief.lekpr "
    sSQL = sSQL & " into anfue from artikel inner join artlief on "
    sSQL = sSQL & " artikel.artnr = artlief.artnr "
    sSQL = sSQL & " Where artlief.linr = " & cAnfuLinr
    sSQL = sSQL & " and artikel.artnr not in(select artnr from " & gcWEdatei & ")"
    
    If Trim(Text2.Text) <> "" Then
        sSQL = sSQL & " and artikel.artnr= " & Trim(Text2.Text)
    End If
    
    If Text4.Text <> "" Then
        cBezeichSuche = Text4.Text
        cBezeichSuche = SwapStr(cBezeichSuche, " ", "*")
        sSQL = sSQL & " and bezeich like '*" & cBezeichSuche & "*'"
    End If
    
    If Text3.Text <> "" Then
        sSQL = sSQL & " and ean like '" & Text3.Text & "*'"
    End If
    
    If Text6.Text <> "" Then
        sSQL = sSQL & " and artlief.libesnr like '" & Text6.Text & "*'"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    artikel_zeigen "Bezeich"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikel_suchen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ArtikelAnfuegenWKL15a()
    On Error GoTo LOKAL_ERROR

    Dim lcount      As Long
    Dim cSQL        As String
    Dim cLBSatz     As String
    Dim iTmp        As Integer
    Dim rsrs        As Recordset
    Dim rsRsF       As Recordset

    cLBSatz = ""
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            cLBSatz = Trim$(List3.list(lcount))
            cLBSatz = Left(cLBSatz, 6)
            cLBSatz = Trim$(cLBSatz)
            cSQL = "Select * from anfue where ARTNR = " & cLBSatz
            Set rsRsF = gdBase.OpenRecordset(cSQL)

            cSQL = "Select * from  " & sZufall & "  where ARTNR = " & cLBSatz
            Set rsrs = gdBase.OpenRecordset(cSQL)

            If rsrs.EOF Then
                rsrs.AddNew
                rsrs!artnr = cLBSatz
                rsrs!BEZEICH = rsRsF!BEZEICH
                cAnfuegenBez = rsRsF!BEZEICH
                If Not IsNull(rsRsF!lekpr) Then
                    rsrs!lekpr = rsRsF!lekpr
                    dAnfuLEKPR = Format$(rsRsF!lekpr, "#####0.00")
                Else
                    rsrs!lekpr = 0
                    dAnfuLEKPR = 0
                End If
                rsrs!BESTELLT = 0
                rsrs!GELIEFERT = 0
                rsrs!BERECHNET = 0
                rsrs!LIEFBETRAG = 0
                rsrs!ZEILEN_RAB = 0
                rsrs!ZEILENWERT = 0
                rsrs!RECHN_RAB = 0
                rsrs!RECHN_WERT = 0
                rsrs!STCK_PREIS = 0
                rsrs!linr = cAnfuLinr
                rsrs!LIBESNR = rsRsF!LIBESNR
                rsrs!KVKPR1 = rsRsF!KVKPR1
                rsrs.Update
            End If
            rsrs.Close: Set rsrs = Nothing
            rsRsF.Close
            Exit For
        End If
    Next lcount
    
    zeigegrid cSort
    
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    MSFlexGrid1.Col = SpaltennummerGELIEFERT
    MSFlexGrid1.SetFocus

    fraArtAnfuegen.Visible = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelAnfuegenWKL15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub artikel_zeigen(corder As String)
    On Error GoTo LOKAL_ERROR:
    
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim sFeld       As String
    Dim rec         As Recordset
    
    List4.Clear
    List4.AddItem "ArtNr" & Space(5) & "Artikelbezeichnung" & Space(21) & "EAN" & Space(14) & "BestellNr"
    List3.Clear
    
    sSQL = " Select * from anfue order by  " & corder
    Set rec = gdBase.OpenRecordset(sSQL)
    If Not rec.EOF Then
        rec.MoveFirst
        Do While Not rec.EOF
        
           If Not IsNull(rec!artnr) Then
               sFeld = rec!artnr
           End If
           
           sFeld = sFeld & Space$(10 - Len(sFeld))
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!BEZEICH) Then
               If Len(rec!BEZEICH) > 35 Then
                   sFeld = Mid$(rec!BEZEICH, 1, 32) & "..."
               Else
                   sFeld = rec!BEZEICH
               End If
           End If
           
           sFeld = sFeld & Space$(37 - Len(sFeld))
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!EAN) Then
               sFeld = rec!EAN
           Else
               sFeld = ""
           End If
           
           sFeld = Space$(15 - Len(sFeld)) & sFeld
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!LIBESNR) Then
               sFeld = rec!LIBESNR
           End If
           
           sFeld = Space$(13 - Len(sFeld)) & sFeld
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           List3.AddItem cLBSatz
           cLBSatz = ""
           rec.MoveNext
        Loop
    End If
    rec.Close: Set rec = Nothing
    
    
    List4.Refresh
    List3.Refresh
    List3.Visible = True
    List4.Visible = True
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikel_zeigen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet    As Integer
    
    Select Case Index
        Case Is = 0
        
            If IsAktionZulaessig("Lieferung übernehmen") = False Then
                Exit Sub
            End If

            iRet = MsgBox("Die angezeigten Daten in den Artikelbestand übernehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
            
                Screen.MousePointer = 11
                
                MSFlexGrid1_SelChange
                
'                EingangDerArtikel

                If gbGescheitert = True Then
                    gbGescheitert = False
                    Frame2.Visible = True
                    Frame0.Visible = True
                    Frame1.Enabled = False
                Else
                    LeseInhaltWK15a
                    
                    Label2(1).Caption = ""
                    Label2(2).Caption = ""
                    Frame1.Visible = True
                    Frame1.Enabled = True
                    Frame2.Visible = False
                    Frame0.Visible = False
                End If
            End If

            AktionAustragen "Lieferung übernehmen"
        Case Is = 1
        
            MSFlexGrid1_SelChange
            iRet = MsgBox("Möchten Sie wirklich den Wareneingang verlassen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            
            If iRet = vbYes Then
                frame2close
            End If
        Case Is = 2
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text6.Text = ""
            List4.Clear
            List4.Visible = True
            List3.Clear
            List3.Visible = True
            
            cmdAnfuegen_Click 0
            fraArtAnfuegen.Visible = True
            Text4.SetFocus
        Case Is = 3
            Zwischenspeichern
            LeseInhaltWK15a
            frame2close
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub frame2close()
On Error GoTo LOKAL_ERROR

    Frame0.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    Frame1.Enabled = True
    
    Text5.Visible = False
    Label3(1).Visible = False
    Command6.Visible = False
    Check2.Visible = False
    
    lblUeberschrift(1).Caption = ""
    Text1.Text = ""
    
    If sZufall <> "" Then
        loeschNEW sZufall, gdBase
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2close"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Zwischenspeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim j           As Integer
    Dim cTabelle    As String
    Dim cdatei      As String
    
    cdatei = List2.list(List2.ListIndex)
    cdatei = Trim(Left(cdatei, 10))
    cdatei = UCase$(cdatei)
    
    sSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    cTabelle = "Q" & cAnfuLinr & "Z"
    
    sSQL = "Update BESTREST set DATEINAME = '" & cTabelle & "' where DATEINAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW cTabelle, gdBase
    
    sSQL = " Create Table " & cTabelle
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR Long"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LEKPR Double"
    sSQL = sSQL & ", BESTELLT Long"
    sSQL = sSQL & ", GELIEFERT Long"
    sSQL = sSQL & ", BERECHNET Long"
    sSQL = sSQL & ", LIEFBETRAG Double"
    sSQL = sSQL & ", ZEILEN_RAB Double"
    sSQL = sSQL & ", ZEILENWERT Double"
    sSQL = sSQL & ", RECHN_RAB Double"
    sSQL = sSQL & ", RECHN_WERT Double"
    sSQL = sSQL & ", STCK_PREIS Double"
    sSQL = sSQL & ", LINR Long"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", LPZ Long"
    sSQL = sSQL & ", MOPREIS Long"
    cSQL = cSQL & ", AWM Text(2)"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from TABDATUM where TABNAME like '" & cTabelle & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
    sSQL = sSQL & " ( '" & cTabelle & "','" & DateValue(Now) & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into " & cTabelle & " select * from " & sZufall & " "
    gdBase.Execute sSQL, dbFailOnError
    
    If cdatei <> cTabelle Then
        loeschNEW cdatei, gdBase
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zwischenspeichern"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim dWert As Double
    Dim cWert As String
    
    Screen.MousePointer = 11
    
    If Trim$(Text1.Text) = "" Then
        MsgBox "Bitte den Zeilenrabatt eingeben!", vbInformation, "Winkiss Hinweis:"
        Text1.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    Else
        cWert = Text1.Text
        cWert = fnMoveComma2Point$(cWert)
        dWert = Val(cWert)
    End If
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Visible = False
    
    For lrow = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lrow
        Label0(0).Caption = Trim$(Str$(lrow))
        MSFlexGrid1.Col = 0
        Label0(2).Caption = MSFlexGrid1.Text
        MSFlexGrid1.Col = SpaltennummerZEILEN_RAB
        Label0(1).Caption = SpaltennummerZEILEN_RAB
        MSFlexGrid1.Text = Format$(dWert, "#####0.00")
'        gbAender = True
'        MSFlexGrid1_SelChange
    Next lrow
    
'    MoveBestell2GridWK15a
    
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Redraw = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Select Case Index
        Case Is = 0
            cSQL = "Update " & sZufall & " Set GELIEFERT = 0, BERECHNET = 0 "
        Case Is = 1
            cSQL = "Update " & sZufall & " Set GELIEFERT = BESTELLT, BERECHNET = BESTELLT "
    End Select
    gdBase.Execute cSQL, dbFailOnError
    
    MSFlexGrid1.Redraw = False
    
    zeigegrid cSort
    
    MSFlexGrid1.Row = 2
    MSFlexGrid1.Col = SpaltennummerGELIEFERT
    MSFlexGrid1.SetFocus
    
    MSFlexGrid1.Redraw = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim llaeng      As Long
    Dim ctmp        As String
    Dim sArtnr      As String
    Dim rsArt       As Recordset
    Dim sBedingung  As String
    Dim sBez        As String
    Dim i           As Integer
    Dim lRows       As Long
    Dim bFound      As Boolean
    
    ctmp = Trim(Text5.Text)
    llaeng = Len(ctmp)
    
    If llaeng < 7 Then
        sArtnr = ctmp
    ElseIf llaeng >= 7 Then
        sBedingung = sBedingung & "where ean = '" & ctmp & "'"
        sBedingung = sBedingung & " or ean2 = '" & ctmp & "'"
        sBedingung = sBedingung & " or ean3 = '" & ctmp & "'"
        
        sSQL = "select * from artikel " & sBedingung
        Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                sArtnr = rsArt!artnr
            End If
        rsArt.Close: Set rsArt = Nothing
    End If
    
    bFound = False
    
    If sArtnr = "" Then
        Text5.Text = ""
        Text5.SetFocus
    Else
        Text5.Text = ""
        With MSFlexGrid1
            lRows = .Rows
            .Redraw = False
            For i = 2 To lRows - 1
                If .TextMatrix(i, CLng(SpaltennummerArtnr)) = sArtnr Then
                    bFound = True
                    
                    If Check2.Value = vbChecked Then
                        .Col = SpaltennummerGELIEFERT
                        .Row = i
                        .Text = Val(.Text) + 1
                        
                        .Col = SpaltennummerBERECHNET
                        .Row = i
                        .Text = Val(.Text) + 1
                    End If
                    
                    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    
                    .TopRow = i
                    .Col = SpaltennummerGELIEFERT
                    .Row = i
                    
                    Text5.Text = ""
                    
                    If Check2.Value = vbUnchecked Then
                        .SetFocus
                    Else
                        Text5.SetFocus
                    End If
                    
                    Exit For
                Else
                    .TopRow = 1
                    .Row = 1
                    Text5.Text = ""
                    Text5.SetFocus
                End If
            Next i
            .Redraw = True
        End With
    End If
    
    If bFound = False Then
        If sArtnr <> "" Then
            sBez = bezis(sArtnr)
            If sBez <> "" Then
                sSQL = "Select * from  " & sZufall
                Set rsArt = gdBase.OpenRecordset(sSQL)
    
                rsArt.AddNew
                rsArt!artnr = sArtnr
                rsArt!BEZEICH = sBez
                rsArt!lekpr = ermLEKPR(sArtnr, CLng(cAnfuLinr))
                rsArt!BESTELLT = 0
                
                If Check2.Value = vbChecked Then
                    rsArt!GELIEFERT = 1
                    rsArt!BERECHNET = 1
                Else
                    rsArt!GELIEFERT = 0
                    rsArt!BERECHNET = 0
                End If
                
                rsArt!LIEFBETRAG = 0
                rsArt!ZEILEN_RAB = 0
                rsArt!ZEILENWERT = 0
                rsArt!RECHN_RAB = 0
                rsArt!RECHN_WERT = 0
                rsArt!STCK_PREIS = 0
                rsArt!linr = cAnfuLinr
                rsArt!LIBESNR = 0
                rsArt!KVKPR1 = ermKVKPR1(sArtnr)
                    
                rsArt.Update
                rsArt.Close: Set rsArt = Nothing
                
                zeigegrid cSort
    
                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                MSFlexGrid1.Col = SpaltennummerGELIEFERT
                MSFlexGrid1.SetFocus

                
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    PositionierenWK15a
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift(0)
    
    fraArtAnfuegen.Visible = False
    Text1.Text = ""
    Text5.Text = ""
    
    gbAender = False
    gbUpdate = False
    
    If NewTableSuchenDBKombi("E15A", gdApp) Then
        voreinstellungladen
    End If
    
    neuFildatschreiben
    
    LeseInhaltWK15a
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
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
Private Sub LeseInhaltWK15a()
    On Error GoTo LOKAL_ERROR
    
    List1.Clear
    List2.Clear
    List1.AddItem "Dateiname               Bestellinformationen                Auftragswert"
    
    If Option2(0).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "tabname", Label3(3)
    ElseIf Option2(1).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "tabdate", Label3(3)
    ElseIf Option2(2).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "Liefbez", Label3(3)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseInhaltWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWK15a()
    On Error GoTo LOKAL_ERROR
    
    With Frame0
        .Top = 7800
        .Left = 0
        .Height = 1095
        .Width = 11895
    End With

    With Frame1
        .Top = 720
        .Left = 0
        .Height = 7935
        .Width = 11895
    End With
    
    With Frame2
        .Top = 720
        .Left = 0
        .Height = 6975
        .Width = 11895
    End With
    
    With fraArtAnfuegen
        .Top = 720
        .Height = 7935
        .Left = 120
        .Width = 11655
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR

    LogtoEnd Me
    voreinstellungspeichern
    
    If sZufall <> "" Then
        loeschNEW sZufall, gdBase
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Set rsrs = gdApp.OpenRecordset("E15A")
    
    If Not rsrs.EOF Then
        
        Option1(0).Value = rsrs!bo1
        Option1(1).Value = rsrs!bo2
        Option1(2).Value = rsrs!bo3
        Option1(3).Value = rsrs!bo4
        
        Option2(0).Value = rsrs!bo5
        Option2(1).Value = rsrs!bo6
        Option2(2).Value = rsrs!bo7
        
        If rsrs!bo8 = True Then
            Check2.Value = vbUnchecked
        Else
            Check2.Value = vbChecked
        End If

    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    Dim bo7 As Integer
    Dim bo8 As Integer
    
    loeschNEW "E15A", gdApp
    CreateTable "E15A", gdApp
    
    bo1 = Option1(0).Value
    bo2 = Option1(1).Value
    bo3 = Option1(2).Value
    bo4 = Option1(3).Value
    bo5 = Option2(0).Value
    bo6 = Option2(1).Value
    bo7 = Option2(2).Value
    
    If Check2.Value = vbChecked Then
        bo8 = 0
    Else
        bo8 = -1
    End If
    
    sSQL = "Insert into E15A (BO1,BO2,BO3,BO4,BO5,BO6,BO7,BO8) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & "," & bo3 & "," & bo4 & "," & bo5 & "," & bo6 & "," & bo7 & "," & bo8
    sSQL = sSQL & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_Click()
On Error GoTo LOKAL_ERROR
    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    lrow = MSFlexGrid1.Row
    
    Select Case KeyCode
        Case Is = vbKeyF2
            lrow = MSFlexGrid1.Row
            gsARTNR = MSFlexGrid1.TextMatrix(lrow, SpaltennummerArtnr)
            If gsARTNR <> "" Then
    
                frmWKL10.Show 1
                Me.Refresh
                Screen.MousePointer = 0
                MSFlexGrid1.Col = SpaltennummerGELIEFERT
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.TopRow = lrow
                MSFlexGrid1.SetFocus
            End If
            gsARTNR = ""
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_LeaveCell()
On Error GoTo LOKAL_ERROR
    
iKeypress = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub checkDiezeile(grid As MSHFlexGrid, lrow As Long)
On Error GoTo LOKAL_ERROR

Dim lBestellt As Long
Dim lGeliefert As Long

Dim lRowmerker As Long
Dim lColmerker As Long

With grid
    .Redraw = False
    lRowmerker = .Row
    lColmerker = .Col
    .Row = lrow
    .Col = SpaltennummerBESTELLT
    lBestellt = Val(.Text)
    
    .Col = SpaltennummerGELIEFERT
    lGeliefert = Val(.Text)
    
    
    
    
    
    'Farbgebung
    If lGeliefert < lBestellt Then
        .Col = SpaltennummerGELIEFERT
        .CellBackColor = vbRed
        
        .Col = SpaltennummerBERECHNET
        .CellBackColor = vbRed
    Else
        .Col = SpaltennummerGELIEFERT
        .CellBackColor = vbWhite
        
        .Col = SpaltennummerBERECHNET
        .CellBackColor = vbWhite
    End If
    
    .Col = lColmerker
    .Row = lRowmerker
    .Redraw = True
End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkDiezeile"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_SelChange()
    On Error GoTo LOKAL_ERROR

    Dim lColmerker          As Long
    Dim lRowmerker          As Long
    Dim cArtNr              As String
    Dim cartnrzuSpeichern   As String
    Dim cPreis              As String
    
    checkDiezeile MSFlexGrid1, CLng(Label0(0).Caption)
    
    If MSFlexGrid1.Row > 1 Then
    
        MSFlexGrid1.Redraw = False
        
        If gbAenderKVK Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerKVKPR1)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "KVKPR1"
                
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderKVK = False
        End If
        
        If gbAenderLEKPR Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerLEKPR)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LEKPR"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLEKPR = False
        End If
        
        If gbAenderBEZEICH Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerBEZEICH)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(Label0(0).Caption), SpaltennummerArtnr)
        
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "BEZEICH"

            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderBEZEICH = False
        End If
        
        
        
        
        MSFlexGrid1.Redraw = True
        
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFLexGrid1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    Dim cArtNr As String
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    Label0(0).Caption = lrow

    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
    
        Case Is = SpaltennummerBERECHNET
            gbAenderBERECHNET = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
    
        Case Is = SpaltennummerGELIEFERT
            gbAenderGELIEFERT = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                    
                    'auch BERECHNET füllen
                    gbAenderBERECHNET = True
                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.Col = SpaltennummerBERECHNET
                    
                    MSFlexGrid1.Text = cValid
                    
                    MSFlexGrid1.Col = SpaltennummerGELIEFERT
                    'auch BERECHNET füllen ende
                    
                    
                End If
            End If
            
            

        Case Is = SpaltennummerBESTELLT
            gbAenderBESTELLT = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If

        Case Is = SpaltennummerBEZEICH
            gbAenderBEZEICH = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerKVKPR1
            gbAenderKVK = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerLEKPR
            gbAenderLEKPR = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
     End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then  'And KeyCode <> vbKeyReturn
    
        Select Case lcol
            Case Is = SpaltennummerKVKPR1, SpaltennummerLEKPR, SpaltennummerBEZEICH, _
            SpaltennummerBESTELLT, SpaltennummerGELIEFERT, SpaltennummerBERECHNET
        
                If iKeypress = 0 And KeyCode <> vbKeyBack And KeyCode <> vbKeyF2 And KeyCode <> vbKeyReturn Then
                
'                    If Check4.Value = vbChecked Then
                        MSFlexGrid1.Row = lrow
                        MSFlexGrid1.Col = lcol
                        MSFlexGrid1.Text = ""
'                    End If

                ElseIf iKeypress > 0 And KeyCode = 46 Then

                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.Col = lcol
                    MSFlexGrid1.Text = ""

                End If
                iKeypress = iKeypress + 1
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Select Case MSFlexGrid1.Col
        Case SpaltennummerArtnr
            cSort = " Order by Artnr "
        Case SpaltennummerLEKPR
            cSort = " Order by LEKPR "
        Case SpaltennummerKVKPR1
            cSort = " Order by KVKPR1 "
        Case SpaltennummerBESTELLT
            cSort = " Order by BESTELLT "
        Case SpaltennummerGELIEFERT
            cSort = " Order by GELIEFERT "
        Case SpaltennummerBERECHNET
            cSort = " Order by BERECHNET "
        Case SpaltennummerBEZEICH
            cSort = " Order by BEZEICH "
        Case SpaltennummerLIEFBETRAG
            cSort = " Order by LIEFBETRAG "
        Case SpaltennummerZEILEN_RAB
            cSort = " Order by ZEILEN_RAB "
        Case SpaltennummerZEILENWERT
            cSort = " Order by ZEILENWERT "
        Case SpaltennummerRECHN_RAB
            cSort = " Order by RECHN_RAB "
        Case SpaltennummerRECHN_WERT
            cSort = " Order by RECHN_WERT "
        Case SpaltennummerSTCK_PREIS
            cSort = " Order by STCK_PREIS "
        Case SpaltennummerLINR
            cSort = " Order by LINR "
        Case SpaltennummerLIBESNR
            cSort = " Order by LIBESNR "
        Case SpaltennummerLPZ
            cSort = " Order by Lpz "
        Case SpaltennummerMOPREIS
            cSort = " Order by MOPREIS "
        Case Else
            cSort = " Order by MOPREIS,lpz"
    End Select
    
    If MSFlexGrid1.Row > 1 Then
        
    Else
        sortierenHGrid MSFlexGrid1
    End If
    
    If byteSortReihen = 2 Then
        cSort = cSort & " desc "
    ElseIf byteSortReihen = 1 Then
        cSort = cSort & " asc "
    Else
        cSort = cSort & " asc "
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command6_Click
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
